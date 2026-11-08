//! betteroffice（纯 Rust）校验区域重算路由。
//!
//! 1-4 校验区域重算的首选引擎：打开工作副本 → 按配置轮数 `recalculate_all`
//! → 写回 → 统计受限格与错误格 → 由调用方决定采用或回退 COM/LibreOffice。
//!
//! 配置（环境变量，缺省值见 [`BoCalcConfig::from_env`]）：
//! - `PYTOOLS_BO_CALC`：`auto`（默认，启用）/`off`（完全走 COM/LO 旧路由）
//! - `PYTOOLS_BO_CALC_MAX_VISITS`：重算预算上限；`0` 表示用库默认（10,000,000）
//! - `PYTOOLS_BO_CALC_ROUNDS`：重算轮数（默认 2）。betteroffice 的依赖图
//!   看不到 `INDIRECT` 的动态引用，跨表动态链第一轮可能读到尚未更新的值；
//!   第二轮起收敛（实测 S1F1 模板第二轮 changed=0）。
//! - `PYTOOLS_BO_CALC_STRIP_STYLE`：`off`（默认）/`on`。`on` 时在重算前把
//!   工作副本里的纯样式空单元格剥掉（`ooxml_patch::strip_style_only_cells`），
//!   避免"格式刷满全列"的簿把整列引用求值范围撑到内容范围之外。一次性生效，
//!   要么剥要么不剥。
//! - `PYTOOLS_BO_CALC_FALLBACK`：`auto`（按平台回 COM/LibreOffice）/`com`/
//!   `libreoffice`/`none`（超限即报错，绝不静默采用受限结果）。

use std::path::Path;
use std::time::Instant;

use anyhow::{bail, Context, Result};

/// why a betteroffice recalculation cannot be trusted and must fall back.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) enum BoCalcRefusal {
    BudgetExhausted { limited_cells: usize },
    UnsupportedFunctions { name_error_cells: usize },
    Unusable { reason: String },
}

impl std::fmt::Display for BoCalcRefusal {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            BoCalcRefusal::BudgetExhausted { limited_cells } => write!(
                formatter,
                "重算预算耗尽（{limited_cells} 个公式格未求值，保留旧值不可信）"
            ),
            BoCalcRefusal::UnsupportedFunctions { name_error_cells } => write!(
                formatter,
                "{name_error_cells} 个公式格求值为 #NAME?（模板使用了 betteroffice 未实现的函数）"
            ),
            BoCalcRefusal::Unusable { reason } => write!(formatter, "{reason}"),
        }
    }
}

/// what one betteroffice recalculation did to the working copy.
#[derive(Debug, Clone)]
pub(crate) struct BoCalcOutcome {
    pub application: &'static str,
    pub rounds_used: u32,
    pub changed_cells_total: usize,
    pub limited_cells: usize,
    pub error_cells: usize,
    pub name_error_cells: usize,
    pub style_cells_stripped: usize,
    pub open_elapsed: std::time::Duration,
    pub calc_elapsed: std::time::Duration,
    pub save_elapsed: std::time::Duration,
}

/// the five raw strings as read from the 全局配置 sheet. Empty strings mean
/// "not configured"; parsing happens in [`BoCalcConfig::from_parts`].
#[derive(Debug, Clone, Default)]
pub(crate) struct ValidationCalcSettings {
    engine: Option<String>,
    max_visits: Option<String>,
    rounds: Option<String>,
    strip_style: Option<String>,
    fallback: Option<String>,
}

/// read the five 校验计算* keys from the 全局配置 sheet; missing keys come
/// back as `None`.
fn read_validation_settings(config_path: &std::path::Path) -> Result<ValidationCalcSettings> {
    let load =
        |key: &str| -> Result<String> { crate::config::load_global_value(config_path, key, "") };
    Ok(ValidationCalcSettings {
        engine: Some(load("校验计算引擎")?),
        max_visits: Some(load("校验计算预算上限")?),
        rounds: Some(load("校验计算轮数")?),
        strip_style: Some(load("校验计算剥离样式空格")?),
        fallback: Some(load("校验计算回退策略")?),
    })
}

/// routing configuration, read once per call from the process environment.
#[derive(Debug, Clone)]
pub(crate) struct BoCalcConfig {
    pub enabled: bool,
    pub max_visits: Option<u64>,
    pub rounds: u32,
    pub strip_style_cells: bool,
    pub fallback_com: bool,
    pub fallback_libreoffice: bool,
}

impl BoCalcConfig {
    /// 配置优先级：config.xlsx 全局配置 > 环境变量 > 内置默认。
    ///
    /// `config_path` 传 `None`（或文件不存在）时跳过配置表，只看环境变量。
    pub(crate) fn load(config_path: Option<&std::path::Path>) -> Self {
        let global = config_path
            .filter(|path| path.is_file())
            .and_then(|path| read_validation_settings(path).ok());
        Self::from_parts(global.as_ref(), &|key| std::env::var(key).ok())
    }

    /// 纯函数入口，便于测试配置与环境变量的优先级。
    fn from_parts(
        global: Option<&ValidationCalcSettings>,
        env: &dyn Fn(&str) -> Option<String>,
    ) -> Self {
        // 每个键：config 非空值优先；否则环境变量非空值；否则 None。
        let pick = |config_value: Option<&String>, env_key: &str| -> Option<String> {
            config_value
                .map(|value| value.trim().to_owned())
                .filter(|value| !value.is_empty())
                .or_else(|| env(env_key).filter(|value| !value.trim().is_empty()))
        };
        let engine = pick(
            global.and_then(|global| global.engine.as_ref()),
            "PYTOOLS_BO_CALC",
        )
        .unwrap_or_else(|| "betteroffice".to_owned());
        // off 与 com 都表示不启用 betteroffice 路由（直接走平台引擎）。
        let enabled = !matches!(engine.trim().to_ascii_lowercase().as_str(), "off" | "com");
        let max_visits = pick(
            global.and_then(|global| global.max_visits.as_ref()),
            "PYTOOLS_BO_CALC_MAX_VISITS",
        )
        .and_then(|value| value.trim().parse::<u64>().ok())
        .filter(|value| *value > 0);
        let rounds = pick(
            global.and_then(|global| global.rounds.as_ref()),
            "PYTOOLS_BO_CALC_ROUNDS",
        )
        .and_then(|value| value.trim().parse::<u32>().ok())
        .filter(|value| (1..=5).contains(value))
        .unwrap_or(2);
        let strip_style_cells = pick(
            global.and_then(|global| global.strip_style.as_ref()),
            "PYTOOLS_BO_CALC_STRIP_STYLE",
        )
        .map(|value| {
            matches!(
                value.trim(),
                "是" | "Y" | "y" | "true" | "TRUE" | "True" | "1" | "on" | "ON"
            )
        })
        .unwrap_or(false);
        let fallback = pick(
            global.and_then(|global| global.fallback.as_ref()),
            "PYTOOLS_BO_CALC_FALLBACK",
        )
        .unwrap_or_else(|| "auto".to_owned())
        .trim()
        .to_ascii_lowercase();
        let fallback_com = fallback != "none" && fallback != "libreoffice";
        let fallback_libreoffice = fallback != "none" && fallback != "com";
        Self {
            enabled,
            max_visits,
            rounds,
            strip_style_cells,
            fallback_com,
            fallback_libreoffice,
        }
    }

    /// whether this platform can serve the fallback this config asks for.
    pub(crate) fn fallback_available(&self) -> bool {
        #[cfg(windows)]
        {
            self.fallback_com
        }
        #[cfg(target_os = "linux")]
        {
            self.fallback_libreoffice
        }
        #[cfg(not(any(windows, target_os = "linux")))]
        {
            false
        }
    }
}

/// Recalculate `source_copy` in place with betteroffice.
///
/// The path must be a disposable working copy (` .xlsx`/`.xlsm`); the caller
/// owns snapshot and candidate validation. On success the cached values in the
/// file have been rewritten and the report describes what happened; on
/// `Err(BoCalcRefusal)` the caller must fall back (the file may have been
/// rewritten — the fallback engines recalculate everything themselves, so a
/// partially updated copy is safe to hand them).
pub(crate) fn recalculate(source_copy: &Path, config: &BoCalcConfig) -> Result<BoCalcOutcome> {
    let extension = source_copy
        .extension()
        .and_then(|extension| extension.to_str())
        .map(|extension| extension.to_ascii_lowercase());
    if !matches!(extension.as_deref(), Some("xlsx") | Some("xlsm")) {
        bail!(
            "betteroffice 重算需要 .xlsx/.xlsm 工作副本（当前 {}）",
            source_copy.display()
        );
    }

    let mut style_cells_stripped = 0usize;
    if config.strip_style_cells {
        let stripped_path =
            stripped_sibling_path(source_copy).context("为纯样式单元格剥离结果分配临时路径失败")?;
        let report = ooxml_patch::strip_style_only_cells(source_copy, &stripped_path)
            .with_context(|| {
                format!(
                    "剥离工作副本的纯样式空单元格失败（副本：{}）",
                    source_copy.display()
                )
            })?;
        style_cells_stripped = report.style_only_cells_removed;
        if stripped_path.exists() {
            fs_rename(&stripped_path, source_copy)?;
        }
    }

    if let Some(max_visits) = config.max_visits {
        xlsx_calc::eval::set_max_recalculation_visits(max_visits);
    }

    let bytes = std::fs::read(source_copy)
        .with_context(|| format!("读取工作副本失败（{}）", source_copy.display()))?;
    let open_started = Instant::now();
    let mut workbook = betteroffice_xlsx::Workbook::open(&bytes)
        .map_err(|error| anyhow::anyhow!("betteroffice 打开工作副本失败：{}", error))?;
    let open_elapsed = open_started.elapsed();

    let calc_started = Instant::now();
    let options = betteroffice_xlsx::CalculationOptions {
        now_serial: Some(current_excel_serial()),
    };
    let mut changed_cells_total = 0usize;
    let mut result = workbook.recalculate_all(options);
    changed_cells_total += result.changed.len();
    for _ in 1..config.rounds {
        result = workbook.recalculate_all(options);
        changed_cells_total += result.changed.len();
    }
    let limited_cells = result.limited_cells.len();
    let budget_exhausted = result.budget_exhausted;
    let calc_elapsed = calc_started.elapsed();

    let save_started = Instant::now();
    let saved = workbook
        .save()
        .map_err(|error| anyhow::anyhow!("betteroffice 保存工作副本失败：{}", error))?;
    std::fs::write(source_copy, &saved)
        .with_context(|| format!("写回工作副本失败（{}）", source_copy.display()))?;
    let save_elapsed = save_started.elapsed();

    let (error_cells, name_error_cells) = count_error_cells(&workbook);

    if budget_exhausted {
        return Err(anyhow::anyhow!(BoCalcRefusal::BudgetExhausted { limited_cells }).into());
    }
    if name_error_cells > 0 {
        return Err(
            anyhow::anyhow!(BoCalcRefusal::UnsupportedFunctions { name_error_cells }).into(),
        );
    }
    if !result.cycle_cells.is_empty() {
        return Err(anyhow::anyhow!(BoCalcRefusal::Unusable {
            reason: format!("存在 {} 个循环引用格", result.cycle_cells.len())
        })
        .into());
    }

    Ok(BoCalcOutcome {
        application: "betteroffice (pure Rust)",
        rounds_used: config.rounds,
        changed_cells_total,
        limited_cells,
        error_cells,
        name_error_cells,
        style_cells_stripped,
        open_elapsed,
        calc_elapsed,
        save_elapsed,
    })
}

/// the local clock as an Excel 1900-system serial (days since 1899-12-30).
/// betteroffice has no clock of its own: without this, `NOW()`/`TODAY()` and
/// everything built on them evaluate to `#VALUE!`.
fn current_excel_serial() -> f64 {
    use chrono::{Datelike, Timelike};
    let now = chrono::Local::now();
    let date = chrono::NaiveDate::from_ymd_opt(now.year(), now.month(), now.day())
        .unwrap_or_else(|| chrono::NaiveDate::from_ymd_opt(1970, 1, 1).expect("epoch"));
    let epoch = chrono::NaiveDate::from_ymd_opt(1899, 12, 30).expect("excel epoch");
    let days = (date - epoch).num_days() as f64;
    let seconds = now.hour() as f64 * 3600.0
        + now.minute() as f64 * 60.0
        + now.second() as f64
        + now.nanosecond() as f64 / 1e9;
    days + seconds / 86_400.0
}

/// total formula cells evaluating to an Excel error, and how many of those are
/// `#NAME?` (the marker for functions betteroffice has not implemented).
fn count_error_cells(workbook: &betteroffice_xlsx::Workbook) -> (usize, usize) {
    use betteroffice_xlsx::CellValue;
    let mut error_cells = 0usize;
    let mut name_error_cells = 0usize;
    for sheet in &workbook.model().sheets {
        for (_cell_ref, cell) in sheet.iter_cells() {
            if cell.formula.is_none() {
                continue;
            }
            if let CellValue::Error { value } = &cell.value {
                error_cells += 1;
                if matches!(
                    value,
                    betteroffice_xlsx::ErrorValue::Name
                        | betteroffice_xlsx::ErrorValue::Ref
                        | betteroffice_xlsx::ErrorValue::Value
                        | betteroffice_xlsx::ErrorValue::Num
                ) {
                    // #NAME? is the unsupported-function marker; the rest are
                    // ordinary business errors the candidate validation sees.
                    if matches!(value, betteroffice_xlsx::ErrorValue::Name) {
                        name_error_cells += 1;
                    }
                }
            }
        }
    }
    (error_cells, name_error_cells)
}

fn ensure(condition: bool, message: impl std::fmt::Display) -> Result<()> {
    if condition {
        Ok(())
    } else {
        bail!("{message}")
    }
}

/// a sibling temp path for the stripped copy, unique per process and instant.
fn stripped_sibling_path(source_copy: &Path) -> Option<std::path::PathBuf> {
    let file_name = source_copy.file_name()?.to_str()?;
    let nonce = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .ok()?
        .as_nanos();
    let parent = source_copy.parent()?;
    Some(parent.join(format!(
        ".{file_name}.bostrip.{}.{}",
        std::process::id(),
        nonce
    )))
}

fn fs_rename(from: &Path, to: &Path) -> Result<()> {
    #[cfg(windows)]
    {
        // Windows rename fails onto an existing target; remove first. The
        // working copy is disposable, and the strip already succeeded.
        if to.exists() {
            std::fs::remove_file(to)
                .with_context(|| format!("移除旧工作副本以便替换（{}）", to.display()))?;
        }
    }
    std::fs::rename(from, to)
        .with_context(|| format!("用剥离后的副本替换工作副本（{}）", to.display()))
}

#[cfg(test)]
mod tests {
    use super::{BoCalcConfig, ValidationCalcSettings};

    fn no_env(_key: &str) -> Option<String> {
        None
    }

    fn env_of<'a>(pairs: &'a [(&'a str, &'a str)]) -> impl Fn(&str) -> Option<String> + 'a {
        move |key: &str| {
            pairs
                .iter()
                .find(|(name, _)| *name == key)
                .map(|(_, value)| value.to_string())
        }
    }

    #[test]
    fn defaults_enable_betteroffice_with_two_rounds() {
        let config = BoCalcConfig::from_parts(None, &no_env);
        assert!(config.enabled);
        assert!(config.max_visits.is_none());
        assert_eq!(config.rounds, 2);
        assert!(!config.strip_style_cells);
        assert!(config.fallback_com && config.fallback_libreoffice);
    }

    #[test]
    fn config_sheet_takes_priority_over_env() {
        let global = ValidationCalcSettings {
            engine: Some("com".into()),
            max_visits: Some("123456".into()),
            rounds: Some("3".into()),
            strip_style: Some("是".into()),
            fallback: Some("none".into()),
        };
        let env = env_of(&[
            ("PYTOOLS_BO_CALC", "betteroffice"),
            ("PYTOOLS_BO_CALC_MAX_VISITS", "999"),
            ("PYTOOLS_BO_CALC_ROUNDS", "1"),
            ("PYTOOLS_BO_CALC_STRIP_STYLE", "off"),
            ("PYTOOLS_BO_CALC_FALLBACK", "auto"),
        ]);
        let config = BoCalcConfig::from_parts(Some(&global), &env);
        assert!(!config.enabled, "config 的 com 优先于 env 的 betteroffice");
        assert_eq!(config.max_visits, Some(123456));
        assert_eq!(config.rounds, 3);
        assert!(config.strip_style_cells);
        assert!(!config.fallback_com && !config.fallback_libreoffice);
    }

    #[test]
    fn env_used_when_config_key_missing() {
        let global = ValidationCalcSettings {
            engine: None,
            max_visits: None,
            rounds: Some("".into()),
            strip_style: None,
            fallback: None,
        };
        let env = env_of(&[
            ("PYTOOLS_BO_CALC", "off"),
            ("PYTOOLS_BO_CALC_ROUNDS", "4"),
            ("PYTOOLS_BO_CALC_STRIP_STYLE", "on"),
        ]);
        let config = BoCalcConfig::from_parts(Some(&global), &env);
        assert!(!config.enabled, "env 的 off 在 config 缺失时生效");
        assert_eq!(config.rounds, 4);
        assert!(config.strip_style_cells);
    }

    #[test]
    fn builtin_defaults_survive_bad_values() {
        let global = ValidationCalcSettings {
            engine: Some("".into()),
            max_visits: Some("abc".into()),
            rounds: Some("99".into()),
            strip_style: Some("随便".into()),
            fallback: Some("".into()),
        };
        let config = BoCalcConfig::from_parts(Some(&global), &no_env);
        assert!(config.enabled, "空引擎名回落 betteroffice");
        assert_eq!(config.max_visits, None, "非法预算回落默认");
        assert_eq!(config.rounds, 2, "越界轮数回落默认 2");
        assert!(!config.strip_style_cells);
        assert!(config.fallback_com && config.fallback_libreoffice);
    }
}

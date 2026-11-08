//! 组级批量编排：并行跑 betteroffice，只把拒绝项交给平台回退（COM/LibreOffice）。
//!
//! 输入是一组**已复制好模板校验公式的工作副本**：每本对应一个独立
//! `betteroffice_xlsx::Workbook`，bo_calc 内部不共享任何可变状态，因此
//! rayon 可以安全地把整组副本并行重算。拒绝项（预算耗尽、未实现函数、
//! 循环引用、打不开等）由调用方统一交给既有平台路径——回退引擎整簿重算，
//! 不依赖 betteroffice 留下的半新半旧缓存。

use std::path::PathBuf;
use std::time::{Duration, Instant};

use rayon::prelude::*;

use crate::bo_calc::{BoCalcConfig, BoCalcOutcome};

/// 单本副本的批量结果。
#[derive(Debug)]
pub(crate) enum BatchOutcome {
    /// betteroffice 已把重算结果写回该副本。
    Succeeded(BoCalcOutcome),
    /// 不可采用，需要平台引擎整簿重算；`reason` 含具体拒绝原因，
    /// 可直接拼进 application 字段。
    Refused { reason: String },
}

/// 对同一组的全部副本并行重算，按**输入顺序**返回 `(index, outcome, elapsed)`。
///
/// 预算上限 `set_max_recalculation_visits` 是进程级全局值；本函数对所有线程
/// 写入同一份 `config` 的同一个值，因此并发写同一值是幂等的，不会互相干扰。
pub(crate) fn recalculate_parallel(
    copies: &[PathBuf],
    config: &BoCalcConfig,
) -> Vec<(usize, BatchOutcome, Duration)> {
    if let Some(max_visits) = config.max_visits {
        xlsx_calc::eval::set_max_recalculation_visits(max_visits);
    }
    copies
        .par_iter()
        .enumerate()
        .map(|(index, path)| {
            let started = Instant::now();
            let outcome = match crate::bo_calc::recalculate(path, config) {
                Ok(outcome) => BatchOutcome::Succeeded(outcome),
                Err(error) => BatchOutcome::Refused {
                    reason: format!("{error:#}"),
                },
            };
            (index, outcome, started.elapsed())
        })
        .collect()
}

#[cfg(test)]
mod tests {
    use super::{recalculate_parallel, BatchOutcome};
    use crate::bo_calc::BoCalcConfig;
    use std::path::PathBuf;

    fn temp_dir(label: &str) -> PathBuf {
        let nonce = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .expect("clock after epoch")
            .as_nanos();
        let dir = std::env::temp_dir().join(format!(
            "bo_batch_{}_{}_{}",
            std::process::id(),
            nonce,
            label
        ));
        std::fs::create_dir_all(&dir).expect("create temp dir");
        dir
    }

    /// a workbook whose single formula keeps betteroffice busy for a while.
    fn write_formula_workbook(path: &std::path::Path, rows: u32) -> anyhow::Result<()> {
        let mut workbook = rust_xlsxwriter::Workbook::new();
        let sheet = workbook.add_worksheet();
        sheet.set_name("数据")?;
        for row in 1..=rows {
            sheet.write_number(row, 0, f64::from(row))?;
        }
        let calc = workbook.add_worksheet();
        calc.set_name("算")?;
        calc.write_formula(0, 0, "=SUM(数据!A:A)")?;
        workbook.save(path)?;
        Ok(())
    }

    fn batch_config() -> BoCalcConfig {
        // 默认配置：启用 betteroffice、库默认预算、2 轮、不剥样式、允许回退。
        BoCalcConfig::from_parts(None, &|_key| None)
    }

    #[test]
    fn all_copies_succeed_and_order_is_preserved() -> anyhow::Result<()> {
        let dir = temp_dir("all_ok");
        let mut copies = Vec::new();
        for index in 0..4 {
            let path = dir.join(format!("copy{index}.xlsx"));
            write_formula_workbook(&path, 50 + index)?;
            copies.push(path);
        }

        let results = recalculate_parallel(&copies, &batch_config());
        assert_eq!(results.len(), copies.len(), "one outcome per copy");

        let mut seen: Vec<usize> = results.iter().map(|(index, _, _)| *index).collect();
        seen.sort_unstable();
        assert_eq!(
            seen,
            (0..copies.len()).collect::<Vec<_>>(),
            "indices complete"
        );

        for (index, outcome, _) in &results {
            match outcome {
                BatchOutcome::Succeeded(outcome) => {
                    assert_eq!(outcome.application, "betteroffice (pure Rust)");
                    assert_eq!(
                        outcome.limited_cells, 0,
                        "success implies every cell was evaluated"
                    );
                }
                BatchOutcome::Refused { reason } => {
                    panic!("copy {index} should have succeeded: {reason}")
                }
            }
        }
        let _ = std::fs::remove_dir_all(&dir);
        Ok(())
    }

    #[test]
    fn a_broken_copy_is_refused_without_failing_the_group() -> anyhow::Result<()> {
        let dir = temp_dir("mixed");
        let good_a = dir.join("good_a.xlsx");
        let broken = dir.join("broken.xlsx");
        let good_b = dir.join("good_b.xlsx");
        write_formula_workbook(&good_a, 40)?;
        // not a ZIP at all: betteroffice cannot open it, so the group must
        // hand this one copy to the platform engine instead of aborting.
        std::fs::write(&broken, b"not a workbook")?;
        write_formula_workbook(&good_b, 60)?;

        let copies = vec![good_a, broken, good_b];
        let results = recalculate_parallel(&copies, &batch_config());
        let by_index: Vec<(usize, &BatchOutcome)> =
            results.iter().map(|(i, o, _)| (*i, o)).collect();

        assert!(matches!(by_index[0].1, BatchOutcome::Succeeded(_)));
        assert!(
            matches!(by_index[1].1, BatchOutcome::Refused { .. }),
            "unreadable workbook must be refused, not abort the batch"
        );
        assert!(matches!(by_index[2].1, BatchOutcome::Succeeded(_)));
        let _ = std::fs::remove_dir_all(&dir);
        Ok(())
    }
}

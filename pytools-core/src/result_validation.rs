//! Optional post-save validation for workbook edits and calculation copies.

use anyhow::{ensure, Context, Result};
use std::{cell::Cell, marker::PhantomData, path::Path, rc::Rc};

pub(crate) const CONFIG_KEY: &str = "修改与计算后校验";

#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub enum ResultValidation {
    /// Internal, for diagnostics and workflows outside the modification setting.
    Existing,
    #[default]
    Sheets,
    Off,
}

impl ResultValidation {
    pub(crate) fn from_config(value: Option<&str>) -> Result<Self> {
        match value.map(str::trim).filter(|value| !value.is_empty()) {
            None => Ok(Self::default()),
            Some("是" | "软校验") => Ok(Self::Sheets),
            Some("否" | "关闭") => Ok(Self::Off),
            Some(_) => anyhow::bail!("{CONFIG_KEY} 仅接受「是」或「否」"),
        }
    }

    pub fn label(self) -> &'static str {
        match self {
            Self::Existing => "沿用原校验",
            Self::Sheets => "是（检查能否打开、预期工作表是否缺失）",
            Self::Off => "否（只检查文件能否打开）",
        }
    }
}

thread_local! {
    static SELECTED: Cell<ResultValidation> = const { Cell::new(ResultValidation::Sheets) };
}

pub struct ResultValidationScope(ResultValidation, PhantomData<Rc<()>>);

impl ResultValidationScope {
    pub fn enter(policy: ResultValidation) -> Self {
        Self(
            SELECTED.with(|selected| selected.replace(policy)),
            PhantomData,
        )
    }

    pub fn from_config(path: &Path) -> Result<Self> {
        Ok(Self::enter(crate::config::load_result_validation(path)?))
    }
}

impl Drop for ResultValidationScope {
    fn drop(&mut self) {
        SELECTED.with(|selected| selected.set(self.0));
    }
}

pub(crate) fn current() -> ResultValidation {
    SELECTED.with(Cell::get)
}

#[doc(hidden)]
pub fn check_content() -> bool {
    current() == ResultValidation::Existing
}

pub(crate) fn check_saved_sheets(expected: &[String], candidate: &Path) -> Result<()> {
    let actual = excel_reader::verify_readable_workbook(candidate)
        .with_context(|| format!("无法打开结果工作簿 {}", candidate.display()))?;
    check_sheet_roster(expected, &actual)
}

pub(crate) fn check_sheet_roster(expected: &[String], actual: &[String]) -> Result<()> {
    ensure!(!actual.is_empty(), "结果工作簿没有可读取的工作表");
    if current() == ResultValidation::Sheets {
        let missing = expected
            .iter()
            .filter(|name| !actual.contains(name))
            .cloned()
            .collect::<Vec<_>>();
        ensure!(missing.is_empty(), "结果缺少工作表：{}", missing.join("、"));
    }
    Ok(())
}

#[doc(hidden)]
pub fn check_same_sheet_result(input: &Path, candidate: &Path) -> Result<()> {
    let expected = if current() == ResultValidation::Sheets {
        excel_reader::read_sheet_names(input)?
    } else {
        Vec::new()
    };
    check_saved_sheets(&expected, candidate)
}

pub(crate) fn check_edit_plan(plan: &excel_editor::ExcelEditPlan, candidate: &Path) -> Result<()> {
    let expected = if current() == ResultValidation::Sheets {
        excel_editor::expected_sheet_names_after_edit(
            excel_reader::read_sheet_names(&plan.input)?,
            &plan.operations,
        )?
    } else {
        Vec::new()
    };
    check_saved_sheets(&expected, candidate)
}

#[derive(Debug)]
pub(crate) struct SavedResultIssue(pub String);

impl std::fmt::Display for SavedResultIssue {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(&self.0)
    }
}

impl std::error::Error for SavedResultIssue {}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn missing_or_blank_setting_defaults_to_soft_validation() -> Result<()> {
        assert_eq!(
            ResultValidation::from_config(None)?,
            ResultValidation::Sheets
        );
        assert_eq!(
            ResultValidation::from_config(Some(" "))?,
            ResultValidation::Sheets
        );
        assert_eq!(ResultValidation::default(), ResultValidation::Sheets);
        assert_eq!(current(), ResultValidation::Sheets);
        assert_eq!(
            ResultValidation::from_config(Some("是"))?,
            ResultValidation::Sheets
        );
        assert_eq!(
            ResultValidation::from_config(Some("否"))?,
            ResultValidation::Off
        );
        assert!(ResultValidation::from_config(Some("不确定")).is_err());
        Ok(())
    }

    #[test]
    fn soft_checks_only_missing_sheets_and_off_keeps_baseline_check() {
        let expected = vec!["First".to_owned(), "Second".to_owned()];
        let _soft = ResultValidationScope::enter(ResultValidation::Sheets);
        assert!(
            check_sheet_roster(&expected, &["Second".into(), "First".into(), "New".into()]).is_ok()
        );
        assert!(check_sheet_roster(&expected, &["Second".into()]).is_err());
        {
            let _off = ResultValidationScope::enter(ResultValidation::Off);
            assert!(check_sheet_roster(&expected, &["Second".into()]).is_ok());
            assert!(check_sheet_roster(&expected, &[]).is_err());
        }
        assert!(check_sheet_roster(&expected, &["Second".into()]).is_err());
    }
}

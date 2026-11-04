//! Rust implementation package for the Python statistics tools.
//!
//! Workbook access must use the workspace's shared `excel-*` and
//! `ooxml-patch` crates. This package contains no Audit-specific types or flows.

pub mod archive_ops;
pub mod batch_ops;
pub mod batch_rename;
pub mod config;
pub mod config_admin;
pub mod conversion;
pub mod conversion_engine;
pub mod date_parse;
pub mod dedup_ops;
pub mod edit_plan;
pub mod engine_diag;
pub mod excel_core_diag;
pub mod file_commit;
pub mod file_sort_prefix;
pub mod fx_header;
pub mod inplace_edit;
pub mod institution_normalize;
pub mod local_filename_map;
pub mod path_map;
pub mod print_ops;
pub mod run_logging;
pub mod quick;
pub mod region_sum;
pub mod rule_engine;
pub mod sheet_extract;
pub mod sheet_rename;
pub mod source_discovery;
pub mod static_print_engines;
pub mod submission_check;
pub mod summary_ops;
pub mod timeline_extract;
pub mod timeline_ops;
pub mod workbook_extract;

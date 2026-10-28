//! Configuration-workbook schema compatibility tests shaped after the real
//! user config.xlsx (11 sheets, authorized for read-only reference on
//! 2026-10-28). Values are sanitized: no real institution names or paths.
//!
//! Part A builds a synthetic workbook with the real sheet names, real header
//! rows and real value forms (Y/N enables, numeric booleans, composite
//! 标识列序号 "3;4", multi-area "B1:B9;BN1:BP9", absolute path style) and
//! asserts every config.rs loader accepts it.
//!
//! Part B soft-skips unless the project root config.xlsx exists; when it does
//! (local machines only; the file is gitignored) the real workbook itself is
//! loaded and only shape-level invariants are asserted.

use anyhow::Result;
use pytools_core::config::{
    load_dedup_tasks, load_extract_tasks, load_foreign_banks, load_global,
    load_institution_mapping_ordered, load_path_maps, load_print_tasks, load_timeline_rules,
};
use rust_xlsxwriter::Workbook;
use std::{
    fs,
    path::{Path, PathBuf},
};

fn unique_temp_dir(label: &str) -> PathBuf {
    let nonce = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .expect("clock after epoch")
        .as_nanos();
    let path = std::env::temp_dir().join(format!(
        "pytools-real-config-{}-{label}-{:?}-{}",
        std::process::id(),
        std::thread::current().id(),
        nonce
    ));
    fs::create_dir_all(&path).expect("create temp dir");
    path
}

fn add_sheet(
    workbook: &mut Workbook,
    name: &str,
    headers: &[&str],
    rows: &[Vec<ConfigCell>],
) {
    let sheet = workbook.add_worksheet();
    sheet.set_name(name).expect("sheet name");
    for (column, header) in headers.iter().enumerate() {
        sheet
            .write_string(0, column as u16, *header)
            .expect("header cell");
    }
    for (row_index, row) in rows.iter().enumerate() {
        for (column, cell) in row.iter().enumerate() {
            let row = (row_index + 1) as u32;
            let column = column as u16;
            match cell {
                ConfigCell::Text(value) => sheet.write_string(row, column, value.as_str()),
                ConfigCell::Number(value) => sheet.write_number(row, column, *value),
            }
            .expect("config cell");
        }
    }
}

#[derive(Clone, Debug)]
enum ConfigCell {
    Text(String),
    Number(f64),
}

fn text_row(values: &[&str]) -> Vec<ConfigCell> {
    values
        .iter()
        .map(|value| ConfigCell::Text((*value).to_owned()))
        .collect()
}

#[test]
fn loaders_accept_the_real_config_workbook_schema() -> Result<()> {
    let root = unique_temp_dir("schema");
    // load_global resolves relative paths against the config file's
    // grandparent directory (the Python PyCode convention), so the workbook
    // must sit one level below the intended base.
    let config_dir = root.join("config");
    fs::create_dir_all(&config_dir)?;
    let config_path = config_dir.join("config.xlsx");
    let absolute_source = root.join("源数据").join("示例工作簿.xlsx");
    let absolute_target = root.join("历史").join("示例汇总.xlsx");
    let mut workbook = Workbook::new();

    // 全局配置：真实形态为 键/值/备注 三列。
    add_sheet(
        &mut workbook,
        "全局配置",
        &["键", "值", "备注"],
        &[
            text_row(&["输入目录", "input", ""]),
            text_row(&["输出目录", "output", ""]),
            text_row(&["日志目录", "logs", ""]),
            text_row(&["源文件扩展名", ".xlsx;.xlsm;.csv;.et", ""]),
            text_row(&["错误策略", "continue", ""]),
        ],
    );

    // 时序提取规则：真实表为 18 列，没有 是否输出行序号 列；
    // 行头列用列字母，列表头行用逗号分隔多行，启用标记 Y/N。
    add_sheet(
        &mut workbook,
        "时序提取规则",
        &[
            "是否启用",
            "规则名称",
            "工作簿关键字",
            "工作表关键字",
            "行头列",
            "列表头行",
            "必含列头",
            "必含行头",
            "数据起始行",
            "数据结束行",
            "数据起始列",
            "数据结束列",
            "跳过关键字",
            "目标工作簿路径",
            "目标工作表",
            "启用目标写入",
            "set区域",
            "目标去重列",
        ],
        &[
            text_row(&[
                "Y",
                "测试时序规则",
                "月报;测试",
                "余额;分地区",
                "B",
                "38,39",
                "",
                "",
                "40",
                "66",
                "C",
                "",
                "",
                &absolute_target.to_string_lossy(),
                "测试时序汇总",
                "1",
                "",
                "",
            ]),
            text_row(&["N", "停用规则", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""]),
        ],
    );
    // Excel 数值 1 与文本 "1" 都应开启目标写入。
    workbook.worksheet_from_name("时序提取规则")?.write_number(1, 15, 1.0)?;

    // 路径标准化映射：真实表 10 列（含备注），可带末尾空列。
    add_sheet(
        &mut workbook,
        "路径标准化映射",
        &[
            "是否启用",
            "映射名称",
            "适用规则名",
            "工作簿关键字",
            "工作表关键字",
            "作用对象",
            "匹配方式",
            "原始路径",
            "标准路径",
            "备注",
        ],
        &[
            text_row(&[
                "Y",
                "测试列头映射",
                "测试时序规则",
                "月报;测试",
                "余额;分地区",
                "列头",
                "精确",
                "当年累计_测试指标（月）",
                "当月余额_测试指标（月）",
                "按位置",
            ]),
            text_row(&["N", "停用映射", "", "", "", "", "", "", "", ""]),
        ],
    );

    // 去重追加数据配置：真实表头没有 任务名 列，写入列名是 执行模式
    // （Rust loader 的 写入方式/覆盖方式 候选均不命中，write_mode 留空）。
    add_sheet(
        &mut workbook,
        "去重追加数据配置",
        &[
            "是否启用",
            "源数据工作簿",
            "源数据工作表",
            "标识列序号",
            "目标工作簿",
            "目标工作表",
            "执行模式",
            "备注",
        ],
        &[
            text_row(&[
                "Y",
                &absolute_source.to_string_lossy(),
                "两项收入",
                "3;4",
                &absolute_target.to_string_lossy(),
                "两项收入",
                "1",
                "",
            ]),
            text_row(&["N", "", "", "", "", "", "", ""]),
        ],
    );

    // 打印配置：真实形态 打印模式 为文本、布尔列为数字 1。
    add_sheet(
        &mut workbook,
        "打印配置",
        &[
            "是否启用",
            "打印模式",
            "源工作簿",
            "源工作表",
            "源工作表打印区域",
            "目标工作簿",
            "目标工作表",
            "FitToPagesWide",
            "FitToPagesTall",
            "打印方向",
            "水平居中",
            "垂直居中",
            "不输出批注",
            "零值不输出",
            "备注",
        ],
        &[vec![
            ConfigCell::Text("Y".to_owned()),
            ConfigCell::Text("1".to_owned()),
            ConfigCell::Text(absolute_source.to_string_lossy().into_owned()),
            ConfigCell::Text("金融统计概览".to_owned()),
            ConfigCell::Text("A1:K52".to_owned()),
            ConfigCell::Text(absolute_target.to_string_lossy().into_owned()),
            ConfigCell::Text("金融统计概览".to_owned()),
            ConfigCell::Number(1.0),
            ConfigCell::Number(1.0),
            ConfigCell::Text("纵向".to_owned()),
            ConfigCell::Number(1.0),
            ConfigCell::Number(1.0),
            ConfigCell::Number(1.0),
            ConfigCell::Number(1.0),
            ConfigCell::Text(String::new()),
        ]],
    );

    // 机构映射表：真实表头为 原始机构名称/映射后机构名称/是否为外资行。
    add_sheet(
        &mut workbook,
        "机构映射表",
        &["原始机构名称", "映射后机构名称", "是否为外资行"],
        &[
            vec![
                ConfigCell::Text("测试银行甲".to_owned()),
                ConfigCell::Text("测试银行甲".to_owned()),
                ConfigCell::Number(0.0),
            ],
            vec![
                ConfigCell::Text("测试外资银行".to_owned()),
                ConfigCell::Text("示范外资银行".to_owned()),
                ConfigCell::Number(1.0),
            ],
            vec![
                ConfigCell::Text("测试机构丙".to_owned()),
                ConfigCell::Text("丙机构".to_owned()),
                ConfigCell::Number(0.0),
            ],
        ],
    );

    // 工作表提取：真实表 9 列，没有 目标工作表 列；提取区域支持分号多区域。
    add_sheet(
        &mut workbook,
        "工作表提取",
        &[
            "是否启用",
            "工作簿关键字",
            "工作表关键字",
            "排除关键字",
            "提取区域",
            "多区域拼接方式",
            "输出文件夹",
            "文件名",
            "备注",
        ],
        &[
            text_row(&[
                "Y",
                "本外币",
                "本外币;测试;存贷款;分地区",
                "",
                "",
                "",
                "",
                "一批分地区（测试）",
                "",
            ]),
            text_row(&[
                "Y",
                "二批",
                "分地区;贷款按行业;本外币",
                "测试村镇银行",
                "B1:B9;BN1:BP9",
                "",
                "",
                "二批分地区（测试）",
                "",
            ]),
        ],
    );

    // 真实配置还包含以下四张表；它们对 config.rs loader 是惰性的，
    // 写入最小真实表头行以验证 11 表工作簿整体可加载。
    add_sheet(
        &mut workbook,
        "重命名配置",
        &[
            "简称",
            "全称",
            "占位C",
            "代码",
            "全称(代码映射)",
            "占位F",
            "键",
            "值",
            "占位I",
            "原表名",
            "新表名",
            "占位L",
            "原文件名片段",
            "新文件名片段",
            "内容重命名启用",
            "内容重命名取值规则",
            "备注",
        ],
        &[text_row(&[
            "测试",
            "测试银行",
            "",
            "5t01",
            "测试银行",
            "",
            "报表名称",
            "测试报表",
            "",
            "人民币 ",
            "1人民币",
            "",
            "",
            "",
            "N",
            "sheet1@B4",
            "",
        ])],
    );
    add_sheet(
        &mut workbook,
        "归档类型配置",
        &[
            "是否启用",
            "匹配关键词",
            "归档文件夹",
            "备注",
            "需要归档的后缀",
            "排除归档的后缀",
        ],
        &[
            text_row(&["Y", "测试：_金融", "测试：_金融", "更具体的关键词放前面", "", ""]),
            text_row(&["N", "停用类型", "停用类型", "", "", ""]),
        ],
    );
    add_sheet(
        &mut workbook,
        "提交检查配置",
        &[
            "是否启用",
            "日期关键词",
            "县区关键词",
            "机构关键词",
            "备用关键词",
            "命中次数",
            "命中的文件名",
            "备注",
        ],
        &[text_row(&[
            "N",
            "",
            "测试城区",
            "测试银行甲;测试银行甲;测试银行甲;测",
            "单家机构",
            "",
            "",
            "示例：四个非空关键词都要出现在同一文件名",
        ])],
    );
    add_sheet(
        &mut workbook,
        "文件排序配置",
        &["是否启用", "排序前缀", "匹配关键词", "备注"],
        &[text_row(&[
            "Y",
            "1",
            "测试银行甲;测试银行甲;测试银行甲;测",
            "测试排序规则",
        ])],
    );

    workbook.save(&config_path)?;

    // 全局配置：真实三列 + 四种源扩展名。
    let global = load_global(&config_path)?;
    assert_eq!(global.source_exts, [".xlsx", ".xlsm", ".csv", ".et"]);
    assert_eq!(global.error_policy, "continue");
    assert_eq!(global.input_dir, root.join("input"));
    assert_eq!(global.output_dir, root.join("output"));
    assert_eq!(global.log_dir, root.join("logs"));

    // 时序规则：Y/N 过滤、字母行头列、逗号列表头行、18 列缺省 是否输出行序号。
    let rules = load_timeline_rules(&config_path)?;
    assert_eq!(rules.len(), 1, "only the Y row loads");
    let rule = &rules[0];
    assert_eq!(rule.name, "测试时序规则");
    assert_eq!(rule.row_header_col, 2, "行头列 B resolves to column 2");
    assert_eq!(rule.col_header_rows, [38, 39], "列表头行 38,39 splits");
    assert_eq!(rule.data_row_start, 40);
    assert!(rule.target_write_enabled, "启用目标写入 1 is truthy");
    assert!(!rule.output_row_no, "real 18-col sheet has no 行序号 column");
    assert_eq!(
        rule.target_wb_path.as_deref(),
        Some(absolute_target.as_path())
    );

    // 路径映射：Y/N 过滤 + 真实 10 列表头。
    let maps = load_path_maps(&config_path)?;
    assert_eq!(maps.len(), 1);
    assert_eq!(maps[0].target, "列头");
    assert_eq!(maps[0].match_mode, "精确");
    assert_eq!(maps[0].applicable_rule, "测试时序规则");

    // 去重任务：无 任务名 列 → task_name 留空；执行模式 不在 写入方式
    // 候选内 → write_mode 留空（真实形状的既有语义）；复合键 3;4。
    let tasks = load_dedup_tasks(&config_path)?;
    assert_eq!(tasks.len(), 2, "loader keeps disabled rows with flag");
    assert!(tasks[0].enabled);
    assert!(!tasks[1].enabled);
    assert!(tasks[0].task_name.is_empty(), "real sheet has no 任务名");
    assert!(tasks[0].write_mode.is_empty(), "执行模式 is not a 写入方式 alias");
    assert_eq!(tasks[0].key_cols, [3, 4], "复合键 3;4 parses");
    assert_eq!(tasks[0].source_wb.as_deref(), Some(absolute_source.as_path()));

    // 打印任务：文本打印模式、数字布尔、打印区域。
    let prints = load_print_tasks(&config_path)?;
    assert_eq!(prints.len(), 1);
    assert_eq!(prints[0].mode, 1);
    assert_eq!(prints[0].source_ranges, "A1:K52");
    assert!(prints[0].center_h && prints[0].center_v);
    assert!(prints[0].drop_comments && prints[0].hide_zero);
    assert_eq!(prints[0].fit_wide, 1);

    // 机构映射：真实表头 + 0/1 数字外资标记。
    let foreign = load_foreign_banks(&config_path)?;
    assert_eq!(foreign.len(), 1);
    assert!(foreign.contains("示范外资银行"));
    let ordered = load_institution_mapping_ordered(&config_path)?;
    assert_eq!(ordered.len(), 3);
    assert_eq!(ordered[0], ("测试银行甲".to_owned(), "测试银行甲".to_owned()));

    // 工作表提取：Y 行加载、多区域分号保留、无 目标工作表 列 → 留空。
    let extracts = load_extract_tasks(&config_path)?;
    assert_eq!(extracts.len(), 2);
    assert_eq!(extracts[1].area_spec, "B1:B9;BN1:BP9");
    assert_eq!(extracts[1].exclude_keywords, "测试村镇银行");
    assert!(extracts[0].target_sheet.is_empty());
    assert_eq!(extracts[0].output_name, "一批分地区（测试）");

    fs::remove_dir_all(&root).ok();
    Ok(())
}

/// Soft-skipping test over the real project-root config.xlsx when present
/// (local machine only; the file is gitignored and never committed).
#[test]
fn real_project_config_loads_when_present() -> Result<()> {
    let config_path = Path::new(env!("CARGO_MANIFEST_DIR")).join("../config.xlsx");
    if !config_path.is_file() {
        eprintln!("config.xlsx not present; skipping real-config check");
        return Ok(());
    }

    let global = load_global(&config_path)?;
    assert!(
        global.source_exts.iter().any(|ext| ext == ".xlsx"),
        "real config must accept .xlsx sources"
    );
    assert!(!global.output_dir.as_os_str().is_empty());

    let rules = load_timeline_rules(&config_path)?;
    for rule in &rules {
        assert!(!rule.col_header_rows.is_empty());
        assert!(rule.row_header_col >= 1);
    }

    for task in load_dedup_tasks(&config_path)? {
        assert!(
            task.key_cols.iter().all(|column| *column >= 1),
            "标识列序号 must be one-based: row {}",
            task.row_no
        );
    }

    for task in load_print_tasks(&config_path)? {
        assert!(
            (1..=3).contains(&task.mode),
            "打印模式 out of range: row {}",
            task.row_no
        );
    }

    for task in load_extract_tasks(&config_path)? {
        assert!(!task.output_name.is_empty(), "文件名 must not be empty");
    }

    let mapping = load_institution_mapping_ordered(&config_path)?;
    assert!(
        !mapping.is_empty(),
        "real 机构映射表 is expected to contain rows"
    );
    Ok(())
}

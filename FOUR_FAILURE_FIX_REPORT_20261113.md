# 归档三项安全缺口与旧打印失败定点修复

main / fa4ba947a64300bbcac1134529848e34565b5925。此前5个已跟踪修改以起始SHA256核对全部保持不变；此前general_boundary_p0_p1.rs、test_general_boundaries.py测试文件哈希也未变，3项归档失败断言保留。没有修改Linux引擎路由、功能5或菜单；没有全量测试、Release、打包、安装、提交、push或dist更新。

## 归档三项修复

1. 预览/直接复制规划时记录源内容SHA256；计划XLSX保留原8列并追加“源文件SHA256”，JSON保留旧字段并增加source_sha256。执行前校验快照，复制后验证源及目标内容仍等于快照，失败不计成功并清理本次新建副本。读取快照失败记该文件失败，不冒充成功。
2. 在分配目标后缀前比较规范化路径，解析已存在祖先、符号路径、点路径及Windows大小写，源目标相同明确失败，不生成自复制后缀。正常不同文件同名仍使用原来的无覆盖后缀策略。
3. 递归收集源时排除目标子树及报告输出子树，包括它们的规范化别名；不把之前输出再次作为本次归档源。原有日期/类型分类和非目标文件排序不变。

兼容性影响：旧计划没有源SHA256时不能证明源版本，因此按文件失败并提示重新生成预览，不用执行时哈希伪装成规划快照。旧列位置未改，正常新计划和直接归档经过现有菜单及parity回归。SHA256约束的是文件内容，不承诺对只改变时间戳的文件拒绝。

修改：pytools-core/src/archive_ops.rs。补充内部回归验证：错误快照的复制失败清理、旧计划缺快照零复制。原有walk_files测试仅随排除参数更新调用，断言未削弱。

## 旧打印路径失败原因及修复

仍为显式print-config-plan + UMYA capability旧路径，没有切换3-1/3-2默认引擎或启用旧库入口。第一次错误Printed!C2只是首个被校验器发现的问题：

- ImportRange有hide_zero，UMYA正确把数值零/缓存零写成空白，共享校验器却仍要求源数值0。校验现传递该标志，仅对真正数值0要求空白；文本“0”、非零、布尔及其他值不会因此获豁免。
- 打印计划先把源Sheet改名为临时种子，再创建输出并删除种子。样式校验原来只按源旧名判断删除；现按操作顺序跟踪身份。元数据校验也根据计划明确删除的身份处理，保留Sheet仍必须存在且属性相同。
- 删除源种子行列尺寸需显式、按Sheet限定的许可。旧打印计划现仅为将删除的种子声明RowAndColumnDimensions许可，不移除全局保护；输出尺寸继续由ImportRange/设置操作及原用例精确验证。
- 关系校验原来把已明确删除的源Sheet自己的.rels部件视为意外丢失。现仅免除该已删除Sheet自有关系部件；包级、工作簿级和保留Sheet关系继续受保护。

修改：crates/excel-editor/src/openpyxl_backend.rs、pytools-core/src/print_ops.rs。print_config_new_target.rs仅增加种子尺寸许可断言，原失败用例的数值/文本0、公式缓存、样式、尺寸、页面设置、页脚及源/配置不变断言均保留。

新增校验器负向检查：隐藏数值0必须真为空白，关闭隐藏时仍须保留0，文本0/非零缺失必须失败；删除种子不影响保留Sheet的存在/属性/关系检查；包级关系缺失仍失败。没有通过全局关闭Styles/PageSettings/Relationships保护来放行。

## 最终相关测试结果

所有直接Cargo命令设置CARGO_TARGET_DIR=F:/RustBuild/pytools、CARGO_BUILD_JOBS=1、CARGO_PROFILE_TEST_DEBUG=0。Worker环境显式设置到随包隔离Python和仓库Worker，未部署到dist。

- cargo test --offline -p pytools-core --test general_boundary_p0_p1 --test archive_parity --test archive_menu_cli --test print_config_new_target --test print_menu_cli --test dedup_menu_cli --test file_commit --no-fail-fast -- --include-ignored --nocapture --test-threads=1：48通过、0失败、退出0。原4项失败均已转绿。日志four-fixes-final-20261113.log/exit。
- cargo test --offline -p pytools-core --lib archive_ops::tests --no-fail-fast：5通过。
- excel-editor库级分别过滤range_copy_hide_zero、style_validation_tracks、metadata_follows_planned_sheet_identity、deleted_sheet_relationship_exemption、dimensions：合计9通过，包含5项既有维度保护回归。
- 共62项Rust相关检查通过；集成中的一个平台入口另实际执行8项Python/真实CLI检查，全部通过，不将其重复计入Rust总数。合成XLS由本机Excel创建以验证拒绝，不代表XLS追加恢复或WPS通过。
- 修改的4个Rust文件rustfmt --check及git diff --check通过。没有全量运行。

中间失败记录保留：旧打印测试依次暴露零值、已删除种子身份、显式尺寸许可及关系部件的问题；一次新增单元误用了不支持Clone的SheetDump，已改为逐项移除/还原测试数据后通过；一次重建与本任务尚在运行的CLI测试占用同一调试EXE导致拒绝访问，不视为通过或业务断言结论。最终源码固定后串行完成上述门禁。

## 证据与未验证项

test-runs/four-fixes-verification-20261113.json记录HEAD、起始修改未变、原边界测试未变、最终源码哈希与各项结果。four-fixes-final日志为最终集成证据，其他早期失败日志是历史过程，不能代替最终状态。此前GENERAL_BOUNDARY_TEST_REPORT_20261113.md保留为修复前红灯证据，本报告给出修复后结论。

NOT_TESTED：真实业务样本、WPS独立环境、FAT/exFAT、共享盘/UNC断连和服务器端状态、麒麟/Linux执行、实际DEB安装发布、硬链接文件身份等更广的路径别名，以及完整ACL权限矩阵。Windows合成文件锁定/只读保护通过不能替代这些平台验收；没有宣称所有并发文件系统时序已验证。

人工最小复核：用新调试程序在隔离目录预览后更改源内容，执行应失败且目标无副本；把旧无SHA计划重新生成后再执行；选择含目标子目录的源目录，目标旧文件不应入计划；构造同源目标计划应失败不出后缀副本；旧打印显式路径与正常菜单分别用合成样本检查，不用真实业务原件。

dist/RustTools.exe仍为2026-11-11 21:12:29、28,871,168字节，未更新；旧EXE不能验证本轮保护。XLS追加仍为已授权的安全拒绝，未恢复。Git HEAD不变、暂存区为空，没有提交推送合并或操作其他工作树。

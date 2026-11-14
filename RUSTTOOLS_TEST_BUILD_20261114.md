# 2026-11-14 main 测试包与验证记录

本次交付为准备人工测试的包，不代表麒麟、WPS 或业务原件人工验收通过。
构建基线：`9f53c56ef041f7238855bfc36dbab9001c14b082` 加本次提交；源文件指纹见包内/旁边的 `BUILD_INFO.json`。
源码指纹：`c651c51cb49d3ad7990cf87904cd1a6b53dda1d7e83f42e7fa63d1d8e716a96e`。

## 本次修复

- 5-1 XLSX 的 Print_Area 校验与命名引用规范化；保持 Openpyxl 和 UMYA 各自已有的行变动口径。
- RustNative/Auto 追加时更新并验证工作表尺寸；UMYA 区分空命名容器与真实命名区域，保留关键结构。
- 4-5 Linux XLSX 使用隔离 Openpyxl；两个 Worker 修复链式二次替换及大小写引用，Rust 回读校验同步。
- 保留历史基线，补当前安全前缀契约及旧打印路径回归；不恢复 2-3 既有 XLS 追加。
- Windows 打包拒绝覆盖已有暂存目录并校验清理路径；两个包验证器按真实 definedName 元素生成测试清单，未放宽产品校验。

## 验证

- 同一 Rust/Worker 源码完整 workspace：Windows 646/0/21，WSL 638/0/21，Docker 638/0/21（通过/失败/ignored；每平台 74 组）。
- 三平台 Python CLI discovery 为 28 项；Worker Python 各 29 通过；额外 Windows 门控 11、两套 Linux 各 7 通过。
- 出包后重新验证默认包内运行时：启动、4-5 链式/大小写引用、命名区域、公式缓存、重复执行字节不变、三条追加路由和 BIFF 追加安全拒绝。
- Windows ZIP 的修改/输出回读/输入保护校验通过；Linux DEB 在 Rocky 8/glibc 2.28 私有解包目录验证通过；附加打包安全测试 5 通过。
- 首次 ZIP 验证因旧脚本错误声明空 definedNames 容器失败，原包/日志保留在本地证据目录；修正两个验证脚本后复核通过。

## 本地产物

- Windows：`dist/RustTools-windows-x64-TEST-main-20261114-095927.zip`，SHA256 `ffe7aca3f6fe5b34c9b7a92fec92ceb58424436ccf8aa131004e45ea3287d04e`。
- Linux：`dist/rusttools_0.1.0-20261114-test-main_amd64.deb`，SHA256 `716be33389801a36e225242776487e754a2bc35ee21f9067fa504970d182557c`，最高 GLIBC 符号 `2.28`。
- 直接启动：`dist/RustTools.exe`，SHA256 `cddd5f4c8c76e9742879b0302540483bbb0811638a6c5aa353ece3a368258254`。

`dist`、压缩包和本地原始日志按仓库规则不进入 Git。
根目录 `config.xlsx` 与 `dist/config/config.xlsx` 内容不同，两份均原样保留。
压缩包/DEB 采用根目录配置的副本；直接运行 dist 使用 dist 中原配置。Linux 测试前请按目标机调整包内配置目录。
本轮只提交源码、合成测试和不含业务内容的报告，不提交配置、业务文件或构建产物。

## 未验证与验收范围

NOT_TESTED：麒麟 V10 实机、DEB 安装/升级/卸载、干净 Windows、FAT/exFAT/共享盘、WPS 专项重新验收。
没有对宿主执行安装、卸载或改写已有 Linux 安装；DEB 仅构建、解包及容器验证。
Windows 专用 Worker 复杂动态公式范围本轮未扩展验收；此前 WPS ReadOnly 可能修改 BIFF 字节的发现未宣称修复。
请使用合成或备份数据人工验收 2-3 XLSX、3-1 mode 1/2/3、4-3 前缀与冲突、4-5 改名及 5-1 XLSX。
5-8 没有恢复，其他 Linux 引擎路由未调整。

机器可读记录：[20261114_test_build.json](docs/validation/20261114_test_build.json)。
原始失败、终端捕获、打包和保护记录本地保留于 `test-runs/package-push-20261114/`，未打包进提交。

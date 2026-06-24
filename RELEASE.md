# Release Notes

## INAND Benchmark Tool

### 交付物

- GUI 可执行文件：`dist/INANDBenchmark.exe`
- CLI/源码入口：`uv run python main.py`
- GUI/源码入口：`uv run python main.py --gui`
- 一键流程脚本：`scripts/run_all_to_ppt.ps1`
- PPT 粘贴脚本：`scripts/paste_excel_to_ppt.ps1`
- Spotfire 导出校验脚本：`scripts/validate_spotfire_export.py`
- 周二联网后自动运行注册脚本：`scripts/register_tuesday_online_task.ps1`

### Team 使用方式

#### 常用命令

```powershell
# 打开 GUI
dist\INANDBenchmark.exe

# 一键生成 Excel + 更新 PPT + 同步 output 到公共盘
powershell -ExecutionPolicy Bypass -File "scripts/run_all_to_ppt.ps1"

# 注册每周二后台自动任务
powershell -ExecutionPolicy Bypass -File "scripts/register_tuesday_online_task.ps1"

# 查看后台任务状态
powershell -ExecutionPolicy Bypass -File "scripts/manage_tuesday_task.ps1" -Action status

# 查看后台任务日志
powershell -ExecutionPolicy Bypass -File "scripts/manage_tuesday_task.ps1" -Action logs

# 暂停/恢复后台任务
powershell -ExecutionPolicy Bypass -File "scripts/manage_tuesday_task.ps1" -Action disable
powershell -ExecutionPolicy Bypass -File "scripts/manage_tuesday_task.ps1" -Action enable

# 删除后台任务
powershell -ExecutionPolicy Bypass -File "scripts/manage_tuesday_task.ps1" -Action unregister
```

#### GUI

双击：

```text
dist/INANDBenchmark.exe
```

GUI 内部提供：

- 生成 Excel
- 更新 PPT
- 一键 Excel + PPT
- 最近一次运行结果

#### 命令行

```powershell
powershell -ExecutionPolicy Bypass -File "scripts/run_all_to_ppt.ps1"
```

默认数据源：

```text
\\cvpfilip03\SDSS_MFG_Data\ENG_Data\Tempfile\ADolph\Spotfire_file\PN_TEST_STEPID_YIELD.csv
```

PPT 输出会按本次报告所属的上一周 `prev_week` 自动命名，例如当前周为 `2026FW52` 时：

```text
output\SDSS INAND YIELD WW51_2026_benchmark.pptx
```

PPT 内部所有 `Wxx'yy` 文字会同步替换为报告周别，例如 `W51'26`。流程结束后，本地 `output` 目录会同步复制到：

```text
\\cvpfilip03\SDSS_MFG_Data\ENG_Data\Tempfile\ADolph\Spotfire_file\output
```

#### 每周二自动运行

```powershell
powershell -ExecutionPolicy Bypass -File "scripts/register_tuesday_online_task.ps1"
```

任务会在当前用户登录时和每周二定时触发，只有周二且共享 CSV 可访问时才运行；每周成功一次后自动跳过后续触发。

### 发布前验证

```powershell
uv run pytest -q
uv run python scripts/validate_spotfire_export.py --data "PN_TEST_STEPID_YIELD.txt" --max-age-hours 9999 --min-rows 100
powershell -ExecutionPolicy Bypass -File "scripts/build_windows_exe.ps1"
```

### 注意事项

- Spotfire 导出文件同名时优先使用 CSV。
- 默认读取 Spotfire 服务器共享路径下的 `PN_TEST_STEPID_YIELD.csv`。
- GUI exe 仍依赖当前目录下的业务文件（`Rule_list.xlsx`、`bachmark SRC.xlsx`、PPT 模板、Spotfire 导出数据）。
- 参考文件不强制放到服务器；Team 多人共用时建议把 `Rule_list.xlsx`、`bachmark SRC.xlsx` 和 PPT 模板也放到共享目录并统一引用。
- PPT 更新功能依赖本机安装 Excel 和 PowerPoint（COM 自动化）。
- 若要发布到 Git remote，需要先在项目目录初始化 Git 并配置远程仓库。

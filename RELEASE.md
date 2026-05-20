# Release Notes

## INAND Benchmark Tool

### 交付物

- GUI 可执行文件：`dist/INANDBenchmark.exe`
- CLI/源码入口：`uv run python main.py`
- GUI/源码入口：`uv run python main.py --gui`
- 一键流程脚本：`scripts/run_all_to_ppt.ps1`
- PPT 粘贴脚本：`scripts/paste_excel_to_ppt.ps1`
- Spotfire 导出校验脚本：`scripts/validate_spotfire_export.py`

### Team 使用方式

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

### 发布前验证

```powershell
uv run pytest -q
uv run python scripts/validate_spotfire_export.py --data "PN_TEST_STEPID_YIELD.txt" --max-age-hours 9999 --min-rows 100
powershell -ExecutionPolicy Bypass -File "scripts/build_windows_exe.ps1"
```

### 注意事项

- Spotfire 导出文件同名时优先使用 CSV。
- GUI exe 仍依赖当前目录下的业务文件（`Rule_list.xlsx`、`bachmark SRC.xlsx`、PPT 模板、Spotfire 导出数据）。
- PPT 更新功能依赖本机安装 Excel 和 PowerPoint（COM 自动化）。
- 若要发布到 Git remote，需要先在项目目录初始化 Git 并配置远程仓库。

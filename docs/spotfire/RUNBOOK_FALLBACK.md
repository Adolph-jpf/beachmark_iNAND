# Spotfire 导出失败回退 Runbook

## 触发条件

- 定时任务失败（收到告警）
- 导出文件缺失或字段不完整
- 导出延迟超过 SLA

## 快速检查（5 分钟内）

1. 检查共享目录是否生成新文件。
2. 检查文件更新时间是否在 SLA 窗口内。
3. 运行本地校验命令：

```powershell
uv run python scripts/validate_spotfire_export.py --data "PN_TEST_STEPID_YIELD.txt"
```

## 回退路径

### A. 管理员手动触发 Job（首选）

- 让管理员在 Spotfire Job 控制台手动执行一次导出任务。
- 成功后通知你继续运行 benchmark 一键流程。

### B. 业务侧临时手工导出（兜底）

1. 在 Spotfire App 打开导出 Analysis。
2. 手工导出 CSV 到约定目录。
3. 文件命名遵循导出契约（或覆盖 `PN_TEST_STEPID_YIELD.csv`）。
4. 运行一键流程：

```powershell
powershell -ExecutionPolicy Bypass -File "scripts/run_all_to_ppt.ps1"
```

## 恢复后复盘

- 记录失败时间、根因、修复动作、恢复时间
- 更新管理员任务参数（重试、并发、资源池）
- 如有字段变更，同步更新导出契约与下游代码

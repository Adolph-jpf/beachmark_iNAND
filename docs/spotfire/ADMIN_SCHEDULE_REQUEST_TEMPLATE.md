# Spotfire 管理员任务单模板（定期导出）

## 基本信息

- 申请人：
- 业务系统：INAND Benchmark
- 任务名称：`INAND_SRC_Export_Schedule`
- 优先级：High

## 调度需求

- 调度平台：Spotfire Automation Services / Library Job
- 分析文件路径：`<Library path to INAND_SRC_Export>`
- 执行频率：每周二、周四 09:30（可调整）
- 时区：Asia/Shanghai
- 失败重试：2 次，间隔 10 分钟

## 导出配置

- 导出格式：CSV
- 编码：UTF-8-sig（或 UTF-16，与历史保持一致）
- 输出目录：`\\<share>\spotfire_exports\inand\`
- 文件命名：`PN_TEST_STEPID_YIELD_yyyyMMdd_HHmm.csv`
- latest 别名（可选）：`PN_TEST_STEPID_YIELD.csv`

## 告警与通知

- 成功通知：可选（建议日报汇总）
- 失败通知：必须
- 通知渠道：邮件 / Teams
- 收件人：`<owner>`, `<backup>`

## 验收标准

- 连续 2 周成功率 >= 99%
- 单次导出完成时间 <= 预定时间 + 15 分钟
- 下游 benchmark 可直接消费，无人工改名/改格式

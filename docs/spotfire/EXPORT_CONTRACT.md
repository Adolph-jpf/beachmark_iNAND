# Spotfire SRC 导出契约

本文档定义 Spotfire 定期导出的**稳定接口**，用于被 benchmark 流程直接消费。

## 1. 目标文件

- 文件用途：作为 benchmark 流程输入数据源
- 推荐格式：`CSV`（同名 `CSV` 优先于 `TXT`）
- 推荐编码：`UTF-16` 或 `UTF-8-sig`（现有加载器均兼容）
- 分隔符：
  - `.csv` 使用 `,`
  - `.txt` 使用 `\t`

## 2. 文件命名

- 建议：`PN_TEST_STEPID_YIELD_yyyyMMdd_HHmm.csv`
- 如果希望无缝复用当前自动流程，可同时产出（或软链接）：
  - `PN_TEST_STEPID_YIELD.csv`

## 3. 投递目录

- 建议投递到共享目录（示例）：
  - `\\<team-share>\spotfire_exports\inand\`
- benchmark 运行机应具备只读权限。

## 4. 刷新与导出频率（建议）

- 常规：每周 2 次（例如 周二 / 周四 09:30）
- 重试：失败后自动重试 2 次，间隔 10 分钟
- 告警：失败邮件/IM 通知业务 owner + 备份人

## 5. 过滤条件

- 默认导出 Spotfire 现成表全量数据（不在导出侧做额外业务裁剪）
- 若必须裁剪，需在变更单中明确，并保证下游字段完整性。

## 6. 必选字段（最低合同）

以下列必须存在（大小写敏感）：

- `PYDATE`
- `WEEK`
- `STARTTIME`
- `ENDTIME`
- `FLAG`
- `FAMILY`
- `TECH`
- `DIEQTY`
- `MTECH`
- `DEVICE`
- `STEPID_YIELD`
- `LT_INQTY/Volume` 或 `LT_INQTY`（两者之一）
- `LT_OUTQTY`
- `LT_YIELD`

说明：
- 当前程序会自动将 `LT_INQTY` 归一为内部列 `LT_INQTY/Volume`。

## 7. 质量门槛（导出后校验）

- 文件存在，更新时间在 SLA 窗口内（默认不超过 36 小时）
- 行数 > 最小阈值（建议 1000，开发环境可降到 100）
- 必选字段齐全

## 8. 兼容策略（已实现）

- 同名数据源优先级：`PN_TEST_STEPID_YIELD.csv` > `PN_TEST_STEPID_YIELD.txt`
- 编码自动探测与回退
- `%` 字符串自动转 0~1 小数（如 `99.5% -> 0.995`）

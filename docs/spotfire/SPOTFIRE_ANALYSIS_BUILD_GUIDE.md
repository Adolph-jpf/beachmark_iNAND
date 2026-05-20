# Spotfire 可调度导出 Analysis 构建指南

适用场景：你可登录 Spotfire App，但无 Server 管理权限；需由管理员配置调度。

## 目标

创建一个**专用导出 Analysis**，供管理员定时任务调用，避免业务页面改版影响导出。

## 步骤

1. 在 Spotfire Library 新建分析文件（建议命名：`INAND_SRC_Export`）。
2. 仅保留导出所需 Data Table（避免混入展示层可视化逻辑）。
3. 固定字段输出顺序（按 [`EXPORT_CONTRACT.md`](./EXPORT_CONTRACT.md)）。
4. 若使用 Python Data Function：
   - 仅做字段归一、格式化（例如列名映射、时间字段标准化）
   - 不做调度逻辑
5. 保存到可被 Automation Services / Job 读取的位置。

## 推荐参数

- Export format: CSV
- Encoding: UTF-8-sig（或沿用 UTF-16，保持与历史一致）
- Include headers: Yes
- Overwrite policy: 生成带时间戳版本 + 可选 latest 别名

## 变更管理

- 任何字段新增/删除/重命名，必须同步更新导出契约并通知下游。
- 建议在 Analysis 描述中写明：
  - owner
  - 联系方式
  - 最近修改时间

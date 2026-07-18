# 表格数据导出器 · tabular-exporter

> 一行配置把任意 JSON 数组变成漂亮的 Excel 或 CSV——自定义列、表头加粗蓝底、冻结首行、斑马纹，中文 CSV 自带 BOM 不乱码。

把数组 / JSON 导出成带表头冻结、斑马纹、布尔值转「是/否」的漂亮 Excel / CSV。零依赖（xlsx 由纯 JS 生成，CSV 自带 UTF-8 BOM），无需 `npm install` 任何包。

## 适用场景
- 运营报表、数据看板一键导出 Excel
- 数据采集 / 接口返回结果交付为可打开的表格
- 后台「导出」功能、数据迁移、日志归档

## 作为 AI 技能使用
本仓库是一个 AI Agent Skill（兼容 Claude / CodeBuddy / 类 Claude 技能的 Agent）。将 `SKILL.md` 放入 Agent 的 skills 目录即可启用；`scripts/`、`references/`、`assets/` 为配套资源，开箱即用。

## 目录结构
- `SKILL.md`：技能规范与触发词（入口）
- `scripts/`：可执行逻辑
- `references/`：参考实现

## 许可
MIT — 可自由用于商业与个人项目。

---
由教备神器自动发布。欢迎提 PR / Issue。
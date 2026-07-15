---
name: "记录数据一键导出 Excel/CSV（带样式）"
display_name: 表格数据导出器
description: "把数组/JSON 导出成带表头冻结、斑马纹、是/否的漂亮 Excel/CSV。适合运营报表、数据导出、日志归档。"
market_desc: 一行配置把任意 JSON 数组变成漂亮的 Excel 或 CSV——自定义列、表头加粗蓝底、冻结首行、斑马纹，中文 CSV 自带 BOM 不乱码。做数据报表、采集结果导出、后台导出时直接复用，零依赖免安装。
version: 1.0.1
---

# tabular-exporter —— 记录数组导出为表格

> 把任意「数组 of 对象」变成带样式的 Excel 或带 BOM 的 CSV。
> **零依赖**：xlsx 由纯 JS 生成（ZIP STORE 法，内置 CRC32），CSV 自带 UTF-8 BOM。无需 `npm install` 任何包。

## 适用场景

- 把采集 / 接口返回的结构化数据导出成可交付的 Excel / CSV
- 后台「导出」功能、数据报表、数据迁移
- 需要中文不乱码、带表头样式、冻结首行

## 提供的函数

| 函数 | 说明 |
|------|------|
| `exportToExcel(rows, columns, opts?)` | 生成带样式 xlsx（Buffer），可选落盘 `filePath` |
| `exportToCsv(rows, columns, opts?)` | 生成带 BOM 的 CSV 字符串 |
| `exportTable(rows, columns, opts?)` | 统一入口，`opts.format` = `'xlsx' \| 'csv'` |
| `columnsFromKeys(keys, widths?)` | 由 key 列表快速生成列配置 |

`columns` 形如 `[{ key, header?, width? }]`，`rows` 为对象数组。

样式特性：
- 表头：加粗 + 白字 + 蓝底（`4472C4`）
- 冻结首行
- 斑马纹（偶数行浅灰底）
- 布尔值 → `是/否`；数组 → 换行；对象 → JSON
- xlsx 与 Excel / WPS / LibreOffice 兼容（标准 OOXML）

## 用法

```js
import { exportTable, columnsFromKeys } from './scripts/exporter.mjs';

const rows = [
  { title: '无线鼠标', price: 99.5, inStock: true },
  { title: '机械键盘', price: 299, inStock: false },
];
const columns = [
  { key: 'title', header: '标题', width: 40 },
  { key: 'price', header: '价格', width: 12 },
  { key: 'inStock', header: '有货', width: 8 },
];

// 导出 xlsx（直接落盘）
exportTable(rows, columns, { format: 'xlsx', filePath: 'out.xlsx', sheetName: '商品' });

// 导出 csv（返回字符串）
const csv = exportTable(rows, columns, { format: 'csv' });
```

## 自测

```bash
node scripts/exporter.mjs --selftest
```

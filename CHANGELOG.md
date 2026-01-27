# 更新日志

## [1.0.0] - 2026-01-26

### 新增

- 初始版本发布
- Excel 文件导入 (.xlsx, .xls)
- CSV 文件导入
- 完整样式保留（字体、颜色、边框、对齐）
- 公式和计算值保留
- 合并单元格支持
- 条件格式支持
- 数据验证支持
- 超链接和富文本支持
- 图片导入支持（浮动图片和单元格图片）
- 工作簿辅助函数

## [1.2.0] - 2026-01-27

### 变更

- 新增：图表、筛选器、排序、透视表导入支持
- 新增：addConditionalFormatsToWorkbook、addFiltersToWorkbook、addSortsToWorkbook、addChartsToWorkbook、addPivotTablesToWorkbook、addImagesToWorkbook 等 API
- 新增：返回值支持 conditionalFormats、filters、sorts、charts、pivotTables 等结构
- 优化：README 全面美化，文档结构更清晰
- 变更：addImagesToWorkbook 仅保留 univerAPI 和 images 两个参数，去除 options
- 移除：importCsv 不再导出，统一用 importFile

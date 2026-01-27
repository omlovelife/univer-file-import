/**
 * univer-file-import
 *
 * Excel/CSV 文件导入工具，用于将 Excel 和 CSV 文件转换为 Univer 工作簿数据格式
 */

// 主要导入功能
export {
  importFile,
  ImageType,
  // 添加功能到工作簿的函数
  addConditionalFormatsToWorkbook,
  addFiltersToWorkbook,
  addSortsToWorkbook,
  addChartsToWorkbook,
  addPivotTablesToWorkbook,
  addImagesToWorkbook,
  // 类型导出
  type ImportedImage,
  type ImportedConditionalFormat,
  type ImportedFilter,
  type ImportedSort,
  type ImportedChart,
  type ImportedPivotTable,
  type ImageInsertOptions,
  type FileImportOptions,
  type FileImportResult,
} from './fileImport';

// 工作簿辅助函数
export {
  getDefaultWorkbookData,
  getWorkbookDataBySheets,
  getWorkbookData,
  buildSheetFrom2DArray,
  normalizeSource,
  createWorkbookItem,
  DEFAULT_ROW_COUNT,
  DEFAULT_COLUMN_COUNT,
  DEFAULT_COLUMN_WIDTH,
  DEFAULT_ROW_HEIGHT,
  type IWorkbookItem,
} from './workbookHelpers';

// 重新导出 Univer 类型（方便用户使用）
export type { IWorkbookData, IWorksheetData, ICellData } from '@univerjs/presets';

/**
 * Excel/CSV 文件导入工具
 *
 * 功能特性：
 * ✅ Excel 文件 (.xlsx, .xls)
 * ✅ CSV 文件 (.csv)
 * ✅ 保留所有工作表（包括空表）
 * ✅ 处理工作表名称中的特殊字符（>>>等）
 * ✅ 保持工作表顺序
 * ✅ 完整样式保留（字体、颜色、边框、对齐）
 * ✅ 公式和计算值保留（包括 TRANSPOSE 和数组公式）
 * ✅ 共享公式支持
 * ✅ 合并单元格支持
 * ✅ 条件格式
 * ✅ 数据验证
 * ✅ 超链接和富文本
 * ✅ 图片导入支持（浮动图片和单元格图片）
 *
 */

import { LocaleType } from '@univerjs/core';
import type { IWorkbookData, IWorksheetData } from '@univerjs/presets';
import ExcelJS from 'exceljs';
import { nanoid } from 'nanoid';
import {
  getDefaultWorkbookData,
  getWorkbookDataBySheets,
  buildSheetFrom2DArray,
  DEFAULT_ROW_COUNT,
  DEFAULT_COLUMN_COUNT,
  DEFAULT_COLUMN_WIDTH,
  DEFAULT_ROW_HEIGHT,
} from './workbookHelpers';

// 版本号（构建时注入）
declare const UNIVER_VERSION: string;

/**
 * 获取 Univer 版本号
 */
function getUniverVersion(): string {
  try {
    return UNIVER_VERSION;
  } catch {
    return '0.15.0';
  }
}

/**
 * 图片类型枚举
 */
export enum ImageType {
  /** 浮动图片 - 可自由移动，不随单元格变化 */
  FLOATING = 'floating',
  /** 单元格图片 - 嵌入单元格内，随单元格移动和缩放 */
  CELL = 'cell',
}

/**
 * 导入的图片信息
 */
export interface ImportedImage {
  /** 图片唯一 ID */
  id: string;
  /** 图片类型：浮动或单元格 */
  type: ImageType;
  /** 图片数据源（base64 或 URL） */
  source: string;
  /** 所属工作表的原始 ID（对应 IWorkbookData.sheets 的 key） */
  sheetId: string;
  /** 所属工作表名称 */
  sheetName: string;
  /** 图片位置信息 */
  position: {
    /** 起始列（0-based） */
    column: number;
    /** 起始行（0-based） */
    row: number;
    /** 列偏移（像素） */
    columnOffset: number;
    /** 行偏移（像素） */
    rowOffset: number;
  };
  /** 图片尺寸 */
  size: {
    /** 宽度（像素） */
    width: number;
    /** 高度（像素） */
    height: number;
  };
  /** 结束位置（用于浮动图片） */
  endPosition?: {
    column: number;
    row: number;
    columnOffset: number;
    rowOffset: number;
  };
  /** 图片标题 */
  title?: string;
  /** 图片描述 */
  description?: string;
}

/**
 * 导入结果（包含工作簿数据和图片信息）
 */
export interface ImportResult {
  /** 工作簿数据 */
  workbookData: IWorkbookData;
  /** 导入的图片列表 */
  images: ImportedImage[];
}

/**
 * Sheet ID 映射（原始 ID -> 实际创建的 ID）
 */
export type SheetIdMapping = Map<string, string>;

/**
 * 转义工作表名称中的特殊字符
 */
function escapeSheetName(name: string): string {
  if (!name) return 'Sheet1';
  return name;
}

/**
 * 将 ArrayBuffer 转换为 Base64 字符串（浏览器兼容）
 */
function arrayBufferToBase64(buffer: ArrayBuffer): string {
  const bytes = new Uint8Array(buffer);
  let binary = '';
  const chunkSize = 8192;
  for (let i = 0; i < bytes.length; i += chunkSize) {
    const chunk = bytes.subarray(i, Math.min(i + chunkSize, bytes.length));
    for (let j = 0; j < chunk.length; j++) {
      binary += String.fromCharCode(chunk[j]);
    }
  }
  return btoa(binary);
}

/**
 * 根据文件扩展名获取 MIME 类型
 */
function getImageMimeType(extension: string): string {
  const mimeMap: Record<string, string> = {
    png: 'image/png',
    jpg: 'image/jpeg',
    jpeg: 'image/jpeg',
    gif: 'image/gif',
    bmp: 'image/bmp',
    webp: 'image/webp',
    svg: 'image/svg+xml',
    tiff: 'image/tiff',
    tif: 'image/tiff',
    ico: 'image/x-icon',
  };
  return mimeMap[extension.toLowerCase()] || 'image/png';
}

/**
 * 检查数字格式是否为日期格式
 */
function isDateFormat(numFmt: string): boolean {
  if (!numFmt || numFmt === 'General') return false;

  const dateKeywords = [
    'yyyy', 'yy', 'mm', 'dd', 'm/d', 'd/m', 'h:mm', 'hh:mm', 'ss',
    '年', '月', '日', 'AM/PM', 'am/pm', '上午', '下午', '[$-',
  ];

  const lowerFmt = numFmt.toLowerCase();
  const hasDateKeyword = dateKeywords.some((keyword) => lowerFmt.includes(keyword.toLowerCase()));

  const dateFormatCodes = [14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46, 47, 176, 177, 178, 179, 180, 181, 182];

  const formatCode = parseInt(numFmt, 10);
  if (!isNaN(formatCode) && dateFormatCodes.includes(formatCode)) {
    return true;
  }

  return hasDateKeyword;
}

/**
 * 处理常见文件类型（Excel/CSV）
 */
async function handleFileImport(
  file: File,
  type: 'xlsx' | 'xls' | 'csv',
  includeImages: boolean = true,
): Promise<ImportResult> {
  try {
    let result: ImportResult;
    if (['xlsx', 'xls'].includes(type)) {
      result = await importExcelWithImages(file, includeImages);
    } else if (type === 'csv') {
      const workbookData = await importCsv(file);
      result = { workbookData, images: [] };
    } else {
      throw new Error('Unsupported file type');
    }
    return result;
  } catch (error) {
    console.error('handleFileImport error:', error);
    throw error;
  }
}

/**
 * 导入 Excel 文件（增强版 - 可选图片解析）
 */
async function importExcelWithImages(
  file: File,
  includeImages: boolean = true,
): Promise<ImportResult> {
  const fileSize = file.size;
  const fileName = file.name;
  const fileExt = fileName.split('.').pop()?.toLowerCase();

  if (fileSize > 10 * 1024 * 1024) {
    console.warn(`⚠️ 正在导入大文件 (${(fileSize / 1024 / 1024).toFixed(1)}MB)，可能需要较长时间...`);
  }

  const workbook = new ExcelJS.Workbook();
  const arrayBuffer = await file.arrayBuffer();

  try {
    await workbook.xlsx.load(arrayBuffer);
  } catch (error) {
    console.error('Excel 文件加载失败:', error);
    throw new Error(`无法加载 Excel 文件: ${(error as Error).message}. 如果是 .xls 格式，请先转换为 .xlsx`);
  }

  const univerWorkbook: IWorkbookData = {
    id: `workbook-${nanoid()}`,
    name: file.name.replace(/\.[^/.]+$/, ''),
    sheetOrder: [],
    appVersion: getUniverVersion(),
    locale: LocaleType.ZH_CN,
    styles: {},
    sheets: {},
    resources: [],
  };

  const allDrawings: Record<string, unknown> = {};
  const allImages: ImportedImage[] = [];

  workbook.worksheets.forEach((worksheet, sheetIndex) => {
    if (!worksheet) return;

    const sheetKey = `sheet-${nanoid()}`;
    univerWorkbook.sheetOrder.push(sheetKey);

    const cellData: Record<number, Record<number, unknown>> = {};
    const mergeData: unknown[] = [];
    const rowData: Record<number, unknown> = {};
    const columnData: Record<number, unknown> = {};
    const images: unknown[] = [];
    const charts: unknown[] = [];
    const conditionalFormats: unknown[] = [];
    const dataValidations: unknown[] = [];

    let maxRow = 0;
    let maxCol = 0;

    const sheetName = escapeSheetName(worksheet.name || `Sheet${sheetIndex + 1}`);

    worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
      const rowIndex = rowNumber - 1;
      maxRow = Math.max(maxRow, rowIndex);

      const rowInfo: Record<string, unknown> = {};
      let hasRowInfo = false;

      if (row.height && row.height > 0) {
        rowInfo.h = row.height;
        hasRowInfo = true;
      }

      if (row.hidden) {
        rowInfo.hd = 1;
        hasRowInfo = true;
      } else {
        rowInfo.hd = 0;
      }

      if (hasRowInfo) {
        rowData[rowIndex] = rowInfo;
      }

      row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
        const colIndex = colNumber - 1;
        maxCol = Math.max(maxCol, colIndex);

        if (cell.type === ExcelJS.ValueType.Null) return;

        // 处理 DISPIMG 公式
        if (cell.type === ExcelJS.ValueType.Formula && cell.formula) {
          const formula = typeof cell.formula === 'string' ? cell.formula : '';
          const dispImgMatch =
            formula.match(/_?xlfn\.?DISPIMG\s*\(\s*"([^"]+)"/i) ||
            formula.match(/DISPIMG\s*\(\s*"([^"]+)"/i);

          if (dispImgMatch) {
            const imageId = dispImgMatch[1];
            try {
              let foundImage: ExcelJS.Image | null = null;

              for (let imgIdx = 0; imgIdx < 100 && !foundImage; imgIdx++) {
                try {
                  const img = workbook.getImage(imgIdx);
                  if (img && img.buffer) {
                    foundImage = img;
                  }
                } catch {
                  // 索引不存在
                }
              }

              if (foundImage && foundImage.buffer) {
                const drawingId = `cell-img-${nanoid()}`;
                const base64 = arrayBufferToBase64(foundImage.buffer as ArrayBuffer);
                const extension = foundImage.extension || 'png';
                const mimeType = getImageMimeType(extension);
                const imageSource = `data:${mimeType};base64,${base64}`;

                const cellImage: ImportedImage = {
                  id: drawingId,
                  type: ImageType.CELL,
                  source: imageSource,
                  sheetId: sheetKey,
                  sheetName: sheetName,
                  position: {
                    column: colIndex,
                    row: rowIndex,
                    columnOffset: 0,
                    rowOffset: 0,
                  },
                  size: {
                    width: 100,
                    height: 100,
                  },
                  title: `CellImage_${imageId}`,
                  description: '',
                };
                allImages.push(cellImage);
                return;
              }
            } catch (imgError) {
              console.warn(`处理单元格图片失败 (${cell.address}):`, imgError);
            }
            return;
          }
        }

        let rawValue = getCellValue(cell);

        if (rawValue !== null && rawValue !== undefined && rawValue !== '') {
          if (typeof rawValue === 'number' && (isNaN(rawValue) || !isFinite(rawValue))) {
            rawValue = getOriginalCellValue(cell);
          } else if (typeof rawValue === 'string' && (rawValue === 'NaN' || rawValue.includes('NaN'))) {
            rawValue = getOriginalCellValue(cell);
          }
        }

        const hasValue = rawValue !== null && rawValue !== undefined && rawValue !== '';
        const hasFormula = cell.type === ExcelJS.ValueType.Formula && cell.formula;

        if (!hasValue && !hasFormula) return;

        if (!cellData[rowIndex]) {
          cellData[rowIndex] = {};
        }

        const cellValue: Record<string, unknown> = {};

        if (hasValue) {
          cellValue.v = rawValue;
        }

        const cellType = getCellType(cell, rawValue);
        if (cellType !== null && hasValue) {
          cellValue.t = cellType;
        }

        if (cell.type === ExcelJS.ValueType.Formula) {
          const formula = cell.formula;
          if (formula) {
            if (typeof formula === 'object' && 'sharedFormula' in (formula as object)) {
              cellValue.f = (formula as { sharedFormula: string }).sharedFormula;
            } else if (typeof formula === 'object' && 'result' in (formula as object)) {
              cellValue.f = (formula as { formula?: string }).formula || cell.formula;
            } else if (typeof formula === 'string') {
              cellValue.f = formula;
            }
            if (rawValue !== null && rawValue !== undefined) {
              cellValue.v = rawValue;
            }
          }
        }

        if (cell.type === ExcelJS.ValueType.RichText) {
          const richTextValue = cell.value as ExcelJS.CellRichTextValue;
          cellValue.p = convertRichText(richTextValue.richText);
        }

        if (cell.type === ExcelJS.ValueType.Hyperlink) {
          const hyperlinkValue = cell.value as ExcelJS.CellHyperlinkValue;
          cellValue.link = {
            url: hyperlinkValue.hyperlink,
            text: hyperlinkValue.text || hyperlinkValue.hyperlink,
          };
        }

        if (cell.style) {
          const style = convertCellStyle(cell.style);
          if (style) {
            cellValue.s = style;
          }
        }

        if (Object.keys(cellValue).length > 0) {
          cellData[rowIndex][colIndex] = cellValue;
        }
      });
    });

    // 处理列宽和隐藏列
    if (worksheet.columns && Array.isArray(worksheet.columns)) {
      worksheet.columns.forEach((column, index) => {
        if (column) {
          const colInfo: Record<string, unknown> = {};
          let hasColInfo = false;

          if (column.width && column.width > 0) {
            colInfo.w = column.width * 7.5;
            hasColInfo = true;
          }

          if (column.hidden) {
            colInfo.hd = 1;
            hasColInfo = true;
          } else {
            colInfo.hd = 0;
          }

          if (hasColInfo) {
            columnData[index] = colInfo;
          }
        }
      });
    }

    // 处理合并单元格
    if (worksheet.model && (worksheet.model as { merges?: string[] }).merges) {
      ((worksheet.model as { merges?: string[] }).merges || []).forEach((merge: string) => {
        const [start, end] = merge.split(':');
        const startCell = worksheet.getCell(start);
        const endCell = worksheet.getCell(end);

        mergeData.push({
          startRow: Number(startCell.row) - 1,
          endRow: Number(endCell.row) - 1,
          startColumn: Number(startCell.col) - 1,
          endColumn: Number(endCell.col) - 1,
        });
      });
    }

    // 处理条件格式
    const wsAny = worksheet as unknown as { conditionalFormattings?: Record<string, unknown> };
    if (wsAny.conditionalFormattings) {
      try {
        Object.values(wsAny.conditionalFormattings || {}).forEach((cf: unknown) => {
          const cfObj = cf as { ref?: string; type?: string; priority?: number; rules?: unknown[] };
          if (cfObj && cfObj.ref) {
            conditionalFormats.push({
              cfId: `cf-${nanoid()}`,
              ref: cfObj.ref,
              type: cfObj.type,
              priority: cfObj.priority,
              rules: cfObj.rules || [],
            });
          }
        });
      } catch (error) {
        console.warn('处理条件格式时出错:', error);
      }
    }

    // 处理数据验证
    worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      row.eachCell({ includeEmpty: false }, (cell) => {
        if (cell.dataValidation) {
          const validation = cell.dataValidation;
          dataValidations.push({
            row: rowNumber - 1,
            col: Number(cell.col) - 1,
            type: validation.type,
            operator: validation.operator,
            formula1: validation.formulae?.[0],
            formula2: validation.formulae?.[1],
            allowBlank: validation.allowBlank,
            showInputMessage: validation.showInputMessage,
            promptTitle: validation.promptTitle,
            prompt: validation.prompt,
            showErrorMessage: validation.showErrorMessage,
            errorStyle: validation.errorStyle,
            errorTitle: validation.errorTitle,
            error: validation.error,
          });
        }
      });
    });

    // 处理图片
    if (includeImages && worksheet.getImages && typeof worksheet.getImages === 'function') {
      try {
        const worksheetImages = worksheet.getImages();

        worksheetImages.forEach((img, imgIndex) => {
          if (img && img.imageId !== undefined) {
            try {
              const imageMedia = workbook.getImage(Number(img.imageId));
              if (!imageMedia || !imageMedia.buffer) {
                console.warn(`图片 ${img.imageId} 没有有效的数据`);
                return;
              }

              const drawingId = `drawing-${nanoid()}`;
              const base64 = arrayBufferToBase64(imageMedia.buffer as ArrayBuffer);
              const extension = imageMedia.extension || 'png';
              const mimeType = getImageMimeType(extension);
              const imageSource = `data:${mimeType};base64,${base64}`;

              const range = img.range as {
                tl?: { col?: number; nativeCol?: number; row?: number; nativeRow?: number };
                br?: { col?: number; nativeCol?: number; row?: number; nativeRow?: number };
                ext?: { width?: number; height?: number };
              };
              let startColumn = 0;
              let startRow = 0;
              let endColumn = 1;
              let endRow = 1;
              let startColumnOffset = 0;
              let startRowOffset = 0;
              let endColumnOffset = 0;
              let endRowOffset = 0;

              let imageWidth = 0;
              let imageHeight = 0;

              const EMU_PER_PIXEL = 9525;
              if (range?.ext) {
                imageWidth = (range.ext.width ?? 0) / EMU_PER_PIXEL;
                imageHeight = (range.ext.height ?? 0) / EMU_PER_PIXEL;
              }

              if (range) {
                if (range.tl) {
                  startColumn = Math.floor(range.tl.col ?? range.tl.nativeCol ?? 0);
                  startRow = Math.floor(range.tl.row ?? range.tl.nativeRow ?? 0);
                  const colDecimal = (range.tl.col ?? range.tl.nativeCol ?? 0) % 1;
                  const rowDecimal = (range.tl.row ?? range.tl.nativeRow ?? 0) % 1;
                  startColumnOffset = colDecimal * DEFAULT_COLUMN_WIDTH;
                  startRowOffset = rowDecimal * DEFAULT_ROW_HEIGHT;
                }
                if (range.br) {
                  endColumn = Math.floor(range.br.col ?? range.br.nativeCol ?? startColumn + 1);
                  endRow = Math.floor(range.br.row ?? range.br.nativeRow ?? startRow + 1);
                  const colDecimal = (range.br.col ?? range.br.nativeCol ?? 0) % 1;
                  const rowDecimal = (range.br.row ?? range.br.nativeRow ?? 0) % 1;
                  endColumnOffset = colDecimal * DEFAULT_COLUMN_WIDTH;
                  endRowOffset = rowDecimal * DEFAULT_ROW_HEIGHT;
                } else {
                  endColumn = startColumn + 1;
                  endRow = startRow + 1;
                }
              }

              if (imageWidth <= 0 || imageHeight <= 0) {
                imageWidth = (endColumn - startColumn) * DEFAULT_COLUMN_WIDTH + endColumnOffset - startColumnOffset;
                imageHeight = (endRow - startRow) * DEFAULT_ROW_HEIGHT + endRowOffset - startRowOffset;
              }

              const width = Math.max(imageWidth, 20);
              const height = Math.max(imageHeight, 20);

              const sheetDrawing = {
                unitId: univerWorkbook.id,
                subUnitId: sheetKey,
                drawingId: drawingId,
                drawingType: 1,
                imageSourceType: 0,
                source: imageSource,
                title: `Image_${imgIndex + 1}`,
                description: '',
                transform: {
                  left: startColumn * DEFAULT_COLUMN_WIDTH + startColumnOffset,
                  top: startRow * DEFAULT_ROW_HEIGHT + startRowOffset,
                  width: width,
                  height: height,
                  angle: 0,
                  flipX: false,
                  flipY: false,
                  skewX: 0,
                  skewY: 0,
                },
                sheetTransform: {
                  from: {
                    column: startColumn,
                    columnOffset: startColumnOffset,
                    row: startRow,
                    rowOffset: startRowOffset,
                  },
                  to: {
                    column: endColumn,
                    columnOffset: endColumnOffset,
                    row: endRow,
                    rowOffset: endRowOffset,
                  },
                },
                anchorType: 0,
              };

              allDrawings[drawingId] = sheetDrawing;
              images.push({
                drawingId,
                sheetId: sheetKey,
              });

              const importedImage: ImportedImage = {
                id: drawingId,
                type: ImageType.FLOATING,
                source: imageSource,
                sheetId: sheetKey,
                sheetName: sheetName,
                position: {
                  column: startColumn,
                  row: startRow,
                  columnOffset: startColumnOffset,
                  rowOffset: startRowOffset,
                },
                size: {
                  width: width,
                  height: height,
                },
                endPosition: {
                  column: endColumn,
                  row: endRow,
                  columnOffset: endColumnOffset,
                  rowOffset: endRowOffset,
                },
                title: `Image_${imgIndex + 1}`,
                description: '',
              };
              allImages.push(importedImage);
            } catch (imgError) {
              console.warn(`处理图片 ${imgIndex} 失败:`, imgError);
            }
          }
        });
      } catch (error) {
        console.warn('处理图片时出错:', error);
      }
    }

    // 处理图表
    const wsDrawings = worksheet as unknown as { drawings?: unknown[] };
    if (wsDrawings.drawings && Array.isArray(wsDrawings.drawings)) {
      try {
        wsDrawings.drawings.forEach((drawing: unknown) => {
          const drawingObj = drawing as { type?: string; chartType?: string; range?: unknown; data?: unknown };
          if (drawingObj.type === 'chart') {
            charts.push({
              chartId: `chart-${nanoid()}`,
              type: drawingObj.chartType || 'unknown',
              range: drawingObj.range,
              data: drawingObj.data,
            });
          }
        });
      } catch (error) {
        console.warn('处理图表时出错:', error);
      }
    }

    const univerSheet: IWorksheetData = {
      id: sheetKey,
      name: sheetName,
      rowCount: Math.max(DEFAULT_ROW_COUNT, maxRow + 1),
      columnCount: Math.max(DEFAULT_COLUMN_COUNT, maxCol + 1),
      defaultColumnWidth: DEFAULT_COLUMN_WIDTH,
      defaultRowHeight: DEFAULT_ROW_HEIGHT,
      zoomRatio: 1,
      scrollTop: 0,
      scrollLeft: 0,
      hidden: worksheet.state === 'hidden' ? 1 : 0,
      tabColor: '',
      cellData,
      rowData: Object.keys(rowData).length > 0 ? rowData : undefined,
      columnData: Object.keys(columnData).length > 0 ? columnData : undefined,
      mergeData: mergeData.length > 0 ? mergeData : undefined,
      ...(images.length > 0 && { images }),
      ...(charts.length > 0 && { charts }),
      ...(conditionalFormats.length > 0 && { conditionalFormats }),
      ...(dataValidations.length > 0 && { dataValidations }),
    } as IWorksheetData;

    univerWorkbook.sheets[sheetKey] = univerSheet;
  });

  if (Object.keys(allDrawings).length > 0) {
    const drawingResourceData = {
      drawings: allDrawings,
      drawingsOrder: Object.keys(allDrawings),
    };

    univerWorkbook.resources!.push({
      name: 'SHEET_DRAWING_PLUGIN',
      data: JSON.stringify(drawingResourceData),
    });
  }

  return {
    workbookData: univerWorkbook,
    images: allImages,
  };
}

/**
 * 获取单元格的原始值
 */
function getOriginalCellValue(cell: ExcelJS.Cell): unknown {
  const value = cell.value;

  if (value === null || value === undefined) {
    return '';
  }

  if (value instanceof Date) {
    if (!isNaN(value.getTime())) {
      const year = value.getFullYear();
      const month = value.getMonth() + 1;
      const day = value.getDate();
      return `${year}/${month}/${day}`;
    }
    return '';
  }

  if (typeof value === 'number') {
    if (!isNaN(value) && isFinite(value)) {
      return value;
    }
    return '';
  }

  if (typeof value === 'string') {
    return value;
  }

  if (typeof value === 'boolean') {
    return value;
  }

  if (typeof value === 'object' && 'richText' in value) {
    return (value as ExcelJS.CellRichTextValue).richText.map((rt) => rt.text).join('');
  }

  if (typeof value === 'object' && 'hyperlink' in value) {
    return (value as ExcelJS.CellHyperlinkValue).text || (value as ExcelJS.CellHyperlinkValue).hyperlink;
  }

  if (typeof value === 'object' && 'result' in value) {
    const result = (value as { result?: unknown }).result;
    if (result !== null && result !== undefined) {
      if (result instanceof Date) {
        const year = result.getFullYear();
        const month = result.getMonth() + 1;
        const day = result.getDate();
        return `${year}/${month}/${day}`;
      }
      return result;
    }
  }

  try {
    const str = String(value);
    if (str !== '[object Object]' && str !== 'NaN' && str !== 'undefined' && str !== 'null') {
      return str;
    }
  } catch {
    // 忽略转换错误
  }

  return '';
}

/**
 * 获取单元格值
 */
function getCellValue(cell: ExcelJS.Cell): unknown {
  if (cell.type === ExcelJS.ValueType.Formula) {
    return cell.result !== undefined ? cell.result : cell.value;
  }

  if (cell.type === ExcelJS.ValueType.RichText) {
    const richTextValue = cell.value as ExcelJS.CellRichTextValue;
    return richTextValue.richText.map((rt) => rt.text).join('');
  }

  if (cell.type === ExcelJS.ValueType.Date) {
    const dateValue = cell.value;
    const numFmt = cell.style?.numFmt;

    if (dateValue === null || dateValue === undefined) {
      return '';
    }

    if (dateValue instanceof Date) {
      if (!isNaN(dateValue.getTime())) {
        const formatted = formatDateByPattern(dateValue, numFmt);
        if (formatted && formatted !== '' && !formatted.includes('NaN')) {
          return formatted;
        }
        return dateValue.toLocaleDateString('zh-CN');
      }
      return String(dateValue);
    }

    if (typeof dateValue === 'number' && !isNaN(dateValue) && isFinite(dateValue) && dateValue > 0) {
      try {
        const date = excelSerialToDate(dateValue);
        if (date && date instanceof Date && !isNaN(date.getTime())) {
          const formatted = formatDateByPattern(date, numFmt);
          if (formatted && formatted !== '' && !formatted.includes('NaN')) {
            return formatted;
          }
          return date.toLocaleDateString('zh-CN');
        }
      } catch (error) {
        console.warn('日期转换失败:', error, '原始值:', dateValue);
      }
      return dateValue;
    }

    if (typeof dateValue === 'string' && dateValue.trim()) {
      if (numFmt && isDateFormat(numFmt)) {
        const parsedDate = parseDateString(dateValue);
        if (parsedDate && !isNaN(parsedDate.getTime())) {
          const formatted = formatDateByPattern(parsedDate, numFmt);
          if (formatted && formatted !== '' && !formatted.includes('NaN')) {
            return formatted;
          }
        }
      }
      return dateValue;
    }

    return dateValue;
  }

  if (cell.type === ExcelJS.ValueType.Hyperlink) {
    const hyperlinkValue = cell.value as ExcelJS.CellHyperlinkValue;
    return hyperlinkValue.text || hyperlinkValue.hyperlink;
  }

  if (cell.type === ExcelJS.ValueType.Merge) {
    return '';
  }

  if (cell.type === ExcelJS.ValueType.Error) {
    return cell.value;
  }

  if (cell.type === ExcelJS.ValueType.Number) {
    const numValue = cell.value as number;
    const numFmt = cell.style?.numFmt;

    if (numFmt && isDateFormat(numFmt)) {
      if (typeof numValue === 'number' && !isNaN(numValue) && isFinite(numValue) && numValue > 0) {
        try {
          const date = excelSerialToDate(numValue);
          if (date && date instanceof Date && !isNaN(date.getTime())) {
            const formatted = formatDateByPattern(date, numFmt);
            if (formatted && formatted !== '' && !formatted.includes('NaN')) {
              return formatted;
            }
            return date.toLocaleDateString('zh-CN');
          }
        } catch (error) {
          console.warn('日期转换失败:', error, '原始值:', numValue);
        }
        return numValue;
      }
    }

    if (typeof numValue === 'number' && !isNaN(numValue) && isFinite(numValue)) {
      return numValue;
    }
    return numValue;
  }

  if (cell.type === ExcelJS.ValueType.String) {
    const strValue = cell.value as string;
    const numFmt = cell.style?.numFmt;

    if (numFmt && isDateFormat(numFmt) && strValue) {
      const parsedDate = parseDateString(strValue);
      if (parsedDate && !isNaN(parsedDate.getTime())) {
        const formatted = formatDateByPattern(parsedDate, numFmt);
        if (formatted && formatted !== '' && !formatted.includes('NaN')) {
          return formatted;
        }
      }
    }

    return strValue;
  }

  return cell.value;
}

/**
 * 获取单元格类型
 */
function getCellType(cell: ExcelJS.Cell, actualValue?: unknown): number | null {
  if (actualValue !== undefined && actualValue !== null && actualValue !== '') {
    if (typeof actualValue === 'boolean') {
      return 3;
    }
    if (typeof actualValue === 'number') {
      return 2;
    }
    if (typeof actualValue === 'string') {
      return 1;
    }
  }

  if (cell.type === ExcelJS.ValueType.Boolean) {
    return 3;
  }
  if (cell.type === ExcelJS.ValueType.Number) {
    return 2;
  }
  if (cell.type === ExcelJS.ValueType.Date) {
    return 1;
  }
  if (cell.type === ExcelJS.ValueType.String || cell.type === ExcelJS.ValueType.RichText) {
    return 1;
  }
  return null;
}

/**
 * 转换富文本格式
 */
function convertRichText(richText: ExcelJS.RichText[]): unknown {
  if (!richText || richText.length === 0) {
    return undefined;
  }

  const body: { dataStream: string; textRuns: unknown[] } = {
    dataStream: '',
    textRuns: [],
  };

  let currentIndex = 0;
  richText.forEach((rt) => {
    const text = rt.text || '';
    const length = text.length;

    body.dataStream += text;

    if (rt.font) {
      const textRun: Record<string, unknown> = {
        st: currentIndex,
        ed: currentIndex + length,
      };

      const ts: Record<string, unknown> = {};
      if (rt.font.bold) ts.bl = 1;
      if (rt.font.italic) ts.it = 1;
      if (rt.font.underline) ts.ul = { s: 1 };
      if (rt.font.strike) ts.st = { s: 1 };
      if (rt.font.size) ts.fs = rt.font.size;
      if (rt.font.name) ts.ff = rt.font.name;
      if (rt.font.color?.argb) {
        ts.cl = { rgb: `#${rt.font.color.argb.slice(-6)}` };
      }

      if (Object.keys(ts).length > 0) {
        textRun.ts = ts;
        body.textRuns.push(textRun);
      }
    }

    currentIndex += length;
  });

  return body.textRuns.length > 0 ? body : undefined;
}

/**
 * 转换单元格样式
 */
function convertCellStyle(style: Partial<ExcelJS.Style>): unknown {
  const univerStyle: Record<string, unknown> = {};

  if (style.font) {
    if (style.font.bold) univerStyle.bl = 1;
    if (style.font.italic) univerStyle.it = 1;
    if (style.font.underline) univerStyle.ul = { s: 1 };
    if (style.font.strike) univerStyle.st = { s: 1 };
    if (style.font.size) univerStyle.fs = style.font.size;
    if (style.font.name) univerStyle.ff = style.font.name;
    if (style.font.color && style.font.color.argb) {
      univerStyle.cl = { rgb: `#${style.font.color.argb.slice(-6)}` };
    }
  }

  if (style.fill && style.fill.type === 'pattern') {
    const patternFill = style.fill as ExcelJS.FillPattern;
    if (patternFill.fgColor && patternFill.fgColor.argb) {
      univerStyle.bg = { rgb: `#${patternFill.fgColor.argb.slice(-6)}` };
    }
  }

  if (style.alignment) {
    if (style.alignment.horizontal) {
      univerStyle.ht = convertAlignment(style.alignment.horizontal);
    }
    if (style.alignment.vertical) {
      univerStyle.vt = convertVerticalAlignment(style.alignment.vertical);
    }
    if (style.alignment.wrapText) {
      univerStyle.tb = 3;
    }
  }

  if (style.border) {
    const bd = convertBorder(style.border);
    if (bd) {
      univerStyle.bd = bd;
    }
  }

  if (style.numFmt) {
    const numFormat = convertNumFormat(style.numFmt);
    if (numFormat) {
      univerStyle.n = numFormat;
    }
  }

  return Object.keys(univerStyle).length > 0 ? univerStyle : undefined;
}

/**
 * 转换水平对齐方式
 */
function convertAlignment(alignment: string): number {
  const alignmentMap: Record<string, number> = {
    left: 1,
    center: 2,
    right: 3,
  };
  return alignmentMap[alignment] || 1;
}

/**
 * 转换垂直对齐方式
 */
function convertVerticalAlignment(alignment: string): number {
  const alignmentMap: Record<string, number> = {
    top: 1,
    middle: 2,
    bottom: 3,
  };
  return alignmentMap[alignment] || 2;
}

/**
 * 转换边框样式
 */
function convertBorder(border: Partial<ExcelJS.Borders>): unknown {
  const result: Record<string, unknown> = {};

  if (border.top) {
    result.t = { s: 1, cl: { rgb: '#000000' } };
  }
  if (border.bottom) {
    result.b = { s: 1, cl: { rgb: '#000000' } };
  }
  if (border.left) {
    result.l = { s: 1, cl: { rgb: '#000000' } };
  }
  if (border.right) {
    result.r = { s: 1, cl: { rgb: '#000000' } };
  }

  return Object.keys(result).length > 0 ? result : undefined;
}

/**
 * 转换数字格式
 */
function convertNumFormat(numFmt: string): unknown {
  if (!numFmt || numFmt === 'General') {
    return undefined;
  }

  const formatResult: Record<string, unknown> = {
    pattern: numFmt,
  };

  const formatInfo = parseNumFormat(numFmt);
  if (formatInfo) {
    Object.assign(formatResult, formatInfo);
  }

  return formatResult;
}

/**
 * 解析 Excel 数字格式字符串
 */
function parseNumFormat(numFmt: string): Record<string, unknown> | undefined {
  const result: Record<string, unknown> = {};

  if (numFmt.includes('%')) {
    result.isPercent = true;
  }

  if (numFmt.includes('$') || numFmt.includes('¥') || numFmt.includes('€') || numFmt.includes('£')) {
    result.isCurrency = true;
  }

  if (numFmt.toUpperCase().includes('E+') || numFmt.toUpperCase().includes('E-')) {
    result.isScientific = true;
  }

  if (numFmt.includes(',')) {
    result.hasThousandsSeparator = true;
  }

  const decimalMatch = numFmt.match(/\.([0#?]+)/);
  if (decimalMatch) {
    result.decimalPlaces = decimalMatch[1].length;
  } else {
    result.decimalPlaces = 0;
  }

  const dateTimePatterns = ['yyyy', 'yy', 'mm', 'dd', 'hh', 'ss', 'AM/PM', 'am/pm'];
  const isDateTime = dateTimePatterns.some((pattern) =>
    numFmt.toLowerCase().includes(pattern.toLowerCase()),
  );
  if (isDateTime) {
    result.isDateTime = true;
  }

  if (numFmt.includes('[Red]') || (numFmt.includes('(') && numFmt.includes(')'))) {
    result.hasNegativeFormat = true;
  }

  return Object.keys(result).length > 0 ? result : undefined;
}

/**
 * 将 Excel 序列号转换为 JavaScript Date 对象
 */
function excelSerialToDate(serial: number): Date | null {
  if (typeof serial !== 'number' || isNaN(serial) || !isFinite(serial)) {
    return null;
  }

  const excelEpoch = new Date(Date.UTC(1899, 11, 30));
  const msPerDay = 24 * 60 * 60 * 1000;

  try {
    const date = new Date(excelEpoch.getTime() + serial * msPerDay);
    if (date instanceof Date && !isNaN(date.getTime())) {
      return date;
    }
  } catch (error) {
    console.warn('Excel 序列号转换日期失败:', error, '序列号:', serial);
  }

  return null;
}

/**
 * 解析日期字符串
 */
function parseDateString(dateStr: string): Date | null {
  if (!dateStr || typeof dateStr !== 'string') {
    return null;
  }

  const trimmed = dateStr.trim();
  if (!trimmed) {
    return null;
  }

  // yyyy/m/d 或 yyyy/mm/dd
  const pattern1 = /^(\d{4})\/(\d{1,2})\/(\d{1,2})$/;
  const match1 = trimmed.match(pattern1);
  if (match1) {
    const year = parseInt(match1[1], 10);
    const month = parseInt(match1[2], 10) - 1;
    const day = parseInt(match1[3], 10);
    if (!isNaN(year) && !isNaN(month) && !isNaN(day)) {
      const date = new Date(year, month, day);
      if (!isNaN(date.getTime())) {
        return date;
      }
    }
  }

  // yyyy-m-d 或 yyyy-mm-dd
  const pattern2 = /^(\d{4})-(\d{1,2})-(\d{1,2})$/;
  const match2 = trimmed.match(pattern2);
  if (match2) {
    const year = parseInt(match2[1], 10);
    const month = parseInt(match2[2], 10) - 1;
    const day = parseInt(match2[3], 10);
    if (!isNaN(year) && !isNaN(month) && !isNaN(day)) {
      const date = new Date(year, month, day);
      if (!isNaN(date.getTime())) {
        return date;
      }
    }
  }

  // yyyy年m月d日
  const pattern3 = /^(\d{4})年(\d{1,2})月(\d{1,2})日$/;
  const match3 = trimmed.match(pattern3);
  if (match3) {
    const year = parseInt(match3[1], 10);
    const month = parseInt(match3[2], 10) - 1;
    const day = parseInt(match3[3], 10);
    if (!isNaN(year) && !isNaN(month) && !isNaN(day)) {
      const date = new Date(year, month, day);
      if (!isNaN(date.getTime())) {
        return date;
      }
    }
  }

  // m/d/yyyy 或 mm/dd/yyyy
  const pattern4 = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/;
  const match4 = trimmed.match(pattern4);
  if (match4) {
    const month = parseInt(match4[1], 10) - 1;
    const day = parseInt(match4[2], 10);
    const year = parseInt(match4[3], 10);
    if (!isNaN(year) && !isNaN(month) && !isNaN(day)) {
      const date = new Date(year, month, day);
      if (!isNaN(date.getTime())) {
        return date;
      }
    }
  }

  try {
    const date = new Date(trimmed);
    if (date instanceof Date && !isNaN(date.getTime())) {
      const year = date.getFullYear();
      if (year >= 1900 && year <= 2100) {
        return date;
      }
    }
  } catch {
    // 忽略解析错误
  }

  return null;
}

/**
 * 根据 Excel 数字格式模式格式化日期
 */
function formatDateByPattern(date: Date, numFmt?: string): string {
  if (!date || !(date instanceof Date) || isNaN(date.getTime())) {
    return '';
  }

  try {
    const year = date.getFullYear();
    const month = date.getMonth() + 1;
    const day = date.getDate();
    const hours = date.getHours();
    const minutes = date.getMinutes();
    const seconds = date.getSeconds();

    if (isNaN(year) || isNaN(month) || isNaN(day) || isNaN(hours) || isNaN(minutes) || isNaN(seconds)) {
      return '';
    }

    if (year < 1900 || year > 2100 || month < 1 || month > 12 || day < 1 || day > 31) {
      return '';
    }

    const safeFormat = (str: string): string => {
      if (!str || str.includes('NaN') || str.includes('undefined') || str.includes('null')) {
        return '';
      }
      return str;
    };

    if (!numFmt || numFmt === 'General') {
      const result = `${year}/${month}/${day}`;
      return safeFormat(result);
    }

    if (numFmt.includes('yyyy') && numFmt.includes('/')) {
      if (numFmt.includes('h:mm') || numFmt.includes('hh:mm')) {
        const mm = String(minutes).padStart(2, '0');
        const ss = String(seconds).padStart(2, '0');
        if (numFmt.includes('mm/dd')) {
          const result = `${year}/${String(month).padStart(2, '0')}/${String(day).padStart(2, '0')} ${hours}:${mm}:${ss}`;
          return safeFormat(result);
        }
        const result = `${year}/${month}/${day} ${hours}:${mm}:${ss}`;
        return safeFormat(result);
      }
      if (numFmt.includes('mm') && numFmt.includes('dd')) {
        const result = `${year}/${String(month).padStart(2, '0')}/${String(day).padStart(2, '0')}`;
        return safeFormat(result);
      }
      const result = `${year}/${month}/${day}`;
      return safeFormat(result);
    }

    if (numFmt.includes('yyyy') && numFmt.includes('-')) {
      if (numFmt.includes('mm') && numFmt.includes('dd')) {
        const result = `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
        return safeFormat(result);
      }
      const result = `${year}-${month}-${day}`;
      return safeFormat(result);
    }

    if (numFmt.match(/m+\/d+\/y+/i)) {
      const yy = String(year).slice(-2);
      if (numFmt.includes('mm') && numFmt.includes('dd')) {
        const result = `${String(month).padStart(2, '0')}/${String(day).padStart(2, '0')}/${yy}`;
        return safeFormat(result);
      }
      const result = `${month}/${day}/${yy}`;
      return safeFormat(result);
    }

    if (numFmt.match(/d+\/m+\/y+/i)) {
      const yy = String(year).slice(-2);
      if (numFmt.includes('mm') && numFmt.includes('dd')) {
        const result = `${String(day).padStart(2, '0')}/${String(month).padStart(2, '0')}/${yy}`;
        return safeFormat(result);
      }
      const result = `${day}/${month}/${yy}`;
      return safeFormat(result);
    }

    if (numFmt.includes('年') && numFmt.includes('月') && numFmt.includes('日')) {
      const result = `${year}年${month}月${day}日`;
      return safeFormat(result);
    }

    const result = `${year}/${month}/${day}`;
    if (result.includes('NaN')) {
      try {
        return date.toLocaleDateString('zh-CN');
      } catch {
        return '';
      }
    }
    return result;
  } catch (error) {
    console.error('formatDateByPattern: 格式化日期时出错', error, date);
    try {
      if (date && date instanceof Date && !isNaN(date.getTime())) {
        return date.toLocaleDateString('zh-CN');
      }
    } catch {
      // 忽略错误
    }
    return '';
  }
}

/**
 * 将列号转换为 Excel 列名
 */
function columnIndexToLetter(columnIndex: number): string {
  let letter = '';
  let temp = columnIndex;
  while (temp >= 0) {
    letter = String.fromCharCode((temp % 26) + 65) + letter;
    temp = Math.floor(temp / 26) - 1;
  }
  return letter;
}

/**
 * 将行列索引转换为单元格地址
 */
function getCellAddress(row: number, column: number): string {
  return `${columnIndexToLetter(column)}${row + 1}`;
}

// ==================== CSV 导入 ====================

/**
 * CSV 导入选项
 */
export interface CsvImportOptions {
  /** 分隔符，默认自动检测 */
  delimiter?: string;
  /** 编码，默认 UTF-8 */
  encoding?: string;
  /** 是否有表头，默认 true */
  hasHeader?: boolean;
  /** 工作表名称 */
  sheetName?: string;
  /** 跳过空行，默认 true */
  skipEmptyLines?: boolean;
  /** 修剪单元格值的空白，默认 true */
  trimValues?: boolean;
}

/**
 * 增强版 CSV 导入函数
 */
export async function importCsv(
  file: File | Blob,
  options: CsvImportOptions = {},
): Promise<IWorkbookData> {
  const {
    delimiter,
    encoding = 'UTF-8',
    hasHeader = true,
    sheetName = 'Sheet1',
    skipEmptyLines = true,
    trimValues = true,
  } = options;

  try {
    let text: string = '';
    try {
      const arrayBuffer = await file.arrayBuffer();
      const decoder = new TextDecoder(encoding);
      text = decoder.decode(arrayBuffer);
    } catch {
      console.warn(`编码 ${encoding} 失败，尝试其他编码`);
      const arrayBuffer = await file.arrayBuffer();
      const fallbackEncodings = ['UTF-8', 'GBK', 'GB2312', 'ISO-8859-1'];
      let decoded = false;

      for (const enc of fallbackEncodings) {
        try {
          const decoder = new TextDecoder(enc);
          text = decoder.decode(arrayBuffer);
          decoded = true;
          break;
        } catch {
          continue;
        }
      }

      if (!decoded) {
        throw new Error('无法解码文件内容');
      }
    }

    const detectedDelimiter = delimiter || detectDelimiter(text);
    const rows = parseCsvText(text, detectedDelimiter, skipEmptyLines);

    if (rows.length === 0) {
      return getDefaultWorkbookData({ name: sheetName });
    }

    const processedRows = rows.map((row) =>
      row.map((cell) => {
        const trimmedCell = trimValues ? cell.trim() : cell;
        return convertCellType(trimmedCell);
      }),
    );

    const sheet: IWorksheetData = buildSheetFrom2DArray(sheetName, processedRows);

    if (hasHeader && processedRows.length > 0) {
      if (!sheet.cellData[0]) {
        sheet.cellData[0] = {};
      }
      Object.keys(sheet.cellData[0]).forEach((colKey) => {
        const col = parseInt(colKey, 10);
        if (sheet.cellData[0][col]) {
          (sheet.cellData[0][col] as Record<string, unknown>).s = {
            bl: 1,
            fs: 11,
            ht: 2,
            vt: 2,
            bg: { rgb: '#F5F5F5' },
            bd: {
              b: { s: 1, cl: { rgb: '#D0D0D0' } },
            },
          };
        }
      });

      if (!sheet.rowData) {
        sheet.rowData = {};
      }
      (sheet.rowData as Record<number, unknown>)[0] = {
        h: 25,
        hd: 0,
      };
    }

    if (processedRows.length > 0) {
      const columnWidths: Record<number, number> = {};

      processedRows.forEach((row) => {
        row.forEach((cell, colIndex) => {
          const cellStr = String(cell);
          const chineseChars = (cellStr.match(/[\u4e00-\u9fa5]/g) || []).length;
          const otherChars = cellStr.length - chineseChars;
          const estimatedWidth = chineseChars * 12 + otherChars * 8 + 20;

          columnWidths[colIndex] = Math.max(
            columnWidths[colIndex] || 80,
            Math.min(estimatedWidth, 400),
          );
        });
      });

      sheet.columnData = {};
      Object.keys(columnWidths).forEach((colKey) => {
        const col = parseInt(colKey, 10);
        (sheet.columnData as Record<number, unknown>)![col] = {
          w: columnWidths[col],
          hd: 0,
        };
      });
    }

    return getWorkbookDataBySheets([sheet]);
  } catch (error) {
    console.error('CSV 导入失败:', error);
    throw new Error(`CSV 导入失败: ${(error as Error).message}`);
  }
}

/**
 * 检测 CSV 分隔符
 */
function detectDelimiter(text: string): string {
  const delimiters = [',', ';', '\t', '|', ':'];
  const firstLines = text.split(/\r?\n/).slice(0, 5);

  let maxScore = 0;
  let detectedDelimiter = ',';

  delimiters.forEach((delimiter) => {
    let score = 0;
    let consistentCount = true;
    let firstLineCount = -1;

    firstLines.forEach((line) => {
      if (!line.trim()) return;

      const count = (line.match(new RegExp(`\\${delimiter}`, 'g')) || []).length;

      if (firstLineCount === -1) {
        firstLineCount = count;
      } else if (firstLineCount !== count) {
        consistentCount = false;
      }

      score += count;
    });

    if (consistentCount && firstLineCount > 0) {
      score += 100;
    }

    if (score > maxScore) {
      maxScore = score;
      detectedDelimiter = delimiter;
    }
  });

  return detectedDelimiter;
}

/**
 * 解析 CSV 文本
 */
function parseCsvText(text: string, delimiter: string, skipEmptyLines: boolean): string[][] {
  const rows: string[][] = [];
  let currentRow: string[] = [];
  let currentCell = '';
  let inQuotes = false;
  let i = 0;

  while (i < text.length) {
    const char = text[i];
    const nextChar = text[i + 1];

    if (char === '"') {
      if (inQuotes && nextChar === '"') {
        currentCell += '"';
        i += 2;
        continue;
      } else if (inQuotes) {
        inQuotes = false;
        i++;
        continue;
      } else if (currentCell === '' || text[i - 1] === delimiter) {
        inQuotes = true;
        i++;
        continue;
      } else {
        currentCell += char;
        i++;
        continue;
      }
    } else if (char === delimiter && !inQuotes) {
      currentRow.push(currentCell);
      currentCell = '';
      i++;
    } else if ((char === '\n' || char === '\r') && !inQuotes) {
      if (char === '\r' && nextChar === '\n') {
        i++;
      }

      currentRow.push(currentCell);
      currentCell = '';

      if (!skipEmptyLines || currentRow.some((cell) => cell !== '')) {
        rows.push(currentRow);
      }
      currentRow = [];
      i++;
    } else {
      currentCell += char;
      i++;
    }
  }

  if (currentCell !== '' || currentRow.length > 0) {
    currentRow.push(currentCell);
    if (!skipEmptyLines || currentRow.some((cell) => cell !== '')) {
      rows.push(currentRow);
    }
  }

  if (rows.length > 0) {
    const maxCols = Math.max(...rows.map((row) => row.length));
    rows.forEach((row) => {
      while (row.length < maxCols) {
        row.push('');
      }
    });
  }

  return rows;
}

/**
 * 转换单元格类型
 */
function convertCellType(value: string): unknown {
  if (value === '') return '';

  const lowerValue = value.toLowerCase();
  if (lowerValue === 'true' || lowerValue === 'yes' || lowerValue === 'y') return true;
  if (lowerValue === 'false' || lowerValue === 'no' || lowerValue === 'n') return false;

  const cleanedValue = value.replace(/[,$¥€£]/g, '');
  const numberPattern = /^-?\d+\.?\d*([eE][+-]?\d+)?$/;
  if (numberPattern.test(cleanedValue)) {
    const num = parseFloat(cleanedValue);
    if (!isNaN(num)) return num;
  }

  if (value.endsWith('%')) {
    const num = parseFloat(value.slice(0, -1));
    if (!isNaN(num)) return num / 100;
  }

  const datePattern1 = /^\d{4}-\d{2}-\d{2}$/;
  if (datePattern1.test(value)) {
    const date = new Date(value);
    if (!isNaN(date.getTime())) {
      return date.toLocaleDateString('zh-CN');
    }
  }

  const datePattern2 = /^\d{1,2}\/\d{1,2}\/\d{4}$/;
  if (datePattern2.test(value)) {
    const date = new Date(value);
    if (!isNaN(date.getTime())) {
      return date.toLocaleDateString('zh-CN');
    }
  }

  const datePattern3 = /^\d{4}年\d{1,2}月\d{1,2}日$/;
  if (datePattern3.test(value)) {
    const matches = value.match(/(\d{4})年(\d{1,2})月(\d{1,2})日/);
    if (matches) {
      const date = new Date(parseInt(matches[1]), parseInt(matches[2]) - 1, parseInt(matches[3]));
      if (!isNaN(date.getTime())) {
        return date.toLocaleDateString('zh-CN');
      }
    }
  }

  return value;
}

// ==================== 图片插入 ====================

/**
 * 插入图片到工作表
 *
 * @description
 * 该函数将导入的图片插入到指定的工作表中。支持两种图片类型：
 * - 浮动图片 (FLOATING): 可以自由放置在工作表上方
 * - 单元格图片 (CELL): 嵌入到特定单元格中
 *
 * 图片通过 sheetId 直接定位工作表。
 *
 * @param univerAPI - Univer API 实例
 * @param images - 要插入的图片列表
 * @param options - 插入选项
 * @returns 插入结果，包含成功数、失败数和错误详情
 */
async function insertImagesAfterImport(
  univerAPI: unknown,
  images: ImportedImage[],
  options: {
    defaultType?: ImageType;
    continueOnError?: boolean;
    onProgress?: (current: number, total: number, image: ImportedImage) => void;
  } = {},
): Promise<{
  success: number;
  failed: number;
  errors: Array<{ image: ImportedImage; error: Error }>;
}> {
  const { defaultType = ImageType.FLOATING, continueOnError = true, onProgress } = options;

  const result = {
    success: 0,
    failed: 0,
    errors: [] as Array<{ image: ImportedImage; error: Error }>,
  };

  if (!images || images.length === 0) {
    return result;
  }

  const api = univerAPI as {
    getActiveWorkbook: () => {
      getSheetBySheetId: (id: string) => {
        getRange: (address: string) => {
          insertCellImageAsync: (source: string) => Promise<unknown>;
        } | null;
        newOverGridImage: () => {
          setSource: (source: string, type: number) => unknown;
          setColumn: (col: number) => unknown;
          setRow: (row: number) => unknown;
          setColumnOffset: (offset: number) => unknown;
          setRowOffset: (offset: number) => unknown;
          setWidth: (width: number) => unknown;
          setHeight: (height: number) => unknown;
          buildAsync: () => Promise<unknown>;
        };
        insertImages: (images: unknown[]) => void;
      } | null;
    } | null;
    Enum?: {
      ImageSourceType?: {
        BASE64?: number;
        URL?: number;
      };
    };
  };

  const fWorkbook = api.getActiveWorkbook();
  if (!fWorkbook) {
    throw new Error('无法获取当前工作簿');
  }

  for (let i = 0; i < images.length; i++) {
    const image = images[i];

    try {
      // 直接通过 sheetId 查找工作表
      const fWorksheet = fWorkbook.getSheetBySheetId(image.sheetId);
      if (!fWorksheet) {
        throw new Error(`找不到工作表: ${image.sheetName} (ID: ${image.sheetId})`);
      }

      const imageType = image.type || defaultType;

      if (imageType === ImageType.CELL) {
        const cellAddress = getCellAddress(image.position.row, image.position.column);
        const fRange = fWorksheet.getRange(cellAddress);

        if (!fRange) {
          throw new Error(`无法获取单元格范围: ${cellAddress}`);
        }

        const insertResult = await fRange.insertCellImageAsync(image.source);
        if (!insertResult) {
          throw new Error(`单元格图片插入失败: ${cellAddress}`);
        }
      } else {
        const { column, row, columnOffset, rowOffset } = image.position;
        const { width, height } = image.size;

        const isBase64 = image.source.startsWith('data:');
        const sourceType = isBase64
          ? api.Enum?.ImageSourceType?.BASE64 ?? 0
          : api.Enum?.ImageSourceType?.URL ?? 1;

        const imageBuilder = fWorksheet.newOverGridImage() as {
          setSource: (source: string, type: number) => unknown;
          setColumn: (col: number) => unknown;
          setRow: (row: number) => unknown;
          setColumnOffset: (offset: number) => unknown;
          setRowOffset: (offset: number) => unknown;
          setWidth: (width: number) => unknown;
          setHeight: (height: number) => unknown;
          buildAsync: () => Promise<unknown>;
        };
        imageBuilder.setSource(image.source, sourceType);
        imageBuilder.setColumn(column);
        imageBuilder.setRow(row);
        imageBuilder.setColumnOffset(columnOffset || 0);
        imageBuilder.setRowOffset(rowOffset || 0);
        imageBuilder.setWidth(width);
        imageBuilder.setHeight(height);

        const builtImage = await imageBuilder.buildAsync();
        if (!builtImage) {
          throw new Error(`构建浮动图片失败: (${column}, ${row})`);
        }

        fWorksheet.insertImages([builtImage]);
      }

      result.success++;
      onProgress?.(i + 1, images.length, image);
    } catch (error) {
      console.warn(`插入图片 ${i + 1} 失败:`, error);
      result.failed++;
      result.errors.push({
        image,
        error: error instanceof Error ? error : new Error(String(error)),
      });

      if (!continueOnError) {
        throw error;
      }
    }
  }

  return result;
}

// ==================== 统一导出接口 ====================

/**
 * 图片插入选项
 */
export interface ImageInsertOptions {
  /** 默认图片类型 */
  defaultType?: ImageType;
  /** 是否在插入失败时继续处理其他图片 */
  continueOnError?: boolean;
  /** 插入进度回调 */
  onProgress?: (current: number, total: number) => void;
}

/**
 * 文件导入选项
 */
export interface FileImportOptions {
  /** 是否包含图片（默认 true） */
  includeImages?: boolean;
}

/**
 * 文件导入结果
 *
 * @description
 * 包含工作簿数据、导入的图片列表以及插入图片的便捷方法。
 * 图片信息在导入时被提取，可以在工作簿创建后通过 `insertImages` 方法插入。
 */
export interface FileImportResult {
  /** 工作簿数据，可直接用于创建 Univer 工作簿 */
  workbookData: IWorkbookData;
  /** 导入的图片列表，包含图片的位置、尺寸等信息 */
  images: ImportedImage[];
  /**
   * 插入图片到工作表的便捷方法
   *
   * @description
   * 将导入的图片插入到已创建的工作簿中。该方法会通过 sheetId
   * 直接查找对应的工作表。
   *
   * @param univerAPI - Univer API 实例（通常是 FUniver 对象）
   * @param options - 可选的插入配置项
   * @returns Promise，解析为插入结果统计
   *
   * @example
   * ```ts
   * const result = await importFile(file);
   * if (result.images.length > 0) {
   *   const insertResult = await result.insertImages(univerAPI, {
   *     defaultType: ImageType.FLOATING,
   *     onProgress: (current, total) => console.log(`${current}/${total}`)
   *   });
   *   console.log(`成功: ${insertResult.success}, 失败: ${insertResult.failed}`);
   * }
   * ```
   */
  insertImages: (
    univerAPI: unknown,
    options?: ImageInsertOptions,
  ) => Promise<{
    success: number;
    failed: number;
    errors: Array<{ image: ImportedImage; error: Error }>;
  }>;
}

/**
 * 统一的文件导入接口
 *
 * @description
 * 支持导入 Excel (.xlsx, .xls) 和 CSV 文件，自动解析文件内容并转换为 Univer 工作簿格式。
 * 对于 Excel 文件，还会提取图片信息，可以在工作簿创建后通过返回的 `insertImages` 方法插入。
 *
 * @param file - 要导入的文件对象
 * @param options - 导入选项
 * @returns 导入结果，包含工作簿数据、图片列表和插入图片的方法
 *
 * @example
 * ```ts
 * // 基础用法：导入文件并创建工作簿
 * const result = await importFile(file);
 * const workbook = univerAPI.createWorkbook(result.workbookData);
 *
 * // 如果有图片，插入到工作簿中
 * if (result.images.length > 0) {
 *   const insertResult = await result.insertImages(univerAPI);
 *   console.log(`图片插入完成: 成功 ${insertResult.success}, 失败 ${insertResult.failed}`);
 * }
 * ```
 *
 * @example
 * ```ts
 * // 高级用法：自定义图片插入选项
 * const result = await importFile(file, { includeImages: true });
 *
 * if (result.images.length > 0) {
 *   await result.insertImages(univerAPI, {
 *     defaultType: ImageType.CELL,  // 使用单元格图片类型
 *     continueOnError: true,         // 失败时继续处理其他图片
 *     onProgress: (current, total) => {
 *       console.log(`正在插入图片: ${current}/${total}`);
 *     }
 *   });
 * }
 * ```
 *
 * @throws {Error} 当文件格式不支持时抛出错误
 */
export async function importFile(
  file: File,
  options: FileImportOptions = {},
): Promise<FileImportResult> {
  const { includeImages = true } = options;

  const fileExt = file.name.split('.').pop()?.toLowerCase();
  if (!fileExt || !['xlsx', 'xls', 'csv'].includes(fileExt)) {
    throw new Error(`不支持的文件格式: ${fileExt}`);
  }

  const fileType = fileExt as 'xlsx' | 'xls' | 'csv';
  const result = await handleFileImport(file, fileType, includeImages);
  const { workbookData, images } = result;

  return {
    workbookData,
    images,
    insertImages: async (
      univerAPI: unknown,
      insertOptions?: ImageInsertOptions,
    ) => {
      if (images.length === 0) {
        return { success: 0, failed: 0, errors: [] };
      }
      return insertImagesAfterImport(univerAPI, images, {
        defaultType: insertOptions?.defaultType ?? ImageType.FLOATING,
        continueOnError: insertOptions?.continueOnError ?? true,
        onProgress: insertOptions?.onProgress
          ? (current, total, _img) => insertOptions.onProgress!(current, total)
          : undefined,
      });
    },
  };
}

/* eslint-disable */
// @ts-nocheck

/**
 * Excel/CSV 文件导入工具 oumingliang 20226.1.30
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
 * ✅ 筛选器支持
 * ✅ 排序支持
 * ✅ 图表导入支持（柱状图、折线图、饼图等常见图表类型）
 * ✅ 透视表导入支持
 *
 */

import { LocaleType } from '@univerjs/core';
import type { IWorkbookData, IWorksheetData, ICellData } from '@univerjs/presets';
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
  /** 条件格式规则（按 sheetId 分组）*/
  conditionalFormats: Record<string, ImportedConditionalFormat[]>;
  /** 筛选器（按 sheetId 分组）*/
  filters: Record<string, ImportedFilter>;
  /** 排序（按 sheetId 分组）*/
  sorts: Record<string, ImportedSort>;
  /** 图表（按 sheetId 分组）*/
  charts: Record<string, ImportedChart[]>;
  /** 透视表列表 */
  pivotTables: ImportedPivotTable[];
}

/**
 * 导入的条件格式规则
 */
export interface ImportedConditionalFormat {
  /** 规则类型 */
  type: 'dataBar' | 'colorScale' | 'iconSet' | 'highlightCell' | 'other';
  /** 应用范围（A1 格式） */
  ranges: string[];
  /** 规则配置（根据类型不同） */
  config: any;
  /** 优先级 */
  priority?: number;
  /** 满足条件时停止 */
  stopIfTrue?: boolean;
}

/**
 * 导入的筛选器信息
 */
export interface ImportedFilter {
  /** 筛选范围（A1 格式，如 "A1:D14"） */
  range: string;
}

/**
 * 导入的排序信息
 */
export interface ImportedSort {
  /** 排序范围（A1 格式，如 "A1:D14"） */
  range: string;
  /** 排序条件列表 */
  conditions: Array<{
    /** 排序列索引（0-based，相对于范围起始列） */
    column: number;
    /** 是否升序，默认 true */
    ascending: boolean;
  }>;
}

/**
 * 导入的图表信息
 */
export interface ImportedChart {
  /** 图表唯一 ID */
  chartId: string;
  /** 所属工作表 ID */
  sheetId: string;
  /** 所属工作表名称 */
  sheetName: string;
  /** 图表类型 */
  chartType:
    | 'column'
    | 'bar'
    | 'line'
    | 'area'
    | 'pie'
    | 'doughnut'
    | 'scatter'
    | 'radar'
    | 'bubble'
    | 'combo'
    | 'stackedBar'
    | 'percentStackedBar'
    | 'stackedArea'
    | 'percentStackedArea'
    | 'wordCloud'
    | 'funnel'
    | 'relationship'
    | 'waterfall'
    | 'treemap'
    | 'sankey'
    | 'heatmap'
    | 'boxPlot'
    | 'unknown';
  /** 数据源范围（A1 格式，如 "A1:D6"） */
  dataRange?: string;
  /** 数据源所在工作表名称（如果与图表所在工作表不同） */
  dataSheetName?: string;
  /** 图表位置 */
  position: {
    /** 起始行（0-based） */
    row: number;
    /** 起始列（0-based） */
    column: number;
    /** 行偏移（像素） */
    rowOffset: number;
    /** 列偏移（像素） */
    columnOffset: number;
  };
  /** 图表尺寸 */
  size: {
    /** 宽度（像素） */
    width: number;
    /** 高度（像素） */
    height: number;
  };
  /** 图表标题 */
  title?: string;
  /** 原始图表数据（用于高级配置） */
  rawData?: any;
}

/**
 * 导入的透视表信息
 */
export interface ImportedPivotTable {
  /** 透视表唯一 ID */
  pivotTableId: string;
  /** 所属工作表 ID */
  sheetId: string;
  /** 所属工作表名称 */
  sheetName: string;
  /** 数据源范围 */
  sourceRange: {
    /** 数据源工作表名称 */
    sheetName: string;
    /** 数据源工作表 ID（如果可用） */
    sheetId?: string;
    /** 起始行（0-based） */
    startRow: number;
    /** 起始列（0-based） */
    startColumn: number;
    /** 结束行（0-based） */
    endRow: number;
    /** 结束列（0-based） */
    endColumn: number;
  };
  /** 透视表锚点位置（左上角单元格） */
  anchorCell: {
    /** 行（0-based） */
    row: number;
    /** 列（0-based） */
    col: number;
  };
  /** 透视表占据区域（用于导入时跳过该区域的单元格数据） */
  occupiedRange?: {
    /** 起始行（0-based） */
    startRow: number;
    /** 起始列（0-based） */
    startColumn: number;
    /** 结束行（0-based） */
    endRow: number;
    /** 结束列（0-based） */
    endColumn: number;
  };
  /** 透视表字段配置 */
  fields: {
    /** 行字段索引列表 */
    rowFields: number[];
    /** 列字段索引列表 */
    colFields: number[];
    /** 值字段索引列表 */
    valueFields: number[];
    /** 筛选字段索引列表 */
    filterFields: number[];
  };
  /** 透视表名称 */
  name?: string;
}

/**
 * 解析 Excel 范围引用（如 "Sheet1!$A$1:$D$6"）
 * 返回 { sheetName, startCol, startRow, endCol, endRow }
 */
function parseExcelRange(range: string): {
  sheetName: string;
  startCol: number;
  startRow: number;
  endCol: number;
  endRow: number;
} | null {
  // 匹配格式: Sheet1!$A$1:$D$6 或 Sheet1!$A$1
  const match = range.match(/^(.+?)!\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$/);
  if (!match) return null;

  const sheetName = match[1].replace(/^'|'$/g, ''); // 移除引号
  const startColStr = match[2];
  const startRow = parseInt(match[3], 10);
  const endColStr = match[4] || startColStr;
  const endRow = match[5] ? parseInt(match[5], 10) : startRow;

  // 将列字母转换为数字（A=1, B=2, ...）
  const colToNum = (col: string): number => {
    let num = 0;
    for (let i = 0; i < col.length; i++) {
      num = num * 26 + (col.charCodeAt(i) - 64);
    }
    return num;
  };

  return {
    sheetName,
    startCol: colToNum(startColStr),
    startRow,
    endCol: colToNum(endColStr),
    endRow,
  };
}

/**
 * 将列数字转换为列字母（1=A, 2=B, ...）
 */
function numToCol(num: number): string {
  let col = '';
  while (num > 0) {
    const remainder = (num - 1) % 26;
    col = String.fromCharCode(65 + remainder) + col;
    num = Math.floor((num - 1) / 26);
  }
  return col;
}

/**
 * 合并多个 Excel 范围引用为一个最小包围矩形
 * @param ranges 范围数组，如 ["Sheet1!$A$1:$A$6", "Sheet1!$B$1:$D$6"]
 * @returns 合并后的范围，如 "A1:D6"（不包含工作表名称，因为 Univer 会自动处理）
 */
function mergeChartDataRanges(ranges: string[]): string {
  if (ranges.length === 0) return '';

  let minCol = Infinity;
  let minRow = Infinity;
  let maxCol = -Infinity;
  let maxRow = -Infinity;
  let sheetName = '';

  for (const range of ranges) {
    const parsed = parseExcelRange(range);
    if (!parsed) continue;

    if (!sheetName) sheetName = parsed.sheetName;

    minCol = Math.min(minCol, parsed.startCol, parsed.endCol);
    minRow = Math.min(minRow, parsed.startRow, parsed.endRow);
    maxCol = Math.max(maxCol, parsed.startCol, parsed.endCol);
    maxRow = Math.max(maxRow, parsed.startRow, parsed.endRow);
  }

  if (minCol === Infinity || minRow === Infinity) {
    return ranges[0] || '';
  }

  // 返回格式：A1:D6（Univer addRange 需要这种格式）
  return `${numToCol(minCol)}${minRow}:${numToCol(maxCol)}${maxRow}`;
}

/**
 * 从 xlsx 文件中解析图表信息
 * 由于 ExcelJS 不支持读取图表，需要直接解析 xlsx（zip 格式）
 *
 * xlsx 文件结构：
 * - xl/drawings/drawing1.xml - 包含图表锚点位置
 * - xl/charts/chart1.xml - 包含图表类型和数据范围
 * - xl/drawings/_rels/drawing1.xml.rels - 包含 drawing 到 chart 的关系映射
 * - xl/_rels/workbook.xml.rels - 包含 sheet 到 drawing 的关系映射
 */
async function parseChartsFromXlsx(
  arrayBuffer: ArrayBuffer,
  sheetNameToIdMap: Map<string, string>,
): Promise<Record<string, ImportedChart[]>> {
  const charts: Record<string, ImportedChart[]> = {};

  try {
    // 使用浏览器原生的 JSZip（通过动态导入）
    // ExcelJS 内部依赖 JSZip，我们可以直接使用
    const JSZip = (await import('jszip')).default;
    const zip = await JSZip.loadAsync(arrayBuffer);

    // 解析工作表和 drawing 的关系
    const sheetDrawingMap = new Map<string, string>(); // sheetName -> drawingPath

    // 读取 xl/workbook.xml 获取 sheet 列表
    const workbookXmlFile = zip.file('xl/workbook.xml');
    const workbookRelsFile = zip.file('xl/_rels/workbook.xml.rels');

    if (!workbookXmlFile || !workbookRelsFile) {
      console.log('[fileImport] 无法找到 workbook.xml 或其关系文件');
      return charts;
    }

    const workbookXml = await workbookXmlFile.async('string');
    const workbookRels = await workbookRelsFile.async('string');

    // 解析 sheet 名称和 rId 的映射
    const sheetRIdMap = new Map<string, string>(); // sheetName -> rId
    const sheetMatches = workbookXml.matchAll(/<sheet[^>]*name="([^"]*)"[^>]*r:id="([^"]*)"/g);
    for (const match of sheetMatches) {
      sheetRIdMap.set(match[1], match[2]);
    }

    // 解析 rId 和 sheet 文件路径的映射
    const rIdToPathMap = new Map<string, string>(); // rId -> sheetPath
    const relMatches = workbookRels.matchAll(
      /<Relationship[^>]*Id="([^"]*)"[^>]*Target="([^"]*)"/g,
    );
    for (const match of relMatches) {
      rIdToPathMap.set(match[1], match[2]);
    }

    // 遍历每个 sheet 的关系文件，找到 drawing
    for (const [sheetName, rId] of sheetRIdMap) {
      const sheetPath = rIdToPathMap.get(rId);
      if (!sheetPath) continue;

      // 获取 sheet 的关系文件
      const sheetRelsPath = sheetPath
        .replace('worksheets/', 'worksheets/_rels/')
        .replace('.xml', '.xml.rels');
      const sheetRelsFile = zip.file(`xl/${sheetRelsPath}`);

      if (!sheetRelsFile) continue;

      const sheetRels = await sheetRelsFile.async('string');

      // 查找 drawing 关系
      const drawingMatch = sheetRels.match(
        /<Relationship[^>]*Type="[^"]*drawing"[^>]*Target="([^"]*)"/,
      );
      if (drawingMatch) {
        const drawingPath = drawingMatch[1].replace('../', '');
        sheetDrawingMap.set(sheetName, drawingPath);
      }
    }

    // 遍历每个 drawing 文件，解析图表信息
    for (const [sheetName, drawingPath] of sheetDrawingMap) {
      const sheetId = sheetNameToIdMap.get(sheetName);
      if (!sheetId) {
        console.log(`[fileImport] 未找到 sheet "${sheetName}" 的 ID`);
        continue;
      }

      // 读取 drawing 文件
      const drawingFile = zip.file(`xl/${drawingPath}`);
      if (!drawingFile) {
        console.log(`[fileImport] 未找到 drawing 文件: xl/${drawingPath}`);
        continue;
      }

      const drawingXml = await drawingFile.async('string');

      // 读取 drawing 的关系文件
      const drawingRelsPath = drawingPath
        .replace('drawings/', 'drawings/_rels/')
        .replace('.xml', '.xml.rels');
      const drawingRelsFile = zip.file(`xl/${drawingRelsPath}`);

      if (!drawingRelsFile) {
        console.log(`[fileImport] 未找到 drawing 关系文件: xl/${drawingRelsPath}`);
        continue;
      }

      const drawingRels = await drawingRelsFile.async('string');

      // 解析图表 rId 到 chart 路径的映射
      const chartRIdMap = new Map<string, string>();
      const chartRelMatches = drawingRels.matchAll(
        /<Relationship[^>]*Id="([^"]*)"[^>]*Target="([^"]*)"/g,
      );
      for (const match of chartRelMatches) {
        if (match[2].includes('chart')) {
          chartRIdMap.set(match[1], match[2].replace('../', ''));
        }
      }

      // 解析 drawing 中的图表锚点
      // 匹配 twoCellAnchor 或 oneCellAnchor 中的图表引用
      const anchorRegex =
        /<(?:xdr:)?twoCellAnchor[^>]*>([\s\S]*?)<\/(?:xdr:)?twoCellAnchor>|<(?:xdr:)?oneCellAnchor[^>]*>([\s\S]*?)<\/(?:xdr:)?oneCellAnchor>/g;
      let anchorMatch;
      let chartIndex = 0;

      while ((anchorMatch = anchorRegex.exec(drawingXml)) !== null) {
        const anchorContent = anchorMatch[1] || anchorMatch[2];

        // 检查是否是图表（包含 graphicFrame 和 chart 引用）
        const chartRefMatch = anchorContent.match(/<c:chart[^>]*r:id="([^"]*)"/);
        if (!chartRefMatch) continue;

        const chartRId = chartRefMatch[1];
        const chartPath = chartRIdMap.get(chartRId);
        if (!chartPath) continue;

        // 解析位置信息
        const fromColMatch = anchorContent.match(
          /<(?:xdr:)?from>[\s\S]*?<(?:xdr:)?col>(\d+)<\/(?:xdr:)?col>/,
        );
        const fromRowMatch = anchorContent.match(
          /<(?:xdr:)?from>[\s\S]*?<(?:xdr:)?row>(\d+)<\/(?:xdr:)?row>/,
        );
        const fromColOffMatch = anchorContent.match(
          /<(?:xdr:)?from>[\s\S]*?<(?:xdr:)?colOff>(\d+)<\/(?:xdr:)?colOff>/,
        );
        const fromRowOffMatch = anchorContent.match(
          /<(?:xdr:)?from>[\s\S]*?<(?:xdr:)?rowOff>(\d+)<\/(?:xdr:)?rowOff>/,
        );

        const position = {
          column: fromColMatch ? parseInt(fromColMatch[1], 10) : 0,
          row: fromRowMatch ? parseInt(fromRowMatch[1], 10) : 0,
          columnOffset: fromColOffMatch ? Math.round(parseInt(fromColOffMatch[1], 10) / 9525) : 0, // EMUs to pixels
          rowOffset: fromRowOffMatch ? Math.round(parseInt(fromRowOffMatch[1], 10) / 9525) : 0,
        };

        // 解析尺寸信息（从 to 锚点计算或使用默认值）
        let width = 600;
        let height = 400;

        const extMatch = anchorContent.match(/<(?:xdr:)?ext[^>]*cx="(\d+)"[^>]*cy="(\d+)"/);
        if (extMatch) {
          width = Math.round(parseInt(extMatch[1], 10) / 9525); // EMUs to pixels
          height = Math.round(parseInt(extMatch[2], 10) / 9525);
        }

        // 读取图表文件，获取图表类型和数据范围
        const chartFile = zip.file(`xl/${chartPath}`);
        let chartType: ImportedChart['chartType'] = 'column';
        let dataRange = '';
        let title = '';
        let isPivotChart = false; // 是否是透视图

        if (chartFile) {
          const chartXml = await chartFile.async('string');

          // 检查是否是透视图（pivotSource 元素表示这是透视图）
          // 透视图应该跳过，让透视表插件处理
          if (chartXml.includes('<c:pivotSource') || chartXml.includes('<c15:pivotSource')) {
            isPivotChart = true;
            continue;
          }

          // 解析图表类型
          if (chartXml.includes('<c:barChart')) {
            // 检查是否是水平条形图
            const barDirMatch = chartXml.match(/<c:barDir[^>]*val="([^"]*)"/);
            const groupingMatch = chartXml.match(/<c:grouping[^>]*val="([^"]*)"/);
            const grouping = groupingMatch ? groupingMatch[1] : 'clustered';

            if (barDirMatch && barDirMatch[1] === 'bar') {
              // 水平条形图
              if (grouping === 'stacked') {
                chartType = 'stackedBar';
              } else if (grouping === 'percentStacked') {
                chartType = 'percentStackedBar';
              } else {
                chartType = 'bar';
              }
            } else {
              // 垂直柱状图
              if (grouping === 'stacked') {
                chartType = 'column'; // 堆叠柱状图，暂时使用 column，后续可能需要扩展
              } else if (grouping === 'percentStacked') {
                chartType = 'column'; // 百分比堆叠柱状图，暂时使用 column
              } else {
                chartType = 'column';
              }
            }
          } else if (chartXml.includes('<c:lineChart')) {
            chartType = 'line';
          } else if (chartXml.includes('<c:pieChart')) {
            chartType = 'pie';
          } else if (chartXml.includes('<c:doughnutChart')) {
            chartType = 'doughnut';
          } else if (chartXml.includes('<c:areaChart')) {
            // 检查是否是堆叠面积图
            const groupingMatch = chartXml.match(/<c:grouping[^>]*val="([^"]*)"/);
            const grouping = groupingMatch ? groupingMatch[1] : 'standard';

            if (grouping === 'stacked') {
              chartType = 'stackedArea';
            } else if (grouping === 'percentStacked') {
              chartType = 'percentStackedArea';
            } else {
              chartType = 'area';
            }
          } else if (chartXml.includes('<c:scatterChart')) {
            chartType = 'scatter';
          } else if (chartXml.includes('<c:radarChart')) {
            chartType = 'radar';
          } else if (chartXml.includes('<c:bubbleChart')) {
            chartType = 'bubble';
          } else if (chartXml.includes('<c:ofPieChart')) {
            // 复合饼图（可能包含其他变体）
            chartType = 'pie';
          } else if (chartXml.includes('<c:surfaceChart')) {
            // 曲面图，映射到散点图
            chartType = 'scatter';
          } else if (chartXml.includes('<c:stockChart')) {
            // 股价图，映射到组合图
            chartType = 'combo';
          }

          // 解析数据范围 - 需要获取完整的数据区域
          // Excel 图表的数据存储在多个位置：
          // - <c:cat>/<c:strRef>/<c:f> 或 <c:cat>/<c:numRef>/<c:f> - 分类轴数据
          // - <c:val>/<c:numRef>/<c:f> - 数值数据
          // 我们需要收集所有范围并计算最小包围矩形
          const allRanges: string[] = [];
          const rangeMatches = chartXml.matchAll(/<c:f>([^<]+)<\/c:f>/g);
          for (const match of rangeMatches) {
            const range = match[1];
            // 过滤掉非范围引用（如纯文本）
            if (range.includes('!') && range.includes('$')) {
              allRanges.push(range);
            }
          }

          // 合并所有范围为一个最小包围矩形
          if (allRanges.length > 0) {
            dataRange = mergeChartDataRanges(allRanges);
          }

          // 解析标题
          const titleMatch = chartXml.match(/<c:title>[\s\S]*?<c:t>([^<]+)<\/c:t>/);
          if (titleMatch) {
            title = titleMatch[1];
          }
        }

        const chart: ImportedChart = {
          chartId: `chart-${nanoid()}`,
          sheetId,
          sheetName,
          chartType,
          dataRange,
          position,
          size: { width, height },
          title,
        };

        if (!charts[sheetId]) {
          charts[sheetId] = [];
        }
        charts[sheetId].push(chart);
        chartIndex++;
      }
    }
  } catch (error) {
    console.error('[fileImport] 解析图表失败:', error);
  }

  return charts;
}

/**
 * 从 xlsx 文件中解析透视表信息
 * xlsx 透视表结构：
 * - xl/pivotTables/pivotTable*.xml - 透视表定义
 * - xl/pivotCache/pivotCacheDefinition*.xml - 数据源定义
 * - xl/worksheets/sheet*.xml - 包含透视表引用
 */
async function parsePivotTablesFromXlsx(
  arrayBuffer: ArrayBuffer,
  sheetNameToIdMap: Map<string, string>,
): Promise<ImportedPivotTable[]> {
  const pivotTables: ImportedPivotTable[] = [];

  try {
    const JSZip = (await import('jszip')).default;
    const zip = await JSZip.loadAsync(arrayBuffer);

    // 读取 xl/workbook.xml 获取 sheet 列表
    const workbookXmlFile = zip.file('xl/workbook.xml');
    const workbookRelsFile = zip.file('xl/_rels/workbook.xml.rels');

    if (!workbookXmlFile || !workbookRelsFile) {
      console.log('[fileImport] 无法找到 workbook.xml 或其关系文件');
      return pivotTables;
    }

    const workbookXml = await workbookXmlFile.async('string');
    const workbookRels = await workbookRelsFile.async('string');

    // 解析 sheet 名称和 rId 的映射
    const sheetRIdMap = new Map<string, string>(); // sheetName -> rId
    const sheetMatches = workbookXml.matchAll(/<sheet[^>]*name="([^"]*)"[^>]*r:id="([^"]*)"/g);
    for (const match of sheetMatches) {
      sheetRIdMap.set(match[1], match[2]);
    }

    // 解析 rId 和 sheet 文件路径的映射
    const rIdToPathMap = new Map<string, string>(); // rId -> sheetPath
    const relMatches = workbookRels.matchAll(
      /<Relationship[^>]*Id="([^"]*)"[^>]*Target="([^"]*)"/g,
    );
    for (const match of relMatches) {
      rIdToPathMap.set(match[1], match[2]);
    }

    // 解析 workbook.xml 中的 pivotCaches，获取 cacheId -> rId 映射
    // 格式: <pivotCache cacheId="6" r:id="rId8"/>
    const cacheIdToRIdMap = new Map<string, string>(); // cacheId -> rId
    const pivotCacheMatches = workbookXml.matchAll(
      /<pivotCache[^>]*cacheId="(\d+)"[^>]*r:id="([^"]*)"|<pivotCache[^>]*r:id="([^"]*)"[^>]*cacheId="(\d+)"/g,
    );
    for (const match of pivotCacheMatches) {
      const cacheId = match[1] || match[4];
      const rId = match[2] || match[3];
      if (cacheId && rId) {
        cacheIdToRIdMap.set(cacheId, rId);
      }
    }

    // 构建 pivotCacheId -> cacheDefinition 的映射
    const pivotCacheMap = new Map<
      string,
      {
        sourceSheetName: string;
        sourceRange: { startRow: number; startColumn: number; endRow: number; endColumn: number };
      }
    >();

    // 遍历 cacheId -> rId 映射，读取对应的 pivotCacheDefinition 文件
    for (const [cacheId, rId] of cacheIdToRIdMap) {
      const cachePath = rIdToPathMap.get(rId);
      if (!cachePath) {
        console.log(`[fileImport] 未找到 cacheId=${cacheId} 对应的缓存文件路径, rId=${rId}`);
        continue;
      }

      // 路径可能是相对路径，需要处理
      const fullCachePath = cachePath.startsWith('/') ? `xl${cachePath}` : `xl/${cachePath}`;
      const cacheFile = zip.file(fullCachePath);
      if (!cacheFile) {
        console.log(`[fileImport] 未找到缓存文件: ${fullCachePath}`);
        continue;
      }

      const cacheXml = await cacheFile.async('string');

      // 解析数据源范围，例如: <cacheSource type="worksheet"><worksheetSource ref="A1:D6" sheet="Sheet1"/></cacheSource>
      const worksheetSourceMatch = cacheXml.match(
        /<worksheetSource[^>]*ref="([^"]*)"[^>]*sheet="([^"]*)"|<worksheetSource[^>]*sheet="([^"]*)"[^>]*ref="([^"]*)"/,
      );

      if (worksheetSourceMatch) {
        const ref = worksheetSourceMatch[1] || worksheetSourceMatch[4];
        const sheet = worksheetSourceMatch[2] || worksheetSourceMatch[3];

        if (ref && sheet) {
          const rangeMatch = ref.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
          if (rangeMatch) {
            const startColumn = colToNum(rangeMatch[1]);
            const startRow = parseInt(rangeMatch[2], 10) - 1;
            const endColumn = colToNum(rangeMatch[3]);
            const endRow = parseInt(rangeMatch[4], 10) - 1;

            pivotCacheMap.set(cacheId, {
              sourceSheetName: sheet,
              sourceRange: { startRow, startColumn, endRow, endColumn },
            });
          }
        }
      }
    }

    // 遍历每个 sheet 的关系文件，找到透视表
    for (const [sheetName, rId] of sheetRIdMap) {
      const sheetPath = rIdToPathMap.get(rId);
      if (!sheetPath) continue;

      const sheetId = sheetNameToIdMap.get(sheetName);
      if (!sheetId) continue;

      // 获取 sheet 的关系文件
      const sheetRelsPath = sheetPath
        .replace('worksheets/', 'worksheets/_rels/')
        .replace('.xml', '.xml.rels');
      const sheetRelsFile = zip.file(`xl/${sheetRelsPath}`);

      if (!sheetRelsFile) continue;

      const sheetRels = await sheetRelsFile.async('string');

      // 查找透视表关系
      const pivotTableRels = [
        ...sheetRels.matchAll(
          /<Relationship[^>]*Id="([^"]*)"[^>]*Type="[^"]*pivotTable"[^>]*Target="([^"]*)"/g,
        ),
      ];

      for (const pivotRel of pivotTableRels) {
        const pivotTablePath = pivotRel[2].replace('../', '');
        const pivotTableFile = zip.file(`xl/${pivotTablePath}`);

        if (!pivotTableFile) continue;

        const pivotTableXml = await pivotTableFile.async('string');

        // 解析透视表位置，例如: <location ref="A8:C12" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"/>
        const locationMatch = pivotTableXml.match(/<location[^>]*ref="([^"]*)"/);
        let anchorRow = 0;
        let anchorCol = 0;
        let occupiedRange: ImportedPivotTable['occupiedRange'] = undefined;

        if (locationMatch) {
          const refRange = locationMatch[1];
          // 解析完整范围，如 A8:C12
          const fullRangeMatch = refRange.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
          if (fullRangeMatch) {
            anchorCol = colToNum(fullRangeMatch[1]);
            anchorRow = parseInt(fullRangeMatch[2], 10) - 1;
            occupiedRange = {
              startRow: parseInt(fullRangeMatch[2], 10) - 1,
              startColumn: colToNum(fullRangeMatch[1]),
              endRow: parseInt(fullRangeMatch[4], 10) - 1,
              endColumn: colToNum(fullRangeMatch[3]),
            };
          } else {
            // 单个单元格
            const cellMatch = refRange.match(/([A-Z]+)(\d+)/);
            if (cellMatch) {
              anchorCol = colToNum(cellMatch[1]);
              anchorRow = parseInt(cellMatch[2], 10) - 1;
            }
          }
        }

        // 解析透视表名称
        const nameMatch = pivotTableXml.match(/<pivotTableDefinition[^>]*name="([^"]*)"/);
        const name = nameMatch ? nameMatch[1] : undefined;

        // 解析 cacheId
        const cacheIdMatch = pivotTableXml.match(/<pivotTableDefinition[^>]*cacheId="(\d+)"/);
        const cacheId = cacheIdMatch ? cacheIdMatch[1] : '1';

        // 获取数据源信息
        const cacheInfo = pivotCacheMap.get(cacheId);
        if (!cacheInfo) {
          console.log(`[fileImport] 未找到透视表的数据源缓存: cacheId=${cacheId}`);
          continue;
        }

        // 解析字段配置
        const rowFields: number[] = [];
        const colFields: number[] = [];
        const valueFields: number[] = [];
        const filterFields: number[] = [];

        // 解析行字段: <rowFields count="1"><field x="0"/></rowFields>
        const rowFieldsMatch = pivotTableXml.match(/<rowFields[^>]*>([\s\S]*?)<\/rowFields>/);
        if (rowFieldsMatch) {
          const fieldMatches = rowFieldsMatch[1].matchAll(/<field[^>]*x="(-?\d+)"/g);
          for (const match of fieldMatches) {
            const fieldIndex = parseInt(match[1], 10);
            if (fieldIndex >= 0) {
              rowFields.push(fieldIndex);
            }
          }
        }

        // 解析列字段: <colFields count="1"><field x="1"/></colFields>
        const colFieldsMatch = pivotTableXml.match(/<colFields[^>]*>([\s\S]*?)<\/colFields>/);
        if (colFieldsMatch) {
          const fieldMatches = colFieldsMatch[1].matchAll(/<field[^>]*x="(-?\d+)"/g);
          for (const match of fieldMatches) {
            const fieldIndex = parseInt(match[1], 10);
            if (fieldIndex >= 0) {
              colFields.push(fieldIndex);
            }
          }
        }

        // 解析值字段: <dataFields count="1"><dataField name="Sum of Sales" fld="2"/></dataFields>
        const dataFieldsMatch = pivotTableXml.match(/<dataFields[^>]*>([\s\S]*?)<\/dataFields>/);
        if (dataFieldsMatch) {
          const fieldMatches = dataFieldsMatch[1].matchAll(/<dataField[^>]*fld="(\d+)"/g);
          for (const match of fieldMatches) {
            valueFields.push(parseInt(match[1], 10));
          }
        }

        // 解析筛选字段: <pageFields count="1"><pageField fld="3"/></pageFields>
        const pageFieldsMatch = pivotTableXml.match(/<pageFields[^>]*>([\s\S]*?)<\/pageFields>/);
        if (pageFieldsMatch) {
          const fieldMatches = pageFieldsMatch[1].matchAll(/<pageField[^>]*fld="(\d+)"/g);
          for (const match of fieldMatches) {
            filterFields.push(parseInt(match[1], 10));
          }
        }

        // 筛选字段会在透视表数据区域上方显示，每个筛选字段占用 2 行（名称行 + 值行）
        // 需要调整 anchorRow 和 occupiedRange 来包含筛选区域
        const filterRowCount = filterFields.length * 2;
        if (filterRowCount > 0) {
          // 向上调整 anchorRow 以包含筛选区域
          anchorRow = Math.max(0, anchorRow - filterRowCount);
          // 同时调整 occupiedRange 的起始行
          if (occupiedRange) {
            occupiedRange.startRow = Math.max(0, occupiedRange.startRow - filterRowCount);
          }
        }

        // 通过 sheetName 查找数据源 sheetId
        const sourceSheetId = sheetNameToIdMap.get(cacheInfo.sourceSheetName);

        const pivotTable: ImportedPivotTable = {
          pivotTableId: `pivot-${nanoid()}`,
          sheetId,
          sheetName,
          sourceRange: {
            sheetName: cacheInfo.sourceSheetName,
            sheetId: sourceSheetId, // 保存数据源 sheetId
            ...cacheInfo.sourceRange,
          },
          anchorCell: {
            row: anchorRow,
            col: anchorCol,
          },
          occupiedRange,
          fields: {
            rowFields,
            colFields,
            valueFields,
            filterFields,
          },
          name,
        };

        pivotTables.push(pivotTable);
      }
    }
  } catch (error) {
    console.error('[fileImport] 解析透视表失败:', error);
  }

  return pivotTables;
}

/**
 * 从 xlsx 文件中直接解析排序信息
 * Excel 的排序信息存储在 xl/worksheets/sheet*.xml 的 <sortState> 元素中
 * 格式: <sortState ref="A1:D10"><sortCondition ref="A1:A10" descending="1"/></sortState>
 */
async function parseSortsFromXlsx(
  arrayBuffer: ArrayBuffer,
  sheetIndexToIdMap: string[],
): Promise<Record<string, ImportedSort>> {
  const sorts: Record<string, ImportedSort> = {};

  try {
    // 动态导入 jszip
    const JSZip = (await import('jszip')).default;
    const zip = await JSZip.loadAsync(arrayBuffer);

    // 遍历所有工作表文件
    for (const [path, file] of Object.entries(zip.files)) {
      // 匹配 xl/worksheets/sheet*.xml
      const sheetMatch = path.match(/^xl\/worksheets\/sheet(\d+)\.xml$/);
      if (!sheetMatch || file.dir) continue;

      const sheetIndex = parseInt(sheetMatch[1], 10) - 1; // 转为 0-based
      const sheetId = sheetIndexToIdMap[sheetIndex];
      if (!sheetId) continue;

      const sheetXml = await file.async('string');

      // 解析 <sortState ref="A1:D10">...</sortState>
      const sortStateMatch = sheetXml.match(
        /<sortState[^>]*ref="([^"]+)"[^>]*>([\s\S]*?)<\/sortState>/,
      );
      if (!sortStateMatch) continue;

      const sortRange = sortStateMatch[1];
      const sortStateContent = sortStateMatch[2];

      // 解析排序范围的起始列（用于计算相对列索引）
      const rangeStartCol = sortRange.match(/^([A-Z]+)/)?.[1] || 'A';
      const startColIndex = colToNum(rangeStartCol);

      // 解析排序条件 <sortCondition ref="A1:A10" descending="1"/>
      const conditions: Array<{ column: number; ascending: boolean }> = [];
      const conditionRegex =
        /<sortCondition[^>]*ref="([^"]+)"[^>]*(?:descending="(\d+)")?[^>]*\/?>/g;

      let condMatch;
      while ((condMatch = conditionRegex.exec(sortStateContent)) !== null) {
        const condRef = condMatch[1];
        const descending = condMatch[2] === '1';

        // 从 ref 中提取列字母（如 "A1:A10" -> "A"）
        const colMatch = condRef.match(/^([A-Z]+)/);
        if (colMatch) {
          const colIndex = colToNum(colMatch[1]);
          // 计算相对于排序范围起始列的偏移
          const relativeColumn = colIndex - startColIndex;
          conditions.push({
            column: relativeColumn,
            ascending: !descending,
          });
        }
      }

      if (conditions.length > 0) {
        sorts[sheetId] = {
          range: sortRange,
          conditions,
        };
      }
    }
  } catch (error) {
    console.error('[fileImport] 解析排序失败:', error);
  }

  return sorts;
}

/**
 * 列字母转数字（0-based）
 * A -> 0, B -> 1, Z -> 25, AA -> 26
 */
function colToNum(col: string): number {
  let num = 0;
  for (let i = 0; i < col.length; i++) {
    num = num * 26 + (col.charCodeAt(i) - 64);
  }
  return num - 1;
}

/**
 * 转义工作表名称中的特殊字符
 * 处理类似 >>> 等特殊字符
 */
function escapeSheetName(name: string): string {
  if (!name) return 'Sheet1';
  // 保留原始名称，Univer 应该能处理大部分特殊字符
  // 如果需要特殊处理，可以在这里添加
  return name;
}

/**
 * 反转义工作表名称
 */
function unescapeSheetName(name: string): string {
  return name;
}

/**
 * 将 ArrayBuffer 转换为 Base64 字符串（浏览器兼容）
 */
function arrayBufferToBase64(buffer: ArrayBuffer): string {
  const bytes = new Uint8Array(buffer);
  let binary = '';
  const chunkSize = 8192; // 分块处理，避免栈溢出
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
 * @param numFmt Excel数字格式字符串
 * @returns 是否为日期格式
 */
function isDateFormat(numFmt: string): boolean {
  if (!numFmt || numFmt === 'General') return false;

  // 常见的日期格式关键词
  const dateKeywords = [
    'yyyy',
    'yy',
    'mm',
    'dd',
    'm/d',
    'd/m',
    'h:mm',
    'hh:mm',
    'ss',
    '年',
    '月',
    '日',
    'AM/PM',
    'am/pm',
    '上午',
    '下午',
    '[$-', // Excel本地化日期格式前缀
  ];

  // 检查是否包含日期关键词
  const lowerFmt = numFmt.toLowerCase();
  const hasDateKeyword = dateKeywords.some((keyword) => lowerFmt.includes(keyword.toLowerCase()));

  // 常见的内置日期格式编号（Excel使用）
  const dateFormatCodes = [
    14,
    15,
    16,
    17,
    18,
    19,
    20,
    21,
    22, // 日期格式
    45,
    46,
    47, // 时间格式
    176,
    177,
    178,
    179,
    180,
    181,
    182, // 本地化日期格式
  ];

  // 如果格式是数字编号，检查是否为日期格式编号
  const formatCode = parseInt(numFmt, 10);
  if (!isNaN(formatCode) && dateFormatCodes.includes(formatCode)) {
    return true;
  }

  return hasDateKeyword;
}

/**
 * 处理常见文件类型（Excel/CSV）
 * @param file 文件对象
 * @param type 文件类型
 * @param includeImages 是否解析图片（默认 true）
 * @returns 导入结果（包含工作簿数据和图片信息）
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
      result = {
        workbookData,
        images: [],
        conditionalFormats: {},
        filters: {},
        sorts: {},
        charts: {},
        pivotTables: [],
      };
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
 * 支持 .xlsx 和 .xls 格式
 * 注意：对于大文件（>10万行），此函数会消耗大量内存
 * @param file 文件对象
 * @param includeImages 是否解析图片（默认 true）
 */
async function importExcelWithImages(
  file: File,
  includeImages: boolean = true,
): Promise<ImportResult> {
  const fileSize = file.size;
  const fileName = file.name;
  const fileExt = fileName.split('.').pop()?.toLowerCase();

  // 大文件警告
  if (fileSize > 10 * 1024 * 1024) {
    console.warn(
      `⚠️ 正在导入大文件 (${(fileSize / 1024 / 1024).toFixed(1)}MB)，可能需要较长时间...`,
    );
  }

  const workbook = new ExcelJS.Workbook();
  const arrayBuffer = await file.arrayBuffer();

  try {
    // 根据文件扩展名选择加载方式
    if (fileExt === 'xls') {
      await workbook.xlsx.load(arrayBuffer);
    } else {
      // .xlsx 文件的标准加载
      await workbook.xlsx.load(arrayBuffer);
    }
  } catch (error) {
    console.error('Excel 文件加载失败:', error);
    throw new Error(`无法加载 Excel 文件: ${error.message}. 如果是 .xls 格式，请先转换为 .xlsx`);
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

  // 全局图片收集器 - 用于构建 SHEET_DRAWING_PLUGIN 资源
  const allDrawings: Record<string, any> = {};
  // 收集所有图片信息（用于返回给调用者）
  const allImages: ImportedImage[] = [];
  // 全局条件格式收集器 - 用于构建 SHEET_CONDITIONAL_FORMATTING_PLUGIN 资源
  // Univer 使用 { [subUnitId]: IConditionFormattingRule[] } 结构
  const allConditionalFormats: Record<string, any[]> = {};
  // 全局筛选器收集器 - 按 sheetId 存储筛选范围
  const allFilters: Record<string, ImportedFilter> = {};
  // 全局排序收集器 - 按 sheetId 存储排序信息
  const allSorts: Record<string, ImportedSort> = {};
  // 全局图表收集器 - 按 sheetId 存储图表信息
  const allCharts: Record<string, ImportedChart[]> = {};
  // sheetName 到 sheetKey 的映射（用于图表解析）
  const sheetNameToIdMap = new Map<string, string>();
  // sheetName 到 sheetKey 的映射列表（用于排序解析）
  const sheetIndexToIdMap: string[] = [];

  // 转换每个工作表 - 保留所有工作表包括空表
  // 使用 worksheets 数组而不是 eachSheet 以保持正确顺序
  workbook.worksheets.forEach((worksheet, sheetIndex) => {
    // 跳过 undefined 工作表
    if (!worksheet) {
      return;
    }

    const sheetKey = `sheet-${nanoid()}`;
    univerWorkbook.sheetOrder.push(sheetKey);

    const cellData: Record<number, Record<number, any>> = {};
    const mergeData: any[] = [];
    const rowData: Record<number, any> = {};
    const columnData: Record<number, any> = {};
    const images: any[] = [];
    const charts: any[] = [];
    const conditionalFormats: any[] = [];
    const dataValidations: any[] = [];

    // 冻结窗格信息
    let freezeRow = 0; // 冻结的行数
    let freezeColumn = 0; // 冻结的列数

    // 计算实际使用的行列数
    let maxRow = 0;
    let maxCol = 0;

    // 工作表名称处理 - 处理特殊字符
    const sheetName = escapeSheetName(worksheet.name || `Sheet${sheetIndex + 1}`);
    // 记录 sheetName 到 sheetKey 的映射（用于图表解析）
    sheetNameToIdMap.set(worksheet.name || `Sheet${sheetIndex + 1}`, sheetKey);
    // 记录 sheetIndex 到 sheetKey 的映射（用于排序解析）
    sheetIndexToIdMap[sheetIndex] = sheetKey;

    // 转换单元格数据
    worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
      const rowIndex = rowNumber - 1;
      maxRow = Math.max(maxRow, rowIndex);

      // 记录行高和隐藏状态
      const rowInfo: any = {};
      let hasRowInfo = false;

      // 处理行高
      if (row.height && row.height > 0) {
        rowInfo.h = row.height;
        hasRowInfo = true;
      }

      // 处理隐藏行 - ExcelJS 的 row.hidden 属性
      if (row.hidden) {
        rowInfo.hd = 1; // 1 表示隐藏
        hasRowInfo = true;
      } else {
        rowInfo.hd = 0; // 0 表示不隐藏
      }

      if (hasRowInfo) {
        rowData[rowIndex] = rowInfo;
      }

      row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        const colIndex = colNumber - 1;
        maxCol = Math.max(maxCol, colIndex);

        // 空单元格（Null 类型）但可能有样式（如背景色）
        // 需要检查是否有样式，有样式的空单元格也需要处理
        if (cell.type === ExcelJS.ValueType.Null) {
          // 检查是否有样式（背景色、边框等）
          if (cell.style) {
            const style = convertCellStyle(cell.style);
            if (style) {
              if (!cellData[rowIndex]) {
                cellData[rowIndex] = {};
              }
              cellData[rowIndex][colIndex] = { s: style };
            }
          }
          return;
        }

        // 处理 DISPIMG 公式（Excel 单元格内嵌图片，Office 365 功能）
        // 公式格式：_xlfn.DISPIMG("ID_xxx", 1) 或 DISPIMG("ID_xxx", 1)
        // 图片数据存储在 xl/cellimages.xml 中，通过 workbook.getImage() 获取
        if (cell.type === ExcelJS.ValueType.Formula && cell.formula) {
          const formula = typeof cell.formula === 'string' ? cell.formula : '';
          const dispImgMatch =
            formula.match(/_?xlfn\.?DISPIMG\s*\(\s*"([^"]+)"/i) ||
            formula.match(/DISPIMG\s*\(\s*"([^"]+)"/i);

          if (dispImgMatch) {
            const imageId = dispImgMatch[1];

            try {
              let foundImage = null;

              // 通过 workbook.getImage() 遍历查找图片
              // ExcelJS 使用数字索引存储图片，需要遍历查找
              for (let imgIdx = 0; imgIdx < 100 && !foundImage; imgIdx++) {
                try {
                  const img = workbook.getImage(imgIdx);
                  if (img && img.buffer) {
                    foundImage = img;
                  }
                } catch (e) {
                  // 索引不存在，继续查找
                }
              }

              // 备用方案：从 worksheet 的 media 获取
              if (!foundImage) {
                const wsMedia = (worksheet as any)._media || (worksheet as any).media || [];
                for (const media of wsMedia) {
                  if (media && media.type === 'image' && media.buffer) {
                    foundImage = media;
                    break;
                  }
                }
              }

              if (foundImage && foundImage.buffer) {
                // 提取图片数据并创建单元格图片记录
                const drawingId = `cell-img-${nanoid()}`;
                const base64 = arrayBufferToBase64(foundImage.buffer);
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

                // 跳过此单元格的正常处理（不显示 DISPIMG 公式）
                return;
              }
            } catch (imgError) {
              console.warn(`处理单元格图片失败 (${cell.address}):`, imgError);
            }

            // 无法提取图片时，跳过此单元格（避免显示 DISPIMG 公式）
            return;
          }
        }

        // 获取单元格值
        let rawValue = getCellValue(cell);

        // 最终安全检查：确保值不是 NaN，如果是则使用原始数据
        if (rawValue !== null && rawValue !== undefined && rawValue !== '') {
          // 如果是数字，检查是否为 NaN
          if (typeof rawValue === 'number' && (isNaN(rawValue) || !isFinite(rawValue))) {
            console.warn(
              '检测到 NaN 数字值，单元格:',
              cell.address,
              '原始值:',
              cell.value,
              '类型:',
              cell.type,
            );
            // 使用原始数据而不是空字符串
            rawValue = getOriginalCellValue(cell);
          }
          // 如果是字符串，检查是否包含 NaN
          else if (
            typeof rawValue === 'string' &&
            (rawValue === 'NaN' || rawValue.includes('NaN'))
          ) {
            console.warn(
              '检测到 NaN 字符串值，单元格:',
              cell.address,
              '原始值:',
              cell.value,
              '类型:',
              cell.type,
            );
            // 使用原始数据而不是空字符串
            rawValue = getOriginalCellValue(cell);
          }
        }

        // 判断单元格是否真正有内容
        // 注意：数值 0 是有效值，不应被过滤
        const hasValue = rawValue !== null && rawValue !== undefined && rawValue !== '';
        const hasFormula = cell.type === ExcelJS.ValueType.Formula && cell.formula;

        // 检查是否有实际的样式设置（排除默认空对象和空样式）
        const hasActualStyle =
          cell.style &&
          Object.keys(cell.style).some((key) => {
            const styleValue = (cell.style as any)[key];
            if (styleValue === null || styleValue === undefined) return false;
            if (typeof styleValue === 'object') {
              return Object.keys(styleValue).length > 0;
            }
            return true;
          });

        // 如果单元格完全为空（无值、无公式），则跳过
        // 注意：只有样式没有值的单元格也应该跳过，避免显示为0
        if (!hasValue && !hasFormula) {
          return;
        }

        if (!cellData[rowIndex]) {
          cellData[rowIndex] = {};
        }

        const cellValue: any = {};

        // 只有当值不为空时才设置 v 属性
        if (hasValue) {
          cellValue.v = rawValue;
        }

        // 设置单元格类型（传入实际值以正确判断类型）
        const cellType = getCellType(cell, rawValue);
        if (cellType !== null && hasValue) {
          cellValue.t = cellType;
        }

        // 处理公式 - 保留公式和计算结果
        if (cell.type === ExcelJS.ValueType.Formula) {
          const formula = cell.formula;
          let formulaText;
          if (formula) {
            // 处理共享公式
            if (typeof formula === 'object' && 'sharedFormula' in formula) {
              formulaText = formula.sharedFormula;
            }
            // 处理数组公式
            else if (typeof formula === 'object' && 'result' in formula) {
              formulaText = formula.formula || cell.formula;
            }
            // 普通公式
            else if (typeof formula === 'string') {
              formulaText = formula;
            }
            // 其他情况，尝试从 cell.formula 直接获取
            else if (typeof cell.formula === 'string') {
              formulaText = cell.formula;
            }

            if (formulaText) {
              // 确保公式以 = 开头（Univer 要求）
              if (!formulaText.startsWith('=')) {
                formulaText = '=' + formulaText;
              }
              cellValue.f = formulaText;
            }
            // 公式单元格也需要设置值（计算结果）
            if (rawValue !== null && rawValue !== undefined) {
              cellValue.v = rawValue;
            }
          }
        }

        // 处理富文本
        if (cell.type === ExcelJS.ValueType.RichText) {
          const richTextValue = cell.value as ExcelJS.CellRichTextValue;
          cellValue.p = convertRichText(richTextValue.richText);
        }

        // 处理超链接
        if (cell.type === ExcelJS.ValueType.Hyperlink) {
          const hyperlinkValue = cell.value as ExcelJS.CellHyperlinkValue;
          cellValue.link = {
            url: hyperlinkValue.hyperlink,
            text: hyperlinkValue.text || hyperlinkValue.hyperlink,
          };
        }

        // 处理样式
        if (cell.style) {
          const style = convertCellStyle(cell.style);
          if (style) {
            cellValue.s = style;
          }
        }

        // 只有当单元格有实际内容时才添加
        if (Object.keys(cellValue).length > 0) {
          cellData[rowIndex][colIndex] = cellValue;
        }
      });
    });

    // 处理列宽和隐藏列
    if (worksheet.columns && Array.isArray(worksheet.columns)) {
      worksheet.columns.forEach((column, index) => {
        if (column) {
          const colInfo: any = {};
          let hasColInfo = false;

          // 处理列宽
          if (column.width && column.width > 0) {
            colInfo.w = column.width * 7.5; // ExcelJS 的列宽单位转换为像素
            hasColInfo = true;
          }

          // 处理隐藏列
          if (column.hidden) {
            colInfo.hd = 1; // 1 表示隐藏
            hasColInfo = true;
          } else {
            colInfo.hd = 0; // 0 表示不隐藏
          }

          if (hasColInfo) {
            columnData[index] = colInfo;
          }
        }
      });
    }

    // 处理合并单元格
    if (worksheet.model && worksheet.model.merges) {
      worksheet.model.merges.forEach((merge: string) => {
        // merge 格式如 "A1:B2"
        const [start, end] = merge.split(':');
        const startCell = worksheet.getCell(start);
        const endCell = worksheet.getCell(end);

        const startRow = startCell.row - 1;
        const endRow = endCell.row - 1;
        const startColumn = startCell.col - 1;
        const endColumn = endCell.col - 1;

        mergeData.push({
          startRow,
          endRow,
          startColumn,
          endColumn,
        });

        // 处理合并单元格的边框
        // Excel 只在主单元格（左上角）存储边框，需要将边框应用到合并区域的边缘
        const masterCell = startCell;
        if (masterCell.style && masterCell.style.border) {
          const border = masterCell.style.border;

          // 顶部边框：应用到第一行的所有单元格
          if (border.top) {
            for (let col = startColumn; col <= endColumn; col++) {
              if (!cellData[startRow]) cellData[startRow] = {};
              if (!cellData[startRow][col]) cellData[startRow][col] = {};
              if (!cellData[startRow][col].s) cellData[startRow][col].s = {};
              if (!cellData[startRow][col].s.bd) cellData[startRow][col].s.bd = {};
              cellData[startRow][col].s.bd.t = convertSingleBorder(border.top);
            }
          }

          // 底部边框：应用到最后一行的所有单元格
          if (border.bottom) {
            for (let col = startColumn; col <= endColumn; col++) {
              if (!cellData[endRow]) cellData[endRow] = {};
              if (!cellData[endRow][col]) cellData[endRow][col] = {};
              if (!cellData[endRow][col].s) cellData[endRow][col].s = {};
              if (!cellData[endRow][col].s.bd) cellData[endRow][col].s.bd = {};
              cellData[endRow][col].s.bd.b = convertSingleBorder(border.bottom);
            }
          }

          // 左边框：应用到第一列的所有单元格
          if (border.left) {
            for (let row = startRow; row <= endRow; row++) {
              if (!cellData[row]) cellData[row] = {};
              if (!cellData[row][startColumn]) cellData[row][startColumn] = {};
              if (!cellData[row][startColumn].s) cellData[row][startColumn].s = {};
              if (!cellData[row][startColumn].s.bd) cellData[row][startColumn].s.bd = {};
              cellData[row][startColumn].s.bd.l = convertSingleBorder(border.left);
            }
          }

          // 右边框：应用到最后一列的所有单元格
          if (border.right) {
            for (let row = startRow; row <= endRow; row++) {
              if (!cellData[row]) cellData[row] = {};
              if (!cellData[row][endColumn]) cellData[row][endColumn] = {};
              if (!cellData[row][endColumn].s) cellData[row][endColumn].s = {};
              if (!cellData[row][endColumn].s.bd) cellData[row][endColumn].s.bd = {};
              cellData[row][endColumn].s.bd.r = convertSingleBorder(border.right);
            }
          }
        }
      });
    }

    // 处理条件格式（包括数据条、色阶、图标集等）
    // ExcelJS 通过 worksheet.conditionalFormattings 存储条件格式
    // 条件格式需要通过 Facade API 添加，返回给调用方处理
    if (worksheet.conditionalFormattings && Array.isArray(worksheet.conditionalFormattings)) {
      try {
        const sheetCfRules: ImportedConditionalFormat[] = [];
        worksheet.conditionalFormattings.forEach((cf: any) => {
          if (!cf || !cf.ref) return;

          // 保留原始范围引用（A1 格式，如 "A1:B10"）
          const rangeRefs = cf.ref.split(/[,\s]+/).filter((r: string) => r.trim());

          // 处理每条规则
          if (cf.rules && Array.isArray(cf.rules)) {
            cf.rules.forEach((rule: any, ruleIndex: number) => {
              const cfRule = convertConditionalFormatRuleForFacade(rule, rangeRefs, ruleIndex);
              if (cfRule) {
                sheetCfRules.push(cfRule);
              }
            });
          }
        });
        if (sheetCfRules.length > 0) {
          allConditionalFormats[sheetKey] = sheetCfRules;
        }
      } catch (error) {
        console.warn('处理条件格式时出错:', error);
      }
    }

    // 处理筛选器 (autoFilter)
    // ExcelJS 的 autoFilter 可以是字符串 "A1:D14" 或对象 { from: ..., to: ... }
    if (worksheet.autoFilter) {
      try {
        const filterRange = parseAutoFilter(worksheet.autoFilter);
        if (filterRange) {
          allFilters[sheetKey] = { range: filterRange };
          console.log(`[fileImport] 解析筛选器: sheet=${sheetKey}, range=${filterRange}`);
        }
      } catch (error) {
        console.warn('处理筛选器时出错:', error);
      }
    }

    // 处理数据验证
    worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      row.eachCell({ includeEmpty: false }, (cell) => {
        if (cell.dataValidation) {
          const validation = cell.dataValidation;
          dataValidations.push({
            row: rowNumber - 1,
            col: cell.col - 1,
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

    // 处理冻结窗格（Freeze Panes）
    // ExcelJS 通过 worksheet.views 存储视图信息，包括冻结
    if (worksheet.views && Array.isArray(worksheet.views) && worksheet.views.length > 0) {
      const view = worksheet.views[0]; // 通常只有一个视图
      if (view) {
        // state 为 'frozen' 表示有冻结
        if (view.state === 'frozen') {
          // xSplit: 冻结的列数（从左边开始）
          // ySplit: 冻结的行数（从顶部开始）
          if (typeof view.xSplit === 'number' && view.xSplit > 0) {
            freezeColumn = view.xSplit;
          }
          if (typeof view.ySplit === 'number' && view.ySplit > 0) {
            freezeRow = view.ySplit;
          }
        }
      }
    }

    // 处理图片 - 构建符合 Univer 格式的图片数据（仅当 includeImages 为 true 时）
    if (includeImages && worksheet.getImages && typeof worksheet.getImages === 'function') {
      try {
        const worksheetImages = worksheet.getImages();

        worksheetImages.forEach((img: any, imgIndex: number) => {
          if (img && img.imageId !== undefined) {
            try {
              // 获取图片数据
              const imageMedia = workbook.getImage(img.imageId);
              if (!imageMedia || !imageMedia.buffer) {
                console.warn(`图片 ${img.imageId} 没有有效的数据`);
                return;
              }

              // 生成唯一 ID
              const drawingId = `drawing-${nanoid()}`;

              // 转换图片为 base64（浏览器兼容）
              const base64 = arrayBufferToBase64(imageMedia.buffer);
              const extension = imageMedia.extension || 'png';
              const mimeType = getImageMimeType(extension);
              const imageSource = `data:${mimeType};base64,${base64}`;

              // 获取图片位置信息
              const range = img.range;
              let startColumn = 0;
              let startRow = 0;
              let endColumn = 1;
              let endRow = 1;
              let startColumnOffset = 0;
              let startRowOffset = 0;
              let endColumnOffset = 0;
              let endRowOffset = 0;

              // 图片实际尺寸（优先使用 ext 属性）
              let imageWidth = 0;
              let imageHeight = 0;

              // ExcelJS 图片可能有 ext 属性包含实际尺寸（单位：EMU，1像素 = 9525 EMU）
              const EMU_PER_PIXEL = 9525;
              if (range?.ext) {
                imageWidth = (range.ext.width ?? 0) / EMU_PER_PIXEL;
                imageHeight = (range.ext.height ?? 0) / EMU_PER_PIXEL;
              }

              if (range) {
                // 处理不同的范围格式
                if (range.tl) {
                  startColumn = Math.floor(range.tl.col ?? range.tl.nativeCol ?? 0);
                  startRow = Math.floor(range.tl.row ?? range.tl.nativeRow ?? 0);
                  // 计算偏移量（小数部分）
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
                  // 如果没有 br，默认图片占 1 个单元格
                  endColumn = startColumn + 1;
                  endRow = startRow + 1;
                }
              }

              // 如果没有从 ext 获取到尺寸，则通过单元格范围计算（作为 fallback）
              if (imageWidth <= 0 || imageHeight <= 0) {
                imageWidth =
                  (endColumn - startColumn) * DEFAULT_COLUMN_WIDTH +
                  endColumnOffset -
                  startColumnOffset;
                imageHeight =
                  (endRow - startRow) * DEFAULT_ROW_HEIGHT + endRowOffset - startRowOffset;
              }

              // 确保最小尺寸
              const width = Math.max(imageWidth, 20);
              const height = Math.max(imageHeight, 20);

              // 构建 Univer 图片数据结构 (ISheetDrawing)
              const sheetDrawing = {
                unitId: univerWorkbook.id,
                subUnitId: sheetKey,
                drawingId: drawingId,
                drawingType: 1, // DrawingType.DRAWING_IMAGE = 1
                imageSourceType: 0, // ImageSourceType.BASE64 = 0
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
                anchorType: 0, // SheetDrawingAnchorType.Position = 0 (随单元格移动)
              };

              // 添加到全局图片集合
              allDrawings[drawingId] = sheetDrawing;
              images.push({
                drawingId,
                sheetId: sheetKey,
              });

              // 添加到返回的图片列表（用于后续通过 API 插入）
              const importedImage: ImportedImage = {
                id: drawingId,
                type: ImageType.FLOATING, // Excel 中的图片默认为浮动图片
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

    // 处理图表（ExcelJS 对图表的支持有限）
    // 图表通常作为 drawing 对象存储
    // 尝试从多个来源获取图表信息
    const worksheetCharts: ImportedChart[] = [];

    // 方法1: 检查 drawings 数组
    if ((worksheet as any).drawings && Array.isArray((worksheet as any).drawings)) {
      try {
        console.log(
          `[fileImport] Sheet "${sheetName}" 发现 ${
            (worksheet as any).drawings.length
          } 个 drawings`,
        );
        (worksheet as any).drawings.forEach((drawing: any, drawingIndex: number) => {
          console.log(`[fileImport] Drawing ${drawingIndex}:`, {
            type: drawing.type,
            chartType: drawing.chartType,
            keys: Object.keys(drawing),
          });
          if (drawing.type === 'chart' || drawing.chartType) {
            const chart: ImportedChart = {
              chartId: `chart-${nanoid()}`,
              sheetId: sheetKey,
              sheetName: sheetName,
              chartType: drawing.chartType || 'column',
              dataRange: drawing.range || '',
              position: {
                row: drawing.range?.tl?.nativeRow || 0,
                column: drawing.range?.tl?.nativeCol || 0,
              },
              size: {
                width: 600,
                height: 400,
              },
              title: drawing.title || '',
              rawData: drawing,
            };
            worksheetCharts.push(chart);
          }
        });
      } catch (error) {
        console.warn('[fileImport] 处理 drawings 图表时出错:', error);
      }
    }

    // 方法2: 检查 _charts 属性（某些版本的 ExcelJS 可能使用这个）
    if ((worksheet as any)._charts && Array.isArray((worksheet as any)._charts)) {
      try {
        console.log(
          `[fileImport] Sheet "${sheetName}" 发现 ${(worksheet as any)._charts.length} 个 _charts`,
        );
        (worksheet as any)._charts.forEach((chart: any, chartIndex: number) => {
          console.log(`[fileImport] _chart ${chartIndex}:`, Object.keys(chart));
          const chartData: ImportedChart = {
            chartId: `chart-${nanoid()}`,
            sheetId: sheetKey,
            sheetName: sheetName,
            chartType: chart.type || chart.chartType || 'column',
            dataRange: chart.dataRange || chart.range || '',
            position: {
              row: chart.row || 0,
              column: chart.col || chart.column || 0,
            },
            size: {
              width: chart.width || 600,
              height: chart.height || 400,
            },
            title: chart.title || '',
            rawData: chart,
          };
          worksheetCharts.push(chartData);
        });
      } catch (error) {
        console.warn('[fileImport] 处理 _charts 时出错:', error);
      }
    }

    // 方法3: 检查 _media 中的图表（图表有时作为媒体对象存储）
    if ((worksheet as any)._media && Array.isArray((worksheet as any)._media)) {
      try {
        (worksheet as any)._media.forEach((media: any, mediaIndex: number) => {
          if (media.type === 'chart') {
            console.log(`[fileImport] 在 _media 中发现图表 ${mediaIndex}`);
          }
        });
      } catch (error) {
        console.warn('[fileImport] 处理 _media 时出错:', error);
      }
    }

    // 将收集到的图表添加到全局收集器
    if (worksheetCharts.length > 0) {
      allCharts[sheetKey] = worksheetCharts;
      console.log(`[fileImport] Sheet "${sheetName}" 共发现 ${worksheetCharts.length} 个图表`);
    }

    // 兼容旧代码：同时保留局部 charts 数组
    charts.push(...worksheetCharts);

    // 构建完整的工作表数据
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
      // 冻结窗格：freeze 对象包含 startRow 和 startColumn
      // startRow: 冻结的行数（从顶部开始，0 表示不冻结行）
      // startColumn: 冻结的列数（从左边开始，0 表示不冻结列）
      ...(freezeRow > 0 || freezeColumn > 0
        ? {
            freeze: {
              startRow: freezeRow,
              startColumn: freezeColumn,
              xSplit: freezeColumn,
              ySplit: freezeRow,
            },
          }
        : {}),
      // 添加图片和图表数据
      ...(images.length > 0 && { images }),
      ...(charts.length > 0 && { charts }),
      // 数据验证 (条件格式通过 resources 添加，不在 worksheet 中)
      ...(dataValidations.length > 0 && { dataValidations }),
    } as IWorksheetData;

    univerWorkbook.sheets[sheetKey] = univerSheet;
  });

  // 构建图片资源 (SHEET_DRAWING_PLUGIN)
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

  // 条件格式不通过 resources 加载，而是返回给调用方
  // 调用方需要使用 Facade API (fWorksheet.addConditionalFormattingRule) 来添加

  // 使用直接解析 xlsx 的方式获取图表（ExcelJS 不支持读取图表）
  const parsedCharts = await parseChartsFromXlsx(arrayBuffer, sheetNameToIdMap);
  // 合并解析到的图表到 allCharts
  for (const [sheetId, chartsArr] of Object.entries(parsedCharts)) {
    if (!allCharts[sheetId]) {
      allCharts[sheetId] = [];
    }
    allCharts[sheetId].push(...chartsArr);
  }

  // 使用直接解析 xlsx 的方式获取透视表（ExcelJS 不支持读取透视表）
  const parsedPivotTables = await parsePivotTablesFromXlsx(arrayBuffer, sheetNameToIdMap);

  // 使用直接解析 xlsx 的方式获取排序信息（ExcelJS 不支持读取排序状态）
  const parsedSorts = await parseSortsFromXlsx(arrayBuffer, sheetIndexToIdMap);
  // 合并到 allSorts
  for (const [sheetId, sortInfo] of Object.entries(parsedSorts)) {
    allSorts[sheetId] = sortInfo;
  }

  // 注意：ExcelJS workbook 对象在函数结束后会被 GC 回收
  // 如果内存压力大，可以考虑手动清理
  // @ts-ignore - 帮助 GC 更快回收
  workbook.worksheets.length = 0;

  return {
    workbookData: univerWorkbook,
    images: allImages,
    conditionalFormats: allConditionalFormats,
    filters: allFilters,
    sorts: allSorts,
    charts: allCharts,
    pivotTables: parsedPivotTables,
  };
}

/**
 * 解析 ExcelJS 的 autoFilter 为 A1 格式的范围字符串
 * ExcelJS autoFilter 可以是：
 * - 字符串: "A1:D14"
 * - 对象: { from: "A1", to: "D14" } 或 { from: { row: 1, column: 1 }, to: { row: 14, column: 4 } }
 */
function parseAutoFilter(autoFilter: any): string | null {
  if (!autoFilter) return null;

  // 如果是字符串，直接返回
  if (typeof autoFilter === 'string') {
    return autoFilter;
  }

  // 如果是对象格式
  if (typeof autoFilter === 'object') {
    const { from, to } = autoFilter;
    if (!from || !to) return null;

    // 将 from/to 转换为 A1 格式
    const fromStr = typeof from === 'string' ? from : cellRefToA1(from);
    const toStr = typeof to === 'string' ? to : cellRefToA1(to);

    if (fromStr && toStr) {
      return `${fromStr}:${toStr}`;
    }
  }

  return null;
}

/**
 * 将 { row, column } 格式转换为 A1 格式
 * 注意：ExcelJS 的 row 和 column 是 1-based
 */
function cellRefToA1(ref: { row: number; column: number }): string | null {
  if (!ref || typeof ref.row !== 'number' || typeof ref.column !== 'number') {
    return null;
  }

  const col = ref.column;
  const row = ref.row;

  // 列号转字母 (1 -> A, 2 -> B, ..., 26 -> Z, 27 -> AA)
  let colStr = '';
  let c = col;
  while (c > 0) {
    c--;
    colStr = String.fromCharCode(65 + (c % 26)) + colStr;
    c = Math.floor(c / 26);
  }

  return `${colStr}${row}`;
}

/**
 * 获取单元格的原始值（用于 NaN 回退时保留原始数据）
 * 将各种类型的原始值转换为可显示的字符串
 */
function getOriginalCellValue(cell: ExcelJS.Cell): any {
  const value = cell.value;

  if (value === null || value === undefined) {
    return '';
  }

  // 如果是 Date 对象，格式化为日期字符串
  if (value instanceof Date) {
    if (!isNaN(value.getTime())) {
      const year = value.getFullYear();
      const month = value.getMonth() + 1;
      const day = value.getDate();
      return `${year}/${month}/${day}`;
    }
    return '';
  }

  // 如果是数字，直接返回
  if (typeof value === 'number') {
    if (!isNaN(value) && isFinite(value)) {
      return value;
    }
    return '';
  }

  // 如果是字符串，直接返回
  if (typeof value === 'string') {
    return value;
  }

  // 如果是布尔值，直接返回
  if (typeof value === 'boolean') {
    return value;
  }

  // 如果是富文本对象
  if (typeof value === 'object' && 'richText' in value) {
    return (value as ExcelJS.CellRichTextValue).richText.map((rt) => rt.text).join('');
  }

  // 如果是超链接对象
  if (typeof value === 'object' && 'hyperlink' in value) {
    return (
      (value as ExcelJS.CellHyperlinkValue).text || (value as ExcelJS.CellHyperlinkValue).hyperlink
    );
  }

  // 如果是公式对象
  if (typeof value === 'object' && 'result' in value) {
    const result = (value as any).result;
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

  // 其他情况尝试转换为字符串
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
 * 获取单元格值（增强版 - 支持富文本和超链接）
 */
function getCellValue(cell: ExcelJS.Cell): any {
  // 处理公式
  if (cell.type === ExcelJS.ValueType.Formula) {
    // 优先返回计算结果，如果没有结果则保留公式
    return cell.result !== undefined ? cell.result : cell.value;
  }

  // 处理富文本
  if (cell.type === ExcelJS.ValueType.RichText) {
    const richTextValue = cell.value as ExcelJS.CellRichTextValue;
    // 简单拼接文本，保留富文本信息在单元格属性中
    return richTextValue.richText.map((rt) => rt.text).join('');
  }

  // 处理日期
  if (cell.type === ExcelJS.ValueType.Date) {
    const dateValue = cell.value;
    const numFmt = cell.style?.numFmt;

    // 空值检查
    if (dateValue === null || dateValue === undefined) {
      return '';
    }

    // 如果是 Date 对象，格式化为字符串保持原样显示
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

    // 如果是数字（Excel 序列号），转换为日期字符串
    if (
      typeof dateValue === 'number' &&
      !isNaN(dateValue) &&
      isFinite(dateValue) &&
      dateValue > 0
    ) {
      try {
        const date = excelSerialToDate(dateValue);
        if (date && date instanceof Date && !isNaN(date.getTime())) {
          const formatted = formatDateByPattern(date, numFmt);
          // 确保格式化结果有效
          if (formatted && formatted !== '' && !formatted.includes('NaN')) {
            return formatted;
          }
          // 格式化失败，返回日期的本地化字符串
          return date.toLocaleDateString('zh-CN');
        }
      } catch (error) {
        console.warn('日期转换失败:', error, '原始值:', dateValue);
      }
      // 无法转换则返回原始数字
      return dateValue;
    }

    // 如果是字符串，尝试解析为日期
    if (typeof dateValue === 'string' && dateValue.trim()) {
      // 如果格式是日期格式，尝试解析字符串日期
      if (numFmt && isDateFormat(numFmt)) {
        const parsedDate = parseDateString(dateValue);
        if (parsedDate && !isNaN(parsedDate.getTime())) {
          const formatted = formatDateByPattern(parsedDate, numFmt);
          // 确保格式化结果有效，不包含 NaN
          if (formatted && formatted !== '' && !formatted.includes('NaN')) {
            return formatted;
          }
        }
      }
      // 无法解析或不是日期格式，返回原始字符串
      return dateValue;
    }

    // 其他情况返回原始值
    return dateValue;
  }

  // 处理超链接 - 保留链接信息
  if (cell.type === ExcelJS.ValueType.Hyperlink) {
    const hyperlinkValue = cell.value as ExcelJS.CellHyperlinkValue;
    // 返回显示文本
    return hyperlinkValue.text || hyperlinkValue.hyperlink;
  }

  // 处理合并单元格
  if (cell.type === ExcelJS.ValueType.Merge) {
    return ''; // 合并单元格的非主单元格返回空
  }

  // 处理错误值
  if (cell.type === ExcelJS.ValueType.Error) {
    return cell.value;
  }

  // 处理数字类型 - 检查是否为日期格式
  if (cell.type === ExcelJS.ValueType.Number) {
    const numValue = cell.value as number;
    const numFmt = cell.style?.numFmt;

    // 检查数字格式是否为日期格式
    if (numFmt && isDateFormat(numFmt)) {
      // 将Excel序列号转换为日期
      if (typeof numValue === 'number' && !isNaN(numValue) && isFinite(numValue) && numValue > 0) {
        try {
          const date = excelSerialToDate(numValue);
          if (date && date instanceof Date && !isNaN(date.getTime())) {
            const formatted = formatDateByPattern(date, numFmt);
            // 确保格式化结果不是空字符串或 NaN
            if (formatted && formatted !== '' && !formatted.includes('NaN')) {
              return formatted;
            }
            // 格式化失败，返回日期的本地化字符串
            return date.toLocaleDateString('zh-CN');
          }
        } catch (error) {
          console.warn('日期转换失败:', error, '原始值:', numValue);
        }
        // 转换失败，返回原始数字
        return numValue;
      }
    }

    // 不是日期格式或无法转换，返回原数字
    if (typeof numValue === 'number' && !isNaN(numValue) && isFinite(numValue)) {
      return numValue;
    }
    // 如果数字无效，返回原始值
    return numValue;
  }

  // 处理字符串类型 - 检查是否为日期字符串
  if (cell.type === ExcelJS.ValueType.String) {
    const strValue = cell.value as string;
    const numFmt = cell.style?.numFmt;

    // 如果格式是日期格式，尝试解析字符串日期
    if (numFmt && isDateFormat(numFmt) && strValue) {
      const parsedDate = parseDateString(strValue);
      if (parsedDate && !isNaN(parsedDate.getTime())) {
        const formatted = formatDateByPattern(parsedDate, numFmt);
        // 确保格式化结果有效，不包含 NaN
        if (formatted && formatted !== '' && !formatted.includes('NaN')) {
          return formatted;
        }
      }
    }

    // 无论如何都返回原始字符串值
    return strValue;
  }

  // Boolean, Null
  return cell.value;
}

/**
 * 获取单元格类型
 * @param cell ExcelJS 单元格
 * @param actualValue 实际的单元格值（经过处理后的值）
 */
function getCellType(cell: ExcelJS.Cell, actualValue?: any): number | null {
  // 如果提供了实际值，根据实际值的类型来判断
  if (actualValue !== undefined && actualValue !== null && actualValue !== '') {
    if (typeof actualValue === 'boolean') {
      return 3; // Boolean
    }
    if (typeof actualValue === 'number') {
      return 2; // Number
    }
    if (typeof actualValue === 'string') {
      return 1; // String
    }
  }

  // 回退到根据单元格类型判断
  if (cell.type === ExcelJS.ValueType.Boolean) {
    return 3; // Boolean
  }
  if (cell.type === ExcelJS.ValueType.Number) {
    return 2; // Number
  }
  // 日期类型：如果实际值是字符串（格式化后的日期），应该返回字符串类型
  if (cell.type === ExcelJS.ValueType.Date) {
    // 日期被格式化为字符串后，应该作为字符串类型
    return 1; // String
  }
  if (cell.type === ExcelJS.ValueType.String || cell.type === ExcelJS.ValueType.RichText) {
    return 1; // String
  }
  return null;
}

/**
 * 转换富文本格式
 */
function convertRichText(richText: ExcelJS.RichText[]): any {
  if (!richText || richText.length === 0) {
    return undefined;
  }

  const body = {
    dataStream: '',
    textRuns: [] as any[],
  };

  let currentIndex = 0;
  richText.forEach((rt) => {
    const text = rt.text || '';
    const length = text.length;

    body.dataStream += text;

    if (rt.font) {
      const textRun: any = {
        st: currentIndex,
        ed: currentIndex + length,
      };

      if (rt.font.bold) textRun.ts = { ...textRun.ts, bl: 1 };
      if (rt.font.italic) textRun.ts = { ...textRun.ts, it: 1 };
      if (rt.font.underline) textRun.ts = { ...textRun.ts, ul: { s: 1 } };
      if (rt.font.strike) textRun.ts = { ...textRun.ts, st: { s: 1 } };
      if (rt.font.size) textRun.ts = { ...textRun.ts, fs: rt.font.size };
      if (rt.font.name) textRun.ts = { ...textRun.ts, ff: rt.font.name };
      // 使用增强的颜色解析（支持 theme + tint）
      if (rt.font.color) {
        const fontColor = parseExcelColor(rt.font.color);
        if (fontColor) {
          textRun.ts = {
            ...textRun.ts,
            cl: { rgb: fontColor },
          };
        }
      }

      if (Object.keys(textRun.ts || {}).length > 0) {
        body.textRuns.push(textRun);
      }
    }

    currentIndex += length;
  });

  return body.textRuns.length > 0 ? body : undefined;
}

/**
 * Excel 默认主题颜色（Office 主题）
 * theme 0-9 对应的基础颜色
 */
const EXCEL_THEME_COLORS: Record<number, string> = {
  0: 'FFFFFF', // lt1 (Light 1) - 通常是白色
  1: '000000', // dk1 (Dark 1) - 通常是黑色
  2: 'E7E6E6', // lt2 (Light 2) - 浅灰
  3: '44546A', // dk2 (Dark 2) - 深蓝灰
  4: '4472C4', // accent1 - 蓝色
  5: 'ED7D31', // accent2 - 橙色
  6: 'A5A5A5', // accent3 - 灰色
  7: 'FFC000', // accent4 - 金色
  8: '5B9BD5', // accent5 - 浅蓝
  9: '70AD47', // accent6 - 绿色
};

/**
 * 将 hex 颜色转换为 RGB
 */
function hexToRgb(hex: string): { r: number; g: number; b: number } {
  const result = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex);
  return result
    ? {
        r: parseInt(result[1], 16),
        g: parseInt(result[2], 16),
        b: parseInt(result[3], 16),
      }
    : { r: 0, g: 0, b: 0 };
}

/**
 * 将 RGB 转换为 hex 颜色
 */
function rgbToHex(r: number, g: number, b: number): string {
  return (
    '#' +
    [r, g, b]
      .map((x) => {
        const hex = Math.round(Math.max(0, Math.min(255, x))).toString(16);
        return hex.length === 1 ? '0' + hex : hex;
      })
      .join('')
  );
}

/**
 * 应用 tint 到颜色
 * tint > 0: 向白色混合（变亮）
 * tint < 0: 向黑色混合（变暗）
 * @param hex 原始 hex 颜色（不含 #）
 * @param tint tint 值 (-1 到 1)
 */
function applyTint(hex: string, tint: number): string {
  const rgb = hexToRgb(hex);

  let r: number, g: number, b: number;

  if (tint < 0) {
    // 向黑色混合（变暗）
    const factor = 1 + tint;
    r = rgb.r * factor;
    g = rgb.g * factor;
    b = rgb.b * factor;
  } else {
    // 向白色混合（变亮）
    r = rgb.r + (255 - rgb.r) * tint;
    g = rgb.g + (255 - rgb.g) * tint;
    b = rgb.b + (255 - rgb.b) * tint;
  }

  return rgbToHex(r, g, b);
}

/**
 * 从 ExcelJS 颜色对象解析颜色
 * 支持 argb、theme+tint、indexed 等格式
 */
function parseExcelColor(color: any): string | null {
  if (!color) return null;

  // 1. 直接使用 argb 值
  if (color.argb) {
    return `#${color.argb.slice(-6)}`;
  }

  // 2. 使用主题颜色 + tint
  if (typeof color.theme === 'number') {
    const baseColor = EXCEL_THEME_COLORS[color.theme];
    if (baseColor) {
      if (typeof color.tint === 'number' && color.tint !== 0) {
        return applyTint(baseColor, color.tint);
      }
      return `#${baseColor}`;
    }
  }

  // 3. indexed 颜色（Excel 调色板索引）- 简化处理常用颜色
  if (typeof color.indexed === 'number') {
    // 常用的 indexed 颜色
    const INDEXED_COLORS: Record<number, string> = {
      0: '#000000', // Black
      1: '#FFFFFF', // White
      2: '#FF0000', // Red
      3: '#00FF00', // Green
      4: '#0000FF', // Blue
      5: '#FFFF00', // Yellow
      6: '#FF00FF', // Magenta
      7: '#00FFFF', // Cyan
      8: '#000000', // Black
      9: '#FFFFFF', // White
      64: '#000000', // System foreground
      65: '#FFFFFF', // System background
    };
    return INDEXED_COLORS[color.indexed] || null;
  }

  return null;
}

/**
 * 转换单元格样式
 */
function convertCellStyle(style: Partial<ExcelJS.Style>): any {
  const univerStyle: any = {};

  // 字体样式
  if (style.font) {
    if (style.font.bold) univerStyle.bl = 1;
    if (style.font.italic) univerStyle.it = 1;
    if (style.font.underline) univerStyle.ul = { s: 1 };
    if (style.font.strike) univerStyle.st = { s: 1 };
    if (style.font.size) univerStyle.fs = style.font.size;
    if (style.font.name) univerStyle.ff = style.font.name;
    // 使用增强的颜色解析
    if (style.font.color) {
      const fontColor = parseExcelColor(style.font.color);
      if (fontColor) {
        univerStyle.cl = { rgb: fontColor };
      }
    }
  }

  // 背景颜色
  if (style.fill && style.fill.type === 'pattern') {
    const patternFill = style.fill as ExcelJS.FillPattern;
    // 使用增强的颜色解析
    if (patternFill.fgColor) {
      const bgColor = parseExcelColor(patternFill.fgColor);
      if (bgColor) {
        univerStyle.bg = { rgb: bgColor };
      }
    }
  }

  // 对齐方式
  if (style.alignment) {
    if (style.alignment.horizontal) {
      univerStyle.ht = convertAlignment(style.alignment.horizontal);
    }
    if (style.alignment.vertical) {
      univerStyle.vt = convertVerticalAlignment(style.alignment.vertical);
    }
    if (style.alignment.wrapText) {
      univerStyle.tb = 3; // 自动换行
    }
  }

  // 边框
  if (style.border) {
    const bd = convertBorder(style.border);
    if (bd) {
      univerStyle.bd = bd;
    }
  }

  // 数字格式（小数位数、百分比等）
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
 * 转换单个边框样式（保留颜色和样式）
 */
function convertSingleBorder(border: ExcelJS.Border): any {
  const result: any = { s: 1 }; // 默认样式

  // 边框样式映射
  const styleMap: Record<string, number> = {
    thin: 1,
    medium: 2,
    thick: 3,
    dotted: 4,
    dashed: 5,
    double: 6,
    hair: 7,
    mediumDashed: 8,
    dashDot: 9,
    mediumDashDot: 10,
    dashDotDot: 11,
    mediumDashDotDot: 12,
    slantDashDot: 13,
  };

  if (border.style) {
    result.s = styleMap[border.style] || 1;
  }

  // 解析边框颜色
  if (border.color) {
    const color = parseExcelColor(border.color);
    if (color) {
      result.cl = { rgb: color };
    } else {
      result.cl = { rgb: '#000000' };
    }
  } else {
    result.cl = { rgb: '#000000' };
  }

  return result;
}

/**
 * 转换边框样式
 */
function convertBorder(border: Partial<ExcelJS.Borders>): any {
  const result: any = {};

  if (border.top) {
    result.t = convertSingleBorder(border.top);
  }
  if (border.bottom) {
    result.b = convertSingleBorder(border.bottom);
  }
  if (border.left) {
    result.l = convertSingleBorder(border.left);
  }
  if (border.right) {
    result.r = convertSingleBorder(border.right);
  }

  return Object.keys(result).length > 0 ? result : undefined;
}

/**
 * 解析单元格范围引用（如 "A1:C10" 或 "A1:C10 D1:F10"）
 */
function parseRangeRef(
  ref: string,
): Array<{ startRow: number; endRow: number; startColumn: number; endColumn: number }> {
  const ranges: Array<{
    startRow: number;
    endRow: number;
    startColumn: number;
    endColumn: number;
  }> = [];

  // 支持多个范围（空格分隔）
  const refParts = ref.split(' ');

  refParts.forEach((part) => {
    const match = part.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/i);
    if (match) {
      const startColumn = columnLetterToIndex(match[1]);
      const startRow = parseInt(match[2], 10) - 1;
      const endColumn = columnLetterToIndex(match[3]);
      const endRow = parseInt(match[4], 10) - 1;

      ranges.push({ startRow, endRow, startColumn, endColumn });
    } else {
      // 单个单元格
      const singleMatch = part.match(/([A-Z]+)(\d+)/i);
      if (singleMatch) {
        const column = columnLetterToIndex(singleMatch[1]);
        const row = parseInt(singleMatch[2], 10) - 1;
        ranges.push({ startRow: row, endRow: row, startColumn: column, endColumn: column });
      }
    }
  });

  return ranges;
}

/**
 * 列字母转索引（A -> 0, B -> 1, AA -> 26）
 */
function columnLetterToIndex(letter: string): number {
  let index = 0;
  for (let i = 0; i < letter.length; i++) {
    index = index * 26 + (letter.charCodeAt(i) - 64);
  }
  return index - 1;
}

/**
 * 转换条件格式规则为 Facade API 可用的格式
 * 返回 ImportedConditionalFormat 结构
 */
function convertConditionalFormatRuleForFacade(
  rule: any,
  rangeRefs: string[],
  index: number,
): ImportedConditionalFormat | null {
  if (!rule || !rule.type) return null;

  // 调试日志：输出原始规则数据
  console.log('[CF Debug] Original rule:', JSON.stringify(rule, null, 2));

  const baseConfig = {
    stopIfTrue: rule.stopIfTrue || false,
    priority: rule.priority || index,
  };

  switch (rule.type) {
    case 'dataBar':
      // 尝试从多个可能的属性中获取颜色
      const positiveColor =
        parseExcelColor(rule.color) || parseExcelColor(rule.fillColor) || '#638EC6';
      console.log(
        '[CF Debug] dataBar positiveColor:',
        positiveColor,
        'from rule.color:',
        rule.color,
      );

      return {
        type: 'dataBar',
        ranges: rangeRefs,
        config: {
          ...baseConfig,
          // 正值颜色
          positiveColor,
          // 负值颜色
          negativeColor:
            parseExcelColor(rule.negativeFillColor) ||
            parseExcelColor(rule.negativeBarColor) ||
            '#FF0000',
          // 是否使用渐变
          gradient: rule.gradient !== false,
          // 是否显示值
          showValue: rule.showValue !== false,
          // 最小值和最大值配置
          minValue: rule.cfvo?.[0]
            ? {
                type: mapCfvoType(rule.cfvo[0].type),
                value: rule.cfvo[0].value,
              }
            : { type: 'min' },
          maxValue: rule.cfvo?.[1]
            ? {
                type: mapCfvoType(rule.cfvo[1].type),
                value: rule.cfvo[1].value,
              }
            : { type: 'max' },
        },
      };

    case 'colorScale':
      return {
        type: 'colorScale',
        ranges: rangeRefs,
        config: {
          ...baseConfig,
          // 色阶配置：颜色和值的对应关系
          colorScale: (rule.cfvo || []).map((cfvo: any, i: number) => ({
            color:
              parseExcelColor(rule.color?.[i]) ||
              (i === 0 ? '#F8696B' : i === (rule.cfvo?.length || 1) - 1 ? '#63BE7B' : '#FFEB84'),
            value: {
              type: mapCfvoType(cfvo.type),
              value: cfvo.value,
            },
          })),
        },
      };

    case 'iconSet':
      return {
        type: 'iconSet',
        ranges: rangeRefs,
        config: {
          ...baseConfig,
          // 图标集类型
          iconSet: rule.iconSet || '3TrafficLights1',
          // 是否显示值
          showValue: rule.showValue !== false,
          // 是否反转
          reverse: rule.reverse || false,
          // 图标配置
          icons: (rule.cfvo || []).map((cfvo: any, i: number) => ({
            type: mapCfvoType(cfvo.type),
            value: cfvo.value,
            operator: cfvo.gte !== false ? 'greaterThanOrEqual' : 'greaterThan',
          })),
        },
      };

    default:
      // 其他类型（highlightCell 等）暂不支持
      return {
        type: 'other',
        ranges: rangeRefs,
        config: {
          ...baseConfig,
          originalType: rule.type,
          originalRule: rule,
        },
      };
  }
}

/**
 * 转换条件格式规则（旧版，保留兼容）
 * 支持：数据条 (dataBar)、色阶 (colorScale)、图标集 (iconSet)、单元格值规则等
 */
function convertConditionalFormatRule(
  rule: any,
  ranges: Array<{ startRow: number; endRow: number; startColumn: number; endColumn: number }>,
  index: number,
): any {
  if (!rule || !rule.type) return null;

  const cfId = `cf-${nanoid()}`;

  // 基础结构
  const baseRule: any = {
    cfId,
    ranges,
    stopIfTrue: rule.stopIfTrue || false,
    priority: rule.priority || index,
  };

  switch (rule.type) {
    case 'dataBar':
      return convertDataBarRule(baseRule, rule);

    case 'colorScale':
      return convertColorScaleRule(baseRule, rule);

    case 'iconSet':
      return convertIconSetRule(baseRule, rule);

    case 'cellIs':
      return convertCellIsRule(baseRule, rule);

    case 'expression':
      return convertExpressionRule(baseRule, rule);

    case 'top10':
      return convertTop10Rule(baseRule, rule);

    case 'aboveAverage':
      return convertAboveAverageRule(baseRule, rule);

    case 'containsText':
    case 'notContainsText':
    case 'beginsWith':
    case 'endsWith':
      return convertTextRule(baseRule, rule);

    case 'duplicateValues':
    case 'uniqueValues':
      return convertDuplicateRule(baseRule, rule);

    default:
      // 通用处理
      return {
        ...baseRule,
        rule: {
          type: rule.type,
          ...rule,
        },
      };
  }
}

/**
 * 转换数据条规则
 * Univer 期望的格式：
 * {
 *   type: 'dataBar',
 *   isShowValue: boolean,
 *   config: {
 *     min: { type: 'min' | 'num' | 'percent' | 'percentile' | 'formula', value?: number | string },
 *     max: { type: 'max' | 'num' | 'percent' | 'percentile' | 'formula', value?: number | string },
 *     isGradient: boolean,
 *     positiveColor: string,
 *     nativeColor: string, // 负值颜色
 *   }
 * }
 */
function convertDataBarRule(baseRule: any, rule: any): any {
  // 解析 cfvo 数组（通常包含 min 和 max 配置）
  let minConfig: any = { type: 'min' };
  let maxConfig: any = { type: 'max' };

  if (rule.cfvo && Array.isArray(rule.cfvo) && rule.cfvo.length >= 2) {
    // Excel cfvo 数组: [min配置, max配置]
    const minCfvo = rule.cfvo[0];
    const maxCfvo = rule.cfvo[1];

    minConfig = {
      type: mapCfvoType(minCfvo.type),
      value: minCfvo.value,
    };
    maxConfig = {
      type: mapCfvoType(maxCfvo.type),
      value: maxCfvo.value,
    };
  }

  // 解析颜色
  const positiveColor = parseExcelColor(rule.color) || '#638EC6';
  const nativeColor = parseExcelColor(rule.negativeFillColor) || '#FF0000';

  const dataBar: any = {
    type: 'dataBar',
    isShowValue: rule.showValue !== false, // 默认显示值
    config: {
      min: minConfig,
      max: maxConfig,
      isGradient: rule.gradient !== false, // 默认使用渐变
      positiveColor: positiveColor,
      nativeColor: nativeColor,
    },
  };

  return {
    ...baseRule,
    rule: dataBar,
  };
}

/**
 * 映射 Excel cfvo 类型到 Univer CFValueType
 * Excel: 'min', 'max', 'num', 'percent', 'percentile', 'formula'
 * Univer: 'min', 'max', 'num', 'percent', 'percentile', 'formula'
 */
function mapCfvoType(excelType: string): string {
  const typeMap: Record<string, string> = {
    min: 'min',
    max: 'max',
    num: 'num',
    number: 'num',
    percent: 'percent',
    percentile: 'percentile',
    formula: 'formula',
  };
  return typeMap[excelType] || 'num';
}

/**
 * 转换色阶规则
 * Univer 期望的格式：
 * {
 *   type: 'colorScale',
 *   config: [
 *     { index: 0, color: '#F8696B', value: { type: 'min' } },
 *     { index: 1, color: '#FFEB84', value: { type: 'percentile', value: 50 } },
 *     { index: 2, color: '#63BE7B', value: { type: 'max' } },
 *   ]
 * }
 */
function convertColorScaleRule(baseRule: any, rule: any): any {
  const configList: any[] = [];

  // cfvo 和 color 是对应的数组
  const cfvoList = rule.cfvo || [];
  const colorList = rule.color || [];

  for (let i = 0; i < cfvoList.length; i++) {
    const cfvo = cfvoList[i];
    const color = colorList[i];

    configList.push({
      index: i,
      color:
        parseExcelColor(color) ||
        (i === 0 ? '#F8696B' : i === cfvoList.length - 1 ? '#63BE7B' : '#FFEB84'),
      value: {
        type: mapCfvoType(cfvo?.type || 'min'),
        value: cfvo?.value,
      },
    });
  }

  const colorScale: any = {
    type: 'colorScale',
    config: configList,
  };

  return {
    ...baseRule,
    rule: colorScale,
  };
}

/**
 * 转换图标集规则
 * Univer 期望的格式：
 * {
 *   type: 'iconSet',
 *   isShowValue: boolean,
 *   config: [
 *     { operator: 'greaterThanOrEqual', value: { type: 'percent', value: 67 }, iconType: '3TrafficLights1', iconId: '0' },
 *     { operator: 'greaterThanOrEqual', value: { type: 'percent', value: 33 }, iconType: '3TrafficLights1', iconId: '1' },
 *     { operator: 'greaterThanOrEqual', value: { type: 'min' }, iconType: '3TrafficLights1', iconId: '2' },
 *   ]
 * }
 */
function convertIconSetRule(baseRule: any, rule: any): any {
  const iconSetType = rule.iconSet || '3TrafficLights1';
  const configList: any[] = [];

  if (rule.cfvo && Array.isArray(rule.cfvo)) {
    rule.cfvo.forEach((cfvo: any, index: number) => {
      configList.push({
        operator: cfvo.gte !== false ? 'greaterThanOrEqual' : 'greaterThan',
        value: {
          type: mapCfvoType(cfvo.type),
          value: cfvo.value,
        },
        iconType: iconSetType,
        iconId: String(index),
      });
    });
  }

  const iconSet: any = {
    type: 'iconSet',
    isShowValue: rule.showValue !== false, // 默认显示值
    config: configList,
  };

  return {
    ...baseRule,
    rule: iconSet,
  };
}

/**
 * 转换单元格值比较规则
 */
function convertCellIsRule(baseRule: any, rule: any): any {
  return {
    ...baseRule,
    rule: {
      type: 'cellIs',
      operator: rule.operator, // 'equal', 'notEqual', 'greaterThan', 'lessThan', etc.
      formulae: rule.formulae || [],
      style: rule.style || {},
      dxfId: rule.dxfId,
    },
  };
}

/**
 * 转换公式规则
 */
function convertExpressionRule(baseRule: any, rule: any): any {
  return {
    ...baseRule,
    rule: {
      type: 'expression',
      formulae: rule.formulae || [],
      style: rule.style || {},
      dxfId: rule.dxfId,
    },
  };
}

/**
 * 转换 Top10 规则
 */
function convertTop10Rule(baseRule: any, rule: any): any {
  return {
    ...baseRule,
    rule: {
      type: 'top10',
      rank: rule.rank || 10,
      percent: rule.percent || false,
      bottom: rule.bottom || false,
      style: rule.style || {},
      dxfId: rule.dxfId,
    },
  };
}

/**
 * 转换高于/低于平均值规则
 */
function convertAboveAverageRule(baseRule: any, rule: any): any {
  return {
    ...baseRule,
    rule: {
      type: 'aboveAverage',
      aboveAverage: rule.aboveAverage !== false, // 默认 true（高于平均值）
      equalAverage: rule.equalAverage || false,
      stdDev: rule.stdDev, // 标准差
      style: rule.style || {},
      dxfId: rule.dxfId,
    },
  };
}

/**
 * 转换文本规则
 */
function convertTextRule(baseRule: any, rule: any): any {
  return {
    ...baseRule,
    rule: {
      type: rule.type,
      text: rule.text,
      operator: rule.operator,
      formulae: rule.formulae || [],
      style: rule.style || {},
      dxfId: rule.dxfId,
    },
  };
}

/**
 * 转换重复值/唯一值规则
 */
function convertDuplicateRule(baseRule: any, rule: any): any {
  return {
    ...baseRule,
    rule: {
      type: rule.type,
      style: rule.style || {},
      dxfId: rule.dxfId,
    },
  };
}

/**
 * 转换数字格式（支持小数位数和百分比）
 * Excel numFmt 格式说明：
 * - 0：整数
 * - 0.00：两位小数
 * - #,##0：千分位整数
 * - #,##0.00：千分位两位小数
 * - 0%：百分比整数
 * - 0.00%：百分比两位小数
 * - 0.00E+00：科学计数法
 * - @：文本
 */
function convertNumFormat(numFmt: string): any {
  if (!numFmt || numFmt === 'General') {
    return undefined;
  }

  // Univer 的数字格式结构
  const formatResult: any = {
    pattern: numFmt,
  };

  // 解析数字格式获取更多信息
  const formatInfo = parseNumFormat(numFmt);
  if (formatInfo) {
    Object.assign(formatResult, formatInfo);
  }

  return formatResult;
}

/**
 * 解析 Excel 数字格式字符串
 * 返回格式化相关的信息
 */
function parseNumFormat(numFmt: string): any {
  const result: any = {};

  // 检测是否是百分比格式
  if (numFmt.includes('%')) {
    result.isPercent = true;
  }

  // 检测是否是货币格式
  if (
    numFmt.includes('$') ||
    numFmt.includes('¥') ||
    numFmt.includes('€') ||
    numFmt.includes('£')
  ) {
    result.isCurrency = true;
  }

  // 检测是否是科学计数法
  if (numFmt.toUpperCase().includes('E+') || numFmt.toUpperCase().includes('E-')) {
    result.isScientific = true;
  }

  // 检测是否有千分位分隔符
  if (numFmt.includes(',')) {
    result.hasThousandsSeparator = true;
  }

  // 检测小数位数
  const decimalMatch = numFmt.match(/\.([0#?]+)/);
  if (decimalMatch) {
    result.decimalPlaces = decimalMatch[1].length;
  } else {
    result.decimalPlaces = 0;
  }

  // 检测日期/时间格式
  const dateTimePatterns = ['yyyy', 'yy', 'mm', 'dd', 'hh', 'ss', 'AM/PM', 'am/pm'];
  const isDateTime = dateTimePatterns.some((pattern) =>
    numFmt.toLowerCase().includes(pattern.toLowerCase()),
  );
  if (isDateTime) {
    result.isDateTime = true;
  }

  // 检测负数格式（通常用红色或括号表示）
  if (numFmt.includes('[Red]') || (numFmt.includes('(') && numFmt.includes(')'))) {
    result.hasNegativeFormat = true;
  }

  // 常见 Excel 内置格式映射
  const builtInFormats: Record<string, any> = {
    // 常规
    General: { type: 'general' },
    // 数字
    '0': { type: 'number', decimalPlaces: 0 },
    '0.00': { type: 'number', decimalPlaces: 2 },
    '#,##0': { type: 'number', decimalPlaces: 0, hasThousandsSeparator: true },
    '#,##0.00': { type: 'number', decimalPlaces: 2, hasThousandsSeparator: true },
    '#,##0.0': { type: 'number', decimalPlaces: 1, hasThousandsSeparator: true },
    '#,##0.000': { type: 'number', decimalPlaces: 3, hasThousandsSeparator: true },
    '#,##0.0000': { type: 'number', decimalPlaces: 4, hasThousandsSeparator: true },
    '0.0': { type: 'number', decimalPlaces: 1 },
    '0.000': { type: 'number', decimalPlaces: 3 },
    '0.0000': { type: 'number', decimalPlaces: 4 },
    // 百分比
    '0%': { type: 'percent', decimalPlaces: 0, isPercent: true },
    '0.0%': { type: 'percent', decimalPlaces: 1, isPercent: true },
    '0.00%': { type: 'percent', decimalPlaces: 2, isPercent: true },
    '0.000%': { type: 'percent', decimalPlaces: 3, isPercent: true },
    // 科学计数法
    '0.00E+00': { type: 'scientific', decimalPlaces: 2, isScientific: true },
    '##0.0E+0': { type: 'scientific', decimalPlaces: 1, isScientific: true },
    // 货币（人民币）
    '¥#,##0': { type: 'currency', decimalPlaces: 0, isCurrency: true, currencySymbol: '¥' },
    '¥#,##0.00': { type: 'currency', decimalPlaces: 2, isCurrency: true, currencySymbol: '¥' },
    '"¥"#,##0': { type: 'currency', decimalPlaces: 0, isCurrency: true, currencySymbol: '¥' },
    '"¥"#,##0.00': { type: 'currency', decimalPlaces: 2, isCurrency: true, currencySymbol: '¥' },
    // 货币（美元）
    '$#,##0': { type: 'currency', decimalPlaces: 0, isCurrency: true, currencySymbol: '$' },
    '$#,##0.00': { type: 'currency', decimalPlaces: 2, isCurrency: true, currencySymbol: '$' },
    '"$"#,##0': { type: 'currency', decimalPlaces: 0, isCurrency: true, currencySymbol: '$' },
    '"$"#,##0.00': { type: 'currency', decimalPlaces: 2, isCurrency: true, currencySymbol: '$' },
    // 会计格式
    '_-¥* #,##0_-': { type: 'accounting', decimalPlaces: 0, isCurrency: true, currencySymbol: '¥' },
    '_-¥* #,##0.00_-': {
      type: 'accounting',
      decimalPlaces: 2,
      isCurrency: true,
      currencySymbol: '¥',
    },
    // 日期
    'yyyy-mm-dd': { type: 'date', isDateTime: true },
    'yyyy/mm/dd': { type: 'date', isDateTime: true },
    yyyy年m月d日: { type: 'date', isDateTime: true },
    'mm-dd-yy': { type: 'date', isDateTime: true },
    'm/d/yy': { type: 'date', isDateTime: true },
    'd-mmm-yy': { type: 'date', isDateTime: true },
    'd-mmm': { type: 'date', isDateTime: true },
    'mmm-yy': { type: 'date', isDateTime: true },
    // 时间
    'h:mm': { type: 'time', isDateTime: true },
    'h:mm:ss': { type: 'time', isDateTime: true },
    'h:mm AM/PM': { type: 'time', isDateTime: true },
    'h:mm:ss AM/PM': { type: 'time', isDateTime: true },
    // 日期时间
    'yyyy-mm-dd h:mm': { type: 'datetime', isDateTime: true },
    'yyyy-mm-dd h:mm:ss': { type: 'datetime', isDateTime: true },
    'm/d/yy h:mm': { type: 'datetime', isDateTime: true },
    // 文本
    '@': { type: 'text' },
  };

  // 如果是内置格式，使用预定义的信息
  if (builtInFormats[numFmt]) {
    return { ...builtInFormats[numFmt], ...result };
  }

  // 返回解析结果
  return Object.keys(result).length > 0 ? result : undefined;
}

/**
 * 增强版 CSV 导入函数
 * @param file - CSV 文件对象
 * @param options - 导入选项
 * @returns IWorkbookData
 */
export async function importCsv(
  file: File | Blob,
  options: {
    delimiter?: string; // 分隔符，默认自动检测
    encoding?: string; // 编码，默认 UTF-8
    hasHeader?: boolean; // 是否有表头，默认 true
    sheetName?: string; // 工作表名称
    skipEmptyLines?: boolean; // 跳过空行，默认 true
    trimValues?: boolean; // 修剪单元格值的空白，默认 true
  } = {},
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
    // 读取文件内容，支持多种编码
    let text: string;
    try {
      const arrayBuffer = await file.arrayBuffer();
      const decoder = new TextDecoder(encoding);
      text = decoder.decode(arrayBuffer);
    } catch (encodingError) {
      // 如果指定编码失败，尝试其他常见编码
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
        } catch (e) {
          continue;
        }
      }

      if (!decoded) {
        throw new Error('无法解码文件内容');
      }
    }

    // 自动检测分隔符
    const detectedDelimiter = delimiter || detectDelimiter(text);

    // 解析 CSV
    const rows = parseCsvText(text, detectedDelimiter, skipEmptyLines);

    if (rows.length === 0) {
      return getDefaultWorkbookData({ name: sheetName });
    }

    // 处理数据类型转换
    const processedRows = rows.map((row) =>
      row.map((cell) => {
        const trimmedCell = trimValues ? cell.trim() : cell;
        return convertCellType(trimmedCell);
      }),
    );

    // 构建工作表
    const sheet: IWorksheetData = buildSheetFrom2DArray(sheetName, processedRows);

    // 如果有表头，为表头行添加样式
    if (hasHeader && processedRows.length > 0) {
      if (!sheet.cellData[0]) {
        sheet.cellData[0] = {};
      }
      Object.keys(sheet.cellData[0]).forEach((colKey) => {
        const col = parseInt(colKey, 10);
        if (sheet.cellData[0][col]) {
          sheet.cellData[0][col].s = {
            bl: 1, // 加粗
            fs: 11, // 字体大小
            ht: 2, // 居中对齐
            vt: 2, // 垂直居中
            bg: { rgb: '#F5F5F5' }, // 背景色
            bd: {
              b: { s: 1, cl: { rgb: '#D0D0D0' } },
            },
          };
        }
      });

      // 设置表头行高
      if (!sheet.rowData) {
        sheet.rowData = {};
      }
      sheet.rowData[0] = {
        h: 25, // 表头行高
        hd: 0,
      };
    }

    // 自动调整列宽（根据内容长度）
    if (processedRows.length > 0) {
      const columnWidths: Record<number, number> = {};

      // 计算每列的最大宽度
      processedRows.forEach((row, rowIndex) => {
        row.forEach((cell, colIndex) => {
          const cellStr = String(cell);
          // 考虑中文字符占用更多空间
          const chineseChars = (cellStr.match(/[\u4e00-\u9fa5]/g) || []).length;
          const otherChars = cellStr.length - chineseChars;
          const estimatedWidth = chineseChars * 12 + otherChars * 8 + 20;

          columnWidths[colIndex] = Math.max(
            columnWidths[colIndex] || 80, // 最小宽度
            Math.min(estimatedWidth, 400), // 最大宽度
          );
        });
      });

      sheet.columnData = {};
      Object.keys(columnWidths).forEach((colKey) => {
        const col = parseInt(colKey, 10);
        sheet.columnData![col] = {
          w: columnWidths[col],
          hd: 0,
        };
      });
    }

    return getWorkbookDataBySheets([sheet]);
  } catch (error) {
    console.error('CSV 导入失败:', error);
    throw new Error(`CSV 导入失败: ${error.message}`);
  }
}

/**
 * 检测 CSV 分隔符（改进版）
 */
function detectDelimiter(text: string): string {
  const delimiters = [',', ';', '\t', '|', ':'];
  const firstLines = text.split(/\r?\n/).slice(0, 5); // 检查前5行

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

    // 如果每行的分隔符数量一致，加分
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
 * 解析 CSV 文本（改进版 - 支持引号内的分隔符和换行）
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
        // 转义的引号（""表示一个引号字符）
        currentCell += '"';
        i += 2;
        continue;
      } else if (inQuotes) {
        // 结束引号
        inQuotes = false;
        i++;
        continue;
      } else if (currentCell === '' || text[i - 1] === delimiter) {
        // 开始引号（必须在单元格开始或分隔符后）
        inQuotes = true;
        i++;
        continue;
      } else {
        // 单元格中间的引号，作为普通字符处理
        currentCell += char;
        i++;
        continue;
      }
    } else if (char === delimiter && !inQuotes) {
      // 分隔符（不在引号内）
      currentRow.push(currentCell);
      currentCell = '';
      i++;
    } else if ((char === '\n' || char === '\r') && !inQuotes) {
      // 换行符（不在引号内）
      if (char === '\r' && nextChar === '\n') {
        i++; // 跳过 \r\n 中的 \n
      }

      // 保存当前单元格
      currentRow.push(currentCell);
      currentCell = '';

      // 保存当前行
      if (!skipEmptyLines || currentRow.some((cell) => cell !== '')) {
        rows.push(currentRow);
      }
      currentRow = [];
      i++;
    } else {
      // 普通字符
      currentCell += char;
      i++;
    }
  }

  // 处理最后一个单元格和行
  if (currentCell !== '' || currentRow.length > 0) {
    currentRow.push(currentCell);
    if (!skipEmptyLines || currentRow.some((cell) => cell !== '')) {
      rows.push(currentRow);
    }
  }

  // 规范化行长度（确保所有行的列数一致）
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
 * 转换单元格类型（增强版 - 字符串 -> 数字/布尔/日期）
 */
function convertCellType(value: string): any {
  if (value === '') return '';

  // 布尔值
  const lowerValue = value.toLowerCase();
  if (lowerValue === 'true' || lowerValue === 'yes' || lowerValue === 'y') return true;
  if (lowerValue === 'false' || lowerValue === 'no' || lowerValue === 'n') return false;

  // 数字（包括负数、小数、科学计数法、货币符号）
  const cleanedValue = value.replace(/[,$¥€£]/g, '');
  const numberPattern = /^-?\d+\.?\d*([eE][+-]?\d+)?$/;
  if (numberPattern.test(cleanedValue)) {
    const num = parseFloat(cleanedValue);
    if (!isNaN(num)) return num;
  }

  // 百分比
  if (value.endsWith('%')) {
    const num = parseFloat(value.slice(0, -1));
    if (!isNaN(num)) return num / 100;
  }

  // 日期（支持多种格式）
  // YYYY-MM-DD
  const datePattern1 = /^\d{4}-\d{2}-\d{2}$/;
  if (datePattern1.test(value)) {
    const date = new Date(value);
    if (!isNaN(date.getTime())) {
      return date.toLocaleDateString('zh-CN');
    }
  }

  // MM/DD/YYYY or DD/MM/YYYY
  const datePattern2 = /^\d{1,2}\/\d{1,2}\/\d{4}$/;
  if (datePattern2.test(value)) {
    const date = new Date(value);
    if (!isNaN(date.getTime())) {
      return date.toLocaleDateString('zh-CN');
    }
  }

  // YYYY年MM月DD日
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

  // 默认返回字符串
  return value;
}

/**
 * 将 JavaScript Date 对象转换为 Excel 序列号
 * Excel 序列号从 1900-01-01 开始计算（但有 1900 年闰年 bug）
 * @param date JavaScript Date 对象
 * @returns Excel 序列号
 */
function dateToExcelSerial(date: Date): number {
  // Excel 的日期起点是 1899-12-30（因为 Excel 错误地认为 1900 是闰年）
  const excelEpoch = new Date(Date.UTC(1899, 11, 30));
  const msPerDay = 24 * 60 * 60 * 1000;

  // 获取 UTC 时间的差值
  const utcDate = Date.UTC(date.getFullYear(), date.getMonth(), date.getDate());
  const serial = (utcDate - excelEpoch.getTime()) / msPerDay;

  return serial;
}

/**
 * 将 Excel 序列号转换为 JavaScript Date 对象
 * @param serial Excel 序列号
 * @returns JavaScript Date 对象，如果转换失败返回 null
 */
function excelSerialToDate(serial: number): Date | null {
  // 防御性检查
  if (typeof serial !== 'number' || isNaN(serial) || !isFinite(serial)) {
    return null;
  }

  // Excel 的日期起点是 1899-12-30
  const excelEpoch = new Date(Date.UTC(1899, 11, 30));
  const msPerDay = 24 * 60 * 60 * 1000;

  try {
    const date = new Date(excelEpoch.getTime() + serial * msPerDay);
    // 验证日期是否有效
    if (date instanceof Date && !isNaN(date.getTime())) {
      return date;
    }
  } catch (error) {
    console.warn('Excel 序列号转换日期失败:', error, '序列号:', serial);
  }

  return null;
}

/**
 * 解析日期字符串（支持多种格式）
 * @param dateStr 日期字符串，如 "2026/1/7", "2026-01-07", "2026年1月7日" 等
 * @returns Date 对象，如果解析失败返回 null
 */
function parseDateString(dateStr: string): Date | null {
  if (!dateStr || typeof dateStr !== 'string') {
    return null;
  }

  // 移除前后空格
  const trimmed = dateStr.trim();
  if (!trimmed) {
    return null;
  }

  // 格式 1: yyyy/m/d 或 yyyy/mm/dd (如 "2026/1/7", "2026/01/07")
  const pattern1 = /^(\d{4})\/(\d{1,2})\/(\d{1,2})$/;
  const match1 = trimmed.match(pattern1);
  if (match1) {
    const year = parseInt(match1[1], 10);
    const month = parseInt(match1[2], 10) - 1; // 月份从 0 开始
    const day = parseInt(match1[3], 10);
    if (!isNaN(year) && !isNaN(month) && !isNaN(day)) {
      const date = new Date(year, month, day);
      if (!isNaN(date.getTime())) {
        return date;
      }
    }
  }

  // 格式 2: yyyy-m-d 或 yyyy-mm-dd (如 "2026-1-7", "2026-01-07")
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

  // 格式 3: yyyy年m月d日 (如 "2026年1月7日")
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

  // 格式 4: m/d/yyyy 或 mm/dd/yyyy (美式，如 "1/7/2026", "01/07/2026")
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

  // 格式 5: 尝试使用原生 Date 解析（作为后备方案）
  try {
    const date = new Date(trimmed);
    if (date instanceof Date && !isNaN(date.getTime())) {
      // 验证日期是否合理（年份在 1900-2100 之间）
      const year = date.getFullYear();
      if (year >= 1900 && year <= 2100) {
        return date;
      }
    }
  } catch (error) {
    // 忽略解析错误
  }

  return null;
}

/**
 * 根据 Excel 数字格式模式格式化日期
 * @param date JavaScript Date 对象
 * @param numFmt Excel 数字格式字符串
 * @returns 格式化后的日期字符串
 */
function formatDateByPattern(date: Date, numFmt?: string): string {
  // 防御性检查：确保 date 是有效的 Date 对象
  if (!date || !(date instanceof Date) || isNaN(date.getTime())) {
    console.warn('formatDateByPattern: 无效的日期对象', date);
    return '';
  }

  try {
    const year = date.getFullYear();
    const month = date.getMonth() + 1;
    const day = date.getDate();
    const hours = date.getHours();
    const minutes = date.getMinutes();
    const seconds = date.getSeconds();

    // 二次检查：确保提取的值都是有效数字
    if (
      isNaN(year) ||
      isNaN(month) ||
      isNaN(day) ||
      isNaN(hours) ||
      isNaN(minutes) ||
      isNaN(seconds)
    ) {
      console.warn('formatDateByPattern: 日期值包含 NaN', {
        year,
        month,
        day,
        hours,
        minutes,
        seconds,
      });
      return '';
    }

    // 验证年份、月份、日期是否在合理范围内
    if (year < 1900 || year > 2100 || month < 1 || month > 12 || day < 1 || day > 31) {
      console.warn('formatDateByPattern: 日期值超出合理范围', { year, month, day });
      return '';
    }

    // 辅助函数：安全格式化并检查 NaN
    const safeFormat = (str: string): string => {
      if (!str || str.includes('NaN') || str.includes('undefined') || str.includes('null')) {
        return '';
      }
      return str;
    };

    // 如果没有指定格式或是通用格式，使用 yyyy/m/d 格式
    if (!numFmt || numFmt === 'General') {
      const result = `${year}/${month}/${day}`;
      // 最终检查：确保结果不包含 NaN
      return safeFormat(result);
    }

    // 根据常见的 Excel 日期格式进行匹配
    // yyyy/m/d 或 yyyy/mm/dd 格式
    if (numFmt.includes('yyyy') && numFmt.includes('/')) {
      if (numFmt.includes('h:mm') || numFmt.includes('hh:mm')) {
        // 日期时间格式
        const mm = String(minutes).padStart(2, '0');
        const ss = String(seconds).padStart(2, '0');
        if (numFmt.includes('mm/dd')) {
          const result = `${year}/${String(month).padStart(2, '0')}/${String(day).padStart(
            2,
            '0',
          )} ${hours}:${mm}:${ss}`;
          return safeFormat(result);
        }
        const result = `${year}/${month}/${day} ${hours}:${mm}:${ss}`;
        return safeFormat(result);
      }
      // 纯日期格式
      if (numFmt.includes('mm') && numFmt.includes('dd')) {
        const result = `${year}/${String(month).padStart(2, '0')}/${String(day).padStart(2, '0')}`;
        return safeFormat(result);
      }
      const result = `${year}/${month}/${day}`;
      return safeFormat(result);
    }

    // yyyy-mm-dd 格式
    if (numFmt.includes('yyyy') && numFmt.includes('-')) {
      if (numFmt.includes('mm') && numFmt.includes('dd')) {
        const result = `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
        return safeFormat(result);
      }
      const result = `${year}-${month}-${day}`;
      return safeFormat(result);
    }

    // m/d/yy 或 mm/dd/yy 格式（美式）
    if (numFmt.match(/m+\/d+\/y+/i)) {
      const yy = String(year).slice(-2);
      if (numFmt.includes('mm') && numFmt.includes('dd')) {
        const result = `${String(month).padStart(2, '0')}/${String(day).padStart(2, '0')}/${yy}`;
        return safeFormat(result);
      }
      const result = `${month}/${day}/${yy}`;
      return safeFormat(result);
    }

    // d/m/yy 或 dd/mm/yy 格式（欧式）
    if (numFmt.match(/d+\/m+\/y+/i)) {
      const yy = String(year).slice(-2);
      if (numFmt.includes('mm') && numFmt.includes('dd')) {
        const result = `${String(day).padStart(2, '0')}/${String(month).padStart(2, '0')}/${yy}`;
        return safeFormat(result);
      }
      const result = `${day}/${month}/${yy}`;
      return safeFormat(result);
    }

    // yyyy年m月d日 格式
    if (numFmt.includes('年') && numFmt.includes('月') && numFmt.includes('日')) {
      const result = `${year}年${month}月${day}日`;
      return safeFormat(result);
    }

    // 如果无法识别格式，默认使用 yyyy/m/d
    const result = `${year}/${month}/${day}`;
    // 最终检查：确保结果不包含 NaN
    if (result.includes('NaN')) {
      // 尝试使用本地化日期格式作为后备
      try {
        return date.toLocaleDateString('zh-CN');
      } catch {
        return '';
      }
    }
    return result;
  } catch (error) {
    console.error('formatDateByPattern: 格式化日期时出错', error, date);
    // 尝试使用本地化日期格式作为后备
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
 * 将列号转换为 Excel 列名（如 0 -> A, 25 -> Z, 26 -> AA）
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
 * 将行列索引转换为单元格地址（如 (0, 0) -> A1）
 */
function getCellAddress(row: number, column: number): string {
  return `${columnIndexToLetter(column)}${row + 1}`;
}

/**
 * 插入图片到工作表
 * 支持浮动图片和单元格图片两种模式
 *
 * @param univerAPI Univer API 实例
 * @param images 要插入的图片列表
 * @param options 插入选项
 */
async function insertImagesAfterImport(
  univerAPI: any,
  images: ImportedImage[],
  options: {
    /** 默认图片类型（如果图片没有指定类型） */
    defaultType?: ImageType;
    /** 是否在插入失败时继续处理其他图片 */
    continueOnError?: boolean;
    /** 插入完成回调 */
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

  const fWorkbook = univerAPI.getActiveWorkbook();
  if (!fWorkbook) {
    throw new Error('无法获取当前工作簿');
  }

  for (let i = 0; i < images.length; i++) {
    const image = images[i];

    try {
      // 直接使用图片的 sheetId（已经与实际创建的 sheetId 一致）
      const actualSheetId = image.sheetId;

      // 获取对应的工作表
      const fWorksheet = fWorkbook.getSheetBySheetId(actualSheetId);
      if (!fWorksheet) {
        throw new Error(`找不到工作表: ${image.sheetName} (ID: ${actualSheetId})`);
      }

      const imageType = image.type || defaultType;

      if (imageType === ImageType.CELL) {
        // 单元格图片：嵌入到指定单元格内
        const cellAddress = getCellAddress(image.position.row, image.position.column);
        const fRange = fWorksheet.getRange(cellAddress);

        if (!fRange) {
          throw new Error(`无法获取单元格范围: ${cellAddress}`);
        }

        // 使用 insertCellImageAsync 插入单元格图片
        const insertResult = await fRange.insertCellImageAsync(image.source);
        if (!insertResult) {
          throw new Error(`单元格图片插入失败: ${cellAddress}`);
        }
      } else {
        // 浮动图片：使用 newOverGridImage 构建器模式插入
        const { column, row, columnOffset, rowOffset } = image.position;
        const { width, height } = image.size;

        // 判断图片来源类型（base64 或 URL）
        const isBase64 = image.source.startsWith('data:');
        const sourceType = isBase64
          ? univerAPI.Enum?.ImageSourceType?.BASE64 ?? 0
          : univerAPI.Enum?.ImageSourceType?.URL ?? 1;

        // 使用 newOverGridImage 构建器创建图片
        const imageBuilder = fWorksheet
          .newOverGridImage()
          .setSource(image.source, sourceType)
          .setColumn(column)
          .setRow(row)
          .setColumnOffset(columnOffset || 0)
          .setRowOffset(rowOffset || 0)
          .setWidth(width)
          .setHeight(height);

        // 构建图片对象
        const builtImage = await imageBuilder.buildAsync();
        if (!builtImage) {
          throw new Error(`构建浮动图片失败: (${column}, ${row})`);
        }

        // 插入图片
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

/**
 * 批量插入浮动图片
 * 简化版本，只支持浮动图片
 *
 * @param fWorksheet Univer 工作表实例
 * @param images 图片列表
 */
async function insertFloatingImages(
  fWorksheet: any,
  images: Array<{
    source: string;
    column: number;
    row: number;
    offsetX?: number;
    offsetY?: number;
  }>,
): Promise<{ success: number; failed: number }> {
  const result = { success: 0, failed: 0 };

  for (const img of images) {
    try {
      await fWorksheet.insertImage(
        img.source,
        img.column,
        img.row,
        img.offsetX || 0,
        img.offsetY || 0,
      );
      result.success++;
    } catch (error) {
      console.warn('插入浮动图片失败:', error);
      result.failed++;
    }
  }

  return result;
}

/**
 * 批量插入单元格图片
 * 简化版本，只支持单元格图片
 *
 * @param fWorksheet Univer 工作表实例
 * @param images 图片列表
 */
async function insertCellImages(
  fWorksheet: any,
  images: Array<{
    source: string;
    cellAddress: string; // 如 'A1', 'B2'
  }>,
): Promise<{ success: number; failed: number }> {
  const result = { success: 0, failed: 0 };

  for (const img of images) {
    try {
      const fRange = fWorksheet.getRange(img.cellAddress);
      if (fRange) {
        await fRange.insertCellImageAsync(img.source);
        result.success++;
      } else {
        result.failed++;
      }
    } catch (error) {
      console.warn(`插入单元格图片失败 (${img.cellAddress}):`, error);
      result.failed++;
    }
  }

  return result;
}

// ==================== 统一导出接口 ====================

/**
 * 图片插入选项
 */
export interface ImageInsertOptions {
  /** 默认图片类型（如果图片没有指定类型） */
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
 */
export interface FileImportResult {
  /** 工作簿数据 */
  workbookData: IWorkbookData;
  /** 导入的图片列表（仅当 includeImages=true 时有值） */
  images: ImportedImage[];
  /**
   * 条件格式数据，按 sheetId 分组
   * 需要在工作簿创建后通过 Facade API 添加
   */
  conditionalFormats: Record<string, ImportedConditionalFormat[]>;
  /**
   * 筛选器数据，按 sheetId 分组
   * 需要在工作簿创建后通过 Facade API 添加
   */
  filters: Record<string, ImportedFilter>;
  /**
   * 排序数据，按 sheetId 分组
   * 需要在工作簿创建后通过 Facade API 添加
   */
  sorts: Record<string, ImportedSort>;
  /**
   * 图表数据，按 sheetId 分组
   * 需要在工作簿创建后通过 Facade API 添加
   */
  charts: Record<string, ImportedChart[]>;
  /**
   * 透视表数据列表
   * 需要在工作簿创建后通过 Facade API 添加
   */
  pivotTables: ImportedPivotTable[];
}

/**
 * 统一的文件导入接口
 *
 * 支持 Excel (.xlsx, .xls) 和 CSV (.csv) 文件
 *
 * @example
 * ```ts
 * // 导入文件
 * const result = await importFile(file);
 *
 * // 创建 sheet（使用 originalSheetId 作为实际的 sheetId）
 * await createAndSaveSheets({ allSheetsData, ... });
 *
 * // 插入图片
 * if (result.images.length > 0) {
 *   await result.insertImages(univerAPI, {
 *     defaultType: ImageType.FLOATING,
 *   });
 * }
 * ```
 *
 * @param file 文件对象
 * @param options 导入选项
 * @returns 导入结果，包含 workbookData、images 和 insertImages 方法
 */
export async function importFile(
  file: File,
  options: FileImportOptions = {},
): Promise<FileImportResult> {
  const { includeImages = true } = options;

  // 获取文件类型
  const fileExt = file.name.split('.').pop()?.toLowerCase();
  if (!fileExt || !['xlsx', 'xls', 'csv'].includes(fileExt)) {
    throw new Error(`不支持的文件格式: ${fileExt}`);
  }

  const fileType = fileExt as 'xlsx' | 'xls' | 'csv';

  // 使用统一的内部函数处理文件导入
  const result = await handleFileImport(file, fileType, includeImages);
  const { workbookData, images, conditionalFormats, filters, sorts, charts, pivotTables } = result;

  // 返回结果
  return {
    workbookData,
    images,
    conditionalFormats,
    filters,
    sorts,
    charts,
    pivotTables,
  };
}

/**
 * 添加条件格式到工作簿
 * 使用 Univer Facade API (FWorksheet.addConditionalFormattingRule)
 * @param univerAPI Univer API 实例
 * @param conditionalFormats 条件格式数据，按 sheetId 分组
 */
export async function addConditionalFormatsToWorkbook(
  univerAPI: any,
  conditionalFormats: Record<string, ImportedConditionalFormat[]>,
): Promise<void> {
  const fWorkbook = univerAPI.getActiveWorkbook();
  if (!fWorkbook) {
    console.warn('[fileImport] 无法获取活动工作簿');
    return;
  }

  for (const [sheetId, rules] of Object.entries(conditionalFormats)) {
    // 根据 sheetId 获取工作表
    const fWorksheet = fWorkbook.getSheetBySheetId(sheetId);
    if (!fWorksheet) {
      continue;
    }

    for (const cfRule of rules) {
      try {
        await addSingleConditionalFormat(fWorksheet, cfRule);
      } catch (err) {
        console.error(`[fileImport] 添加条件格式失败:`, err, cfRule);
      }
    }
  }
}

/**
 * 添加单个条件格式规则
 * @internal
 */
/**
 * 将列字母转换为数字（1-based，用于范围裁剪）
 * A=1, B=2, ..., Z=26, AA=27, ...
 */
function colToNum1Based(col: string): number {
  let num = 0;
  for (let i = 0; i < col.length; i++) {
    num = num * 26 + (col.charCodeAt(i) - 64);
  }
  return num;
}

/**
 * 解析 A1 格式的范围字符串并裁剪到工作表边界内
 * @param rangeStr A1 格式的范围字符串（如 "A1:B10" 或 "K88:XFD90"）
 * @param maxRows 工作表最大行数
 * @param maxCols 工作表最大列数
 * @returns 裁剪后的范围字符串，如果范围完全超出边界则返回 null
 */
function clipRangeToBounds(rangeStr: string, maxRows: number, maxCols: number): string | null {
  if (!rangeStr || typeof rangeStr !== 'string') {
    return null;
  }

  // 解析范围字符串：A1:B10 或 A1（单个单元格）
  const rangeMatch = rangeStr.match(/^([A-Z]+)(\d+)(:([A-Z]+)(\d+))?$/i);
  if (!rangeMatch) {
    return null;
  }

  const startColStr = rangeMatch[1].toUpperCase();
  let startRow = parseInt(rangeMatch[2], 10);
  const endColStr = (rangeMatch[4] || startColStr).toUpperCase();
  let endRow = parseInt(rangeMatch[5] || rangeMatch[2], 10);

  // 转换为数字索引（1-based）
  let startCol = colToNum1Based(startColStr);
  let endCol = colToNum1Based(endColStr);

  // 裁剪到边界内
  startRow = Math.max(1, Math.min(startRow, maxRows));
  endRow = Math.max(1, Math.min(endRow, maxRows));
  startCol = Math.max(1, Math.min(startCol, maxCols));
  endCol = Math.max(1, Math.min(endCol, maxCols));

  // 如果裁剪后范围无效，返回 null
  if (startRow > endRow || startCol > endCol) {
    return null;
  }

  // 如果范围被裁剪，记录警告
  const originalEndCol = colToNum1Based(endColStr);
  if (originalEndCol > maxCols) {
    console.warn(
      `[fileImport] 条件格式范围 ${rangeStr} 的列超出边界（最大列: ${maxCols}），已裁剪到 ${numToCol(
        endCol,
      )}${endRow}`,
    );
  }

  // 返回裁剪后的范围字符串
  if (startRow === endRow && startCol === endCol) {
    // 单个单元格
    return `${numToCol(startCol)}${startRow}`;
  } else {
    // 范围
    return `${numToCol(startCol)}${startRow}:${numToCol(endCol)}${endRow}`;
  }
}

async function addSingleConditionalFormat(
  fWorksheet: any,
  cfRule: ImportedConditionalFormat,
): Promise<void> {
  const { type, ranges, config } = cfRule;
  if (!ranges || ranges.length === 0) return;

  // 获取工作表的最大行数和列数
  const maxRows = fWorksheet.getMaxRows?.() || 1000;
  const maxCols = fWorksheet.getMaxColumns?.() || 1000;

  // 验证并裁剪所有范围到工作表边界内
  const validRanges = ranges
    .map((r: string) => {
      const clippedRange = clipRangeToBounds(r, maxRows, maxCols);
      if (!clippedRange) {
        console.warn(`[fileImport] 条件格式范围 ${r} 超出工作表边界或无效，已跳过`);
        return null;
      }
      return clippedRange;
    })
    .filter((r): r is string => r !== null);

  if (validRanges.length === 0) {
    console.warn(`[fileImport] 条件格式规则没有有效范围`);
    return;
  }

  // 获取第一个有效范围作为主范围
  const primaryRange = validRanges[0];
  let fRange: any = null;
  try {
    fRange = fWorksheet.getRange(primaryRange);
  } catch (error) {
    console.warn(`[fileImport] 无法获取范围: ${primaryRange}`, error);
    return;
  }

  if (!fRange) {
    console.warn(`[fileImport] 无法获取范围: ${primaryRange}`);
    return;
  }

  // 创建条件格式构建器
  const builder = fWorksheet.newConditionalFormattingRule();
  if (!builder) {
    console.warn('[fileImport] 无法创建条件格式构建器');
    return;
  }

  // 设置范围 - 将所有有效范围转换为 IRange 格式
  const rangeObjects = validRanges
    .map((r: string) => {
      try {
        const range = fWorksheet.getRange(r);
        return range ? range.getRange() : null;
      } catch (error) {
        console.warn(`[fileImport] 获取范围失败: ${r}`, error);
        return null;
      }
    })
    .filter(Boolean);

  if (rangeObjects.length === 0) {
    console.warn('[fileImport] 无法将范围转换为 IRange 格式');
    return;
  }

  // 根据类型设置规则并添加
  let ruleBuilder: any = null;

  switch (type) {
    case 'dataBar':
      // setDataBar 参数格式：{ min, max, positiveColor, nativeColor, isGradient?, isShowValue? }
      ruleBuilder = builder.setDataBar({
        positiveColor: config.positiveColor || '#638EC6',
        nativeColor: config.negativeColor || '#FF0000',
        isGradient: config.gradient !== false,
        isShowValue: config.showValue !== false,
        min: config.minValue || { type: 'min' },
        max: config.maxValue || { type: 'max' },
      });
      break;

    case 'colorScale':
      if (config.colorScale && Array.isArray(config.colorScale)) {
        // setColorScale 参数格式：[{ index, color, value: { type, value? } }]
        const colorScaleConfig = config.colorScale.map((cs: any, index: number) => ({
          index,
          color: cs.color,
          value: cs.value || { type: 'num', value: 0 },
        }));
        ruleBuilder = builder.setColorScale(colorScaleConfig);
      }
      break;

    case 'iconSet':
      // setIconSet 参数格式：{ iconConfigs: [...], isShowValue }
      const iconType = config.iconSet || '3Arrows';
      ruleBuilder = builder.setIconSet({
        isShowValue: config.showValue !== false,
        iconConfigs:
          config.icons?.map((icon: any, index: number) => ({
            iconType,
            iconId: String(index),
            operator: icon.operator || 'greaterThanOrEqual',
            value: icon.value || { type: 'num', value: 0 },
          })) || [],
      });
      break;

    default:
      // 其他类型暂不支持
      console.warn(`[fileImport] 不支持的条件格式类型: ${type}`);
      return;
  }

  // 如果没有成功创建 ruleBuilder，则退出
  if (!ruleBuilder) {
    console.warn('[fileImport] 无法创建条件格式规则构建器');
    return;
  }

  // 设置范围并构建规则
  ruleBuilder.setRanges(rangeObjects);
  const rule = ruleBuilder.build();

  if (rule) {
    fWorksheet.addConditionalFormattingRule(rule);
    console.log(`[fileImport] 已添加 ${type} 条件格式规则，范围: ${validRanges.join(', ')}`);
  }
}

/**
 * 添加筛选器到工作簿
 * 使用 Univer Facade API (FRange.createFilter)
 * @param univerAPI Univer API 实例
 * @param filters 筛选器数据，按 sheetId 分组
 */
export async function addFiltersToWorkbook(
  univerAPI: any,
  filters: Record<string, ImportedFilter>,
): Promise<void> {
  const fWorkbook = univerAPI.getActiveWorkbook();
  if (!fWorkbook) {
    return;
  }

  let successCount = 0;
  const totalCount = Object.keys(filters).length;

  for (const [sheetId, filterInfo] of Object.entries(filters)) {
    const targetSheet = fWorkbook.getSheetBySheetId(sheetId);
    if (!targetSheet) {
      continue;
    }

    try {
      let { range } = filterInfo;
      if (!range) {
        continue;
      }

      // 如果筛选范围只有一行（如 A1:DY1），需要扩展到实际数据区域
      const rangeMatch = range.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
      if (rangeMatch) {
        const [, startCol, startRow, endCol, endRow] = rangeMatch;
        if (startRow === endRow) {
          const maxRow = targetSheet.getMaxRows?.() || 1000;
          range = `${startCol}${startRow}:${endCol}${Math.min(maxRow, 10000)}`;
        }
      }

      // 按照官方文档：先激活工作表，然后使用 getActiveSheet()
      fWorkbook.setActiveSheet(targetSheet);
      const fWorksheet = fWorkbook.getActiveSheet();
      if (!fWorksheet) {
        continue;
      }

      const fRange = fWorksheet.getRange(range);
      if (!fRange) {
        continue;
      }

      // 创建筛选器，如果已存在则先移除
      let fFilter = fRange.createFilter();
      if (!fFilter) {
        const existingFilter = fWorksheet.getFilter();
        if (existingFilter) {
          existingFilter.remove();
          fFilter = fRange.createFilter();
        }
      }

      if (fFilter) {
        successCount++;
      }
    } catch (err) {
      console.error(`[fileImport] 添加筛选器失败:`, err);
    }
  }

  if (successCount > 0) {
    console.log(`[fileImport] 筛选器导入完成: ${successCount}/${totalCount}`);
  }
}

/**
 * Excel 图表类型到 Univer ChartType 的映射
 * 注意：Univer 的图表类型可以通过 univerAPI.Enum.ChartType 获取
 */
const EXCEL_TO_UNIVER_CHART_TYPE: Record<string, string> = {
  column: 'Column',
  bar: 'Bar',
  line: 'Line',
  area: 'Area',
  pie: 'Pie',
  doughnut: 'Doughnut',
  scatter: 'Scatter',
  radar: 'Radar',
  bubble: 'Bubble',
  combo: 'Combination',
  // 堆叠图表
  stackedBar: 'StackedBar',
  percentStackedBar: 'PercentStackedBar',
  stackedArea: 'StackedArea',
  percentStackedArea: 'PercentStackedArea',
  // 特殊图表类型
  wordCloud: 'WordCloud',
  funnel: 'Funnel',
  relationship: 'Relationship',
  waterfall: 'Waterfall',
  treemap: 'Treemap',
  sankey: 'Sankey',
  heatmap: 'Heatmap',
  boxPlot: 'BoxPlot',
  unknown: 'Column', // 默认使用柱状图
};

/**
 * 添加图表到工作簿
 * 使用 Univer Facade API (FWorksheet.newChart / insertChart)
 * @param univerAPI Univer API 实例
 * @param charts 图表数据，按 sheetId 分组
 */
export async function addChartsToWorkbook(
  univerAPI: any,
  charts: Record<string, ImportedChart[]>,
): Promise<void> {
  const fWorkbook = univerAPI.getActiveWorkbook();
  if (!fWorkbook) {
    return;
  }

  let successCount = 0;
  let totalCount = 0;

  for (const [sheetId, chartList] of Object.entries(charts)) {
    if (!chartList || chartList.length === 0) continue;
    totalCount += chartList.length;

    const targetSheet = fWorkbook.getSheetBySheetId(sheetId);
    if (!targetSheet) {
      continue;
    }

    // 按照官方文档：先激活工作表，然后使用 getActiveSheet()
    fWorkbook.setActiveSheet(targetSheet);
    const fWorksheet = fWorkbook.getActiveSheet();
    if (!fWorksheet) {
      continue;
    }

    for (const chartInfo of chartList) {
      try {
        const { chartType, dataRange, position, size, title, dataSheetName } = chartInfo;

        // 如果没有数据范围，跳过
        if (!dataRange) {
          console.warn(`[fileImport] 图表没有数据范围，跳过:`, chartInfo.chartId);
          continue;
        }

        // 获取 Univer 的图表类型枚举
        const univerChartTypeName = EXCEL_TO_UNIVER_CHART_TYPE[chartType] || 'Column';
        const univerChartType = univerAPI.Enum?.ChartType?.[univerChartTypeName];

        if (univerChartType === undefined) {
          console.warn(`[fileImport] 未找到图表类型: ${univerChartTypeName}`);
          continue;
        }

        // 创建图表构建器
        const chartBuilder = fWorksheet.newChart();
        if (!chartBuilder) {
          console.warn('[fileImport] 无法创建图表构建器');
          continue;
        }

        // 设置图表类型
        chartBuilder.setChartType(univerChartType);

        // 设置数据范围
        // 如果数据源在不同的工作表，需要添加工作表名称前缀
        const fullRange = dataSheetName ? `'${dataSheetName}'!${dataRange}` : dataRange;
        chartBuilder.addRange(fullRange);

        // 设置位置
        chartBuilder.setPosition(
          position.row,
          position.column,
          position.rowOffset,
          position.columnOffset,
        );

        // 设置尺寸
        chartBuilder.setWidth(size.width || 600);
        chartBuilder.setHeight(size.height || 400);

        // 设置标题（如果有）
        if (title) {
          chartBuilder.setOptions?.('title', { content: title });
        }

        // 构建并插入图表
        const chartBuildInfo = chartBuilder.build();
        await fWorksheet.insertChart(chartBuildInfo);

        successCount++;
      } catch (err) {
        console.error(`[fileImport] 添加图表失败:`, err);
      }
    }
  }

  if (successCount > 0) {
    console.log(`[fileImport] 图表导入完成: ${successCount}/${totalCount}`);
  }
}

/**
 * 添加图片到工作簿
 * 使用 Univer Facade API
 * @param univerAPI Univer API 实例
 * @param images 图片数据列表
 * @param options 图片插入选项
 */
export async function addImagesToWorkbook(
  univerAPI: any,
  images: ImportedImage[],
  options: ImageInsertOptions = {},
): Promise<{
  success: number;
  failed: number;
  errors: Array<{ image: ImportedImage; error: Error }>;
}> {
  if (!images || images.length === 0) {
    return { success: 0, failed: 0, errors: [] };
  }

  return insertImagesAfterImport(univerAPI, images, {
    defaultType: options.defaultType ?? ImageType.FLOATING,
    continueOnError: options.continueOnError ?? true,
    onProgress: options.onProgress
      ? (current, total, _img) => options.onProgress!(current, total)
      : undefined,
  });
}

/**
 * 添加透视表到工作簿
 * 使用 Univer Facade API (FWorkbook.addPivotTable)
 * @param univerAPI Univer API 实例
 * @param pivotTables 透视表数据列表
 */
export async function addPivotTablesToWorkbook(
  univerAPI: any,
  pivotTables: ImportedPivotTable[],
): Promise<void> {
  if (!pivotTables || pivotTables.length === 0) {
    return;
  }

  const fWorkbook = univerAPI.getActiveWorkbook();
  if (!fWorkbook) {
    return;
  }

  const unitId = fWorkbook.getId();
  let successCount = 0;

  for (const pivotTable of pivotTables) {
    try {
      const { sheetId, sourceRange, anchorCell, fields, name } = pivotTable;

      // 获取数据源工作表 - 优先使用 sheetId，如果没有则使用 sheetName
      let sourceSheet: any = null;
      if (sourceRange.sheetId) {
        sourceSheet = fWorkbook.getSheetBySheetId(sourceRange.sheetId);
      }
      if (!sourceSheet && sourceRange.sheetName) {
        sourceSheet = fWorkbook.getSheetByName(sourceRange.sheetName);
      }
      if (!sourceSheet) {
        console.warn(
          `[fileImport] 未找到数据源工作表: sheetId=${sourceRange.sheetId}, sheetName=${sourceRange.sheetName}`,
        );
        continue;
      }
      const sourceSheetId = sourceSheet.getSheetId();
      const sourceSheetName = sourceSheet.getSheetName();

      // 获取透视表所在工作表
      const targetSheet = fWorkbook.getSheetBySheetId(sheetId);
      if (!targetSheet) {
        console.warn(`[fileImport] 未找到透视表目标工作表: ${sheetId}`);
        continue;
      }
      const targetSheetId = targetSheet.getSheetId();

      // 先激活目标工作表
      fWorkbook.setActiveSheet(targetSheet);

      // 构建数据源信息（按照 Univer API 文档格式）
      const sourceInfo = {
        unitId,
        subUnitId: sourceSheetId,
        sheetName: sourceSheetName,
        range: {
          startRow: sourceRange.startRow,
          startColumn: sourceRange.startColumn,
          endRow: sourceRange.endRow,
          endColumn: sourceRange.endColumn,
        },
      };

      // 构建锚点信息（透视表放置位置）
      const anchorCellInfo = {
        unitId,
        subUnitId: targetSheetId,
        row: anchorCell.row,
        col: anchorCell.col,
      };

      // 清空目标区域的单元格数据，避免 Univer 弹出"目标区域中已有数据"的确认对话框
      // 优先使用 occupiedRange，如果没有则根据 anchorCell 估算一个合理的清空范围
      let clearStartRow = anchorCell.row;
      let clearStartCol = anchorCell.col;
      let clearEndRow = anchorCell.row + 50; // 默认清空 50 行
      let clearEndCol = anchorCell.col + 20; // 默认清空 20 列

      const { occupiedRange } = pivotTable;
      if (occupiedRange) {
        // 使用 occupiedRange，但扩展一些范围以确保完全清空
        clearStartRow = Math.max(0, occupiedRange.startRow - 2); // 向上扩展 2 行
        clearStartCol = Math.max(0, occupiedRange.startColumn - 2); // 向左扩展 2 列
        clearEndRow = occupiedRange.endRow + 10; // 向下扩展 10 行
        clearEndCol = occupiedRange.endColumn + 5; // 向右扩展 5 列
      }

      try {
        // 确保范围在工作表边界内
        const maxRows = targetSheet.getMaxRows?.() || 1000;
        const maxCols = targetSheet.getMaxColumns?.() || 1000;
        clearStartRow = Math.max(0, Math.min(clearStartRow, maxRows - 1));
        clearStartCol = Math.max(0, Math.min(clearStartCol, maxCols - 1));
        clearEndRow = Math.max(clearStartRow, Math.min(clearEndRow, maxRows - 1));
        clearEndCol = Math.max(clearStartCol, Math.min(clearEndCol, maxCols - 1));

        const numRows = clearEndRow - clearStartRow + 1;
        const numCols = clearEndCol - clearStartCol + 1;

        // 使用 getRange 获取范围对象
        const clearRange = targetSheet.getRange(clearStartRow, clearStartCol, numRows, numCols);
        if (clearRange) {
          // 方法1: 使用 setValues 清空值
          const emptyValues = Array(numRows)
            .fill(null)
            .map(() => Array(numCols).fill(null));
          clearRange.setValues(emptyValues);

          // 方法2: 尝试使用 clearContent 或类似方法清空单元格（如果 API 支持）
          try {
            // 尝试使用 clearContent 方法（如果存在）
            if (clearRange.clearContent && typeof clearRange.clearContent === 'function') {
              clearRange.clearContent();
            } else if (clearRange.clear && typeof clearRange.clear === 'function') {
              clearRange.clear();
            } else {
              // 如果 API 不支持批量清空，逐个清空单元格
              for (let r = clearStartRow; r <= clearEndRow; r++) {
                for (let c = clearStartCol; c <= clearEndCol; c++) {
                  try {
                    const cell = targetSheet.getRange(r, c, 1, 1);
                    if (cell) {
                      // 清空单元格值
                      cell.setValue(null);
                      // 如果支持清空格式，也清空格式
                      if (cell.clearFormat && typeof cell.clearFormat === 'function') {
                        cell.clearFormat();
                      } else if (cell.clear && typeof cell.clear === 'function') {
                        cell.clear();
                      }
                    }
                  } catch (cellError) {
                    // 单个单元格清空失败不影响整体流程
                  }
                }
              }
            }
          } catch (cellClearError) {
            // 清空格式失败不影响整体流程
            console.warn('[fileImport] 清空单元格格式失败:', cellClearError);
          }
        }
      } catch (clearError) {
        console.warn(`[fileImport] 清空透视表目标区域失败: ${clearError.message}`, {
          clearStartRow,
          clearStartCol,
          clearEndRow,
          clearEndCol,
        });
        // 清空失败不影响后续操作，但可能会弹出确认对话框
      }

      // 获取 PositionTypeEnum
      const PositionTypeEnum = univerAPI.Enum?.PositionTypeEnum;
      const PivotTableFiledAreaEnum = univerAPI.Enum?.PivotTableFiledAreaEnum;

      if (!PositionTypeEnum || !PivotTableFiledAreaEnum) {
        console.warn('[fileImport] 未找到透视表相关枚举类型，请确保已加载透视表插件');
        continue;
      }

      // 使用 Existing 模式，放在现有工作表中
      const positionType = PositionTypeEnum.Existing || 'existing';

      // 创建透视表
      let fPivotTable: any = null;
      try {
        fPivotTable = await fWorkbook.addPivotTable(sourceInfo, positionType, anchorCellInfo);
      } catch (addError) {
        continue;
      }

      // 如果返回 null，透视表可能已创建但 API 没有返回对象
      if (!fPivotTable) {
        successCount++;
        continue;
      }

      // 需要等待透视表渲染完成后再添加字段
      // 使用 Promise 包装事件监听
      await new Promise<void>((resolve, reject) => {
        let listenerDisposable: any = null;

        const timeoutId = setTimeout(() => {
          listenerDisposable?.dispose?.();
          reject(new Error('透视表渲染超时'));
        }, 10000); // 10秒超时

        listenerDisposable = univerAPI.addEvent(
          univerAPI.Event.PivotTableRendered,
          async (params: any) => {
            try {
              // 检查是否是当前透视表
              const currentPivotTableId = fPivotTable.getPivotTableId?.();
              if (params.pivotTableId === currentPivotTableId) {
                clearTimeout(timeoutId);
                listenerDisposable?.dispose?.();

                // 添加字段配置
                // 添加行字段
                for (let i = 0; i < fields.rowFields.length; i++) {
                  await fPivotTable.addField(fields.rowFields[i], PivotTableFiledAreaEnum.Row, i);
                }

                // 添加列字段
                for (let i = 0; i < fields.colFields.length; i++) {
                  await fPivotTable.addField(
                    fields.colFields[i],
                    PivotTableFiledAreaEnum.Column,
                    i,
                  );
                }

                // 添加值字段
                for (let i = 0; i < fields.valueFields.length; i++) {
                  await fPivotTable.addField(
                    fields.valueFields[i],
                    PivotTableFiledAreaEnum.Value,
                    i,
                  );
                }

                // 添加筛选字段
                for (let i = 0; i < fields.filterFields.length; i++) {
                  await fPivotTable.addField(
                    fields.filterFields[i],
                    PivotTableFiledAreaEnum.Filter,
                    i,
                  );
                }

                resolve();
              }
            } catch (err) {
              clearTimeout(timeoutId);
              listenerDisposable?.dispose?.();
              reject(err);
            }
          },
        );
      });

      successCount++;
    } catch (err) {
      console.error(`[fileImport] 添加透视表失败:`, err);
    }
  }

  if (successCount > 0) {
    console.log(`[fileImport] 透视表导入完成: ${successCount}/${pivotTables.length}`);
  }
}

/**
 * 使用 Univer Facade API 向 workbook 添加排序
 * 注意：Univer 的 sort API 是立即应用排序，而不是保存排序状态
 * 导入时会按照原 Excel 的排序配置重新执行排序
 *
 * @param univerAPI FUniver 实例
 * @param sorts 按 sheetId 组织的排序信息 Record<string, ImportedSort>
 */
export async function addSortsToWorkbook(
  univerAPI: any,
  sorts: Record<string, ImportedSort>,
): Promise<void> {
  if (!univerAPI || !sorts || Object.keys(sorts).length === 0) {
    return;
  }

  const workbook = univerAPI.getActiveWorkbook();
  if (!workbook) {
    console.warn('[fileImport] 无法获取活动 workbook');
    return;
  }

  let successCount = 0;

  for (const [sheetId, sortInfo] of Object.entries(sorts)) {
    try {
      // 通过 sheetId 获取工作表
      const fWorksheet = workbook.getSheetBySheetId(sheetId);
      if (!fWorksheet) {
        console.warn(`[fileImport] 未找到工作表: ${sheetId}`);
        continue;
      }

      // 解析排序范围
      const rangeStr = sortInfo.range;
      if (!rangeStr || sortInfo.conditions.length === 0) {
        continue;
      }

      // 获取 FRange 对象
      const fRange = fWorksheet.getRange(rangeStr);
      if (!fRange) {
        console.warn(`[fileImport] 无法获取范围: ${rangeStr}`);
        continue;
      }

      // 构建排序参数
      // Univer FRange.sort 支持:
      // - sort(column: number) - 升序排序
      // - sort({ column: number, ascending: boolean })
      // - sort([{ column: number, ascending: boolean }, ...])
      const sortCriteria = sortInfo.conditions.map((cond) => ({
        column: cond.column,
        ascending: cond.ascending,
      }));

      // 应用排序
      await fRange.sort(sortCriteria);
      successCount++;
    } catch (err) {
      console.error(`[fileImport] 添加排序失败:`, err);
    }
  }

  if (successCount > 0) {
    console.log(`[fileImport] 排序导入完成: ${successCount}/${Object.keys(sorts).length}`);
  }
}
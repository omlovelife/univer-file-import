/**
 * 工作簿辅助函数
 * 纯函数，无副作用
 */

import { LocaleType } from '@univerjs/core';
import type { IWorkbookData, IWorksheetData, ICellData } from '@univerjs/presets';
import { nanoid } from 'nanoid';

// 版本号（构建时注入）
declare const UNIVER_VERSION: string;

// 默认配置常量（统一管理）
export const DEFAULT_ROW_COUNT = 1000;
export const DEFAULT_COLUMN_COUNT = 26;
export const DEFAULT_COLUMN_WIDTH = 73;
export const DEFAULT_ROW_HEIGHT = 24;

// ========== 类型定义 ==========
export interface IWorkbookItem {
  documentId: string;
  workbookId: string;
  name: string;
  appId?: string;
  createTime?: string;
  modifyTime?: string;
  isUpdated?: boolean;
  source?: Array<{
    sheetId: string;
    type: number;
    id?: string;
    name?: string;
    externInfo?: Record<string, unknown>;
  }>;
  owner?: string;
}

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
 * 获取默认工作簿数据
 */
export function getDefaultWorkbookData(props: Partial<IWorkbookData> = {}): IWorkbookData {
  const sheetId = `sheet-${nanoid()}`;
  return {
    id: `workbook-${nanoid()}`,
    name: '未命名表格',
    appVersion: getUniverVersion(),
    locale: LocaleType.ZH_CN,
    styles: {},
    sheetOrder: [sheetId],
    sheets: {
      [sheetId]: {
        id: sheetId,
        name: 'Sheet1',
        cellData: {},
        columnCount: DEFAULT_COLUMN_COUNT,
        rowCount: DEFAULT_ROW_COUNT,
        defaultColumnWidth: DEFAULT_COLUMN_WIDTH,
        defaultRowHeight: DEFAULT_ROW_HEIGHT,
      } as IWorksheetData,
    },
    resources: [],
    ...props,
  };
}

/**
 * 生成基础 WorkbookData
 */
export function getWorkbookDataBySheets(sheets: IWorksheetData[], name?: string): IWorkbookData {
  const sheetOrder = sheets.map((s) => s.id);
  const sheetsMap: { [sheetId: string]: IWorksheetData } = {};
  sheets.forEach((s) => {
    sheetsMap[s.id] = s;
  });
  return {
    ...getDefaultWorkbookData({ name }),
    sheetOrder,
    sheets: sheetsMap,
  };
}

/**
 * 获取工作簿数据（优先使用后端快照数据 > 默认数据）
 */
export function getWorkbookData(params: {
  snapshotData?: IWorkbookData | null;
  name: string;
}): IWorkbookData {
  const { snapshotData, name } = params;
  return snapshotData || getDefaultWorkbookData({ name });
}

/**
 * 将二维数组转换成 IWorksheetData 的 cellData 结构（行、列均为数字索引）
 */
function arrayToCellData(data: unknown[][]): { [row: number]: { [col: number]: ICellData } } {
  const cellData: { [row: number]: { [col: number]: ICellData } } = {};

  data.forEach((row, rowIndex) => {
    if (!cellData[rowIndex]) cellData[rowIndex] = {};
    row.forEach((value, colIndex) => {
      // 仅在有值时填充 cell，避免生成庞大的稀疏表
      if (value !== undefined && value !== null && value !== '') {
        const cell: ICellData = { v: value as string | number | boolean };
        // 简单类型标记（可选）
        if (typeof value === 'number') {
          (cell as Record<string, unknown>).t = 2; // number
        } else if (typeof value === 'boolean') {
          (cell as Record<string, unknown>).t = 3; // boolean
        } else {
          (cell as Record<string, unknown>).t = 1; // string / others
        }
        cellData[rowIndex][colIndex] = cell;
      }
    });
  });

  return cellData;
}

/**
 * 计算二维数组的行列数（含默认兜底）
 */
function computeDimensions(rows: unknown[][]): { rowCount: number; columnCount: number } {
  const rowCount = Math.max(DEFAULT_ROW_COUNT, rows.length);
  const columnCount = Math.max(
    DEFAULT_COLUMN_COUNT,
    rows.reduce((max, row) => Math.max(max, row.length), 0),
  );
  return { rowCount, columnCount };
}

/**
 * 从二维数组构建 IWorksheetData
 */
export function buildSheetFrom2DArray(name: string, rows: unknown[][]): IWorksheetData {
  const { rowCount, columnCount } = computeDimensions(rows);
  const sheetId = `sheet-${nanoid()}`;
  return {
    id: sheetId,
    name,
    rowCount,
    columnCount,
    defaultColumnWidth: DEFAULT_COLUMN_WIDTH,
    defaultRowHeight: DEFAULT_ROW_HEIGHT,
    zoomRatio: 1,
    scrollTop: 0,
    scrollLeft: 0,
    hidden: 0,
    tabColor: '',
    cellData: arrayToCellData(rows),
  } as IWorksheetData;
}

/**
 * 规范化 source 为数组格式
 */
export function normalizeSource(source: unknown): unknown[] {
  if (!source) return [];
  if (Array.isArray(source)) return source;
  if (typeof source === 'object' && source !== null) return [source];
  return [];
}

/**
 * 构造工作簿项
 */
export function createWorkbookItem(params: {
  documentId: string;
  workbookId: string;
  name: string;
  docuMetaData?: Record<string, unknown>;
  source?: unknown;
  defaultAppId?: string;
  userName?: string;
}): IWorkbookItem {
  const { documentId, workbookId, name, docuMetaData, source, defaultAppId, userName } = params;

  let sourceInfo: unknown[] = [];
  if (docuMetaData?.source) {
    sourceInfo = normalizeSource(docuMetaData.source);
  } else if (source) {
    sourceInfo = normalizeSource(source);
  }

  const now = new Date().toISOString().replace('T', ' ').slice(0, 19);

  return {
    ...docuMetaData,
    documentId,
    workbookId,
    name,
    createTime: (docuMetaData?.createTime as string) || now,
    modifyTime: docuMetaData?.modifyTime as string | undefined,
    appId: (docuMetaData?.appId as string) || defaultAppId,
    owner: (docuMetaData?.owner as string) || userName,
    isUpdated: false,
    source: sourceInfo as IWorkbookItem['source'],
  };
}

# 📊 univer-file-import

<p align="center">
  <strong>Excel/CSV 文件导入工具，将 Excel 和 CSV 文件转换为 <a href="https://univer.ai/">Univer</a> 工作簿数据格式</strong>
</p>

<p align="center">
  <a href="#功能特性">功能特性</a> •
  <a href="#安装">安装</a> •
  <a href="#快速开始">快速开始</a> •
  <a href="#api-文档">API 文档</a> •
  <a href="#类型定义">类型定义</a>
</p>

## 📝 更新日志（Latest）

### 1.2.4 (2026-02-02)
- **csv兼容性**： 优化百分比显示。

### v1.2.3（2026-01-30）
- **图表导入增强**：补充支持堆叠/百分比堆叠条形图、堆叠/百分比堆叠面积图等类型识别。
- **公式兼容性**：导入时确保公式以 `=` 开头（符合 Univer 公式要求）。

更多历史版本请查看 `CHANGELOG.md`。

---

## ✨ 功能特性

<table>
<tr>
<td width="50%">

### 📁 文件格式
- ✅ Excel 文件 (.xlsx, .xls)
- ✅ CSV 文件 (.csv)

### 📝 工作表
- ✅ 保留所有工作表（包括空表）
- ✅ 处理特殊字符（>>>等）
- ✅ 保持工作表顺序

### 🎨 样式
- ✅ 字体、颜色、边框、对齐
- ✅ 条件格式
- ✅ 数据验证

</td>
<td width="50%">

### 📐 公式
- ✅ 公式和计算值保留
- ✅ TRANSPOSE 和数组公式
- ✅ 共享公式支持

### 🖼️ 富内容
- ✅ 图片（浮动/单元格）
- ✅ 超链接和富文本
- ✅ 合并单元格

### 📈 高级功能
- ✅ 图表（当前支持导入：柱状图/条形图、堆叠条形图、百分比堆叠条形图、折线图、面积图、堆叠面积图、百分比堆叠面积图、饼图、圆环图、散点图、雷达图、气泡图、组合图）
- ✅ 透视表
- ✅ 筛选器和排序

</td>
</tr>
</table>

---

## 📦 安装

```bash
# npm
npm install univer-file-import

# pnpm
pnpm add univer-file-import

# yarn
yarn add univer-file-import
```

### Peer Dependencies

```bash
npm install @univerjs/core @univerjs/presets
```

---

## 🚀 快速开始

### 基础用法

```typescript
import { 
  importFile, 
  ImageType,
  addConditionalFormatsToWorkbook,
  addFiltersToWorkbook,
  addSortsToWorkbook,
  addChartsToWorkbook,
  addPivotTablesToWorkbook,
  addImagesToWorkbook,
} from 'univer-file-import';

// 1️⃣ 导入文件
const result = await importFile(file);

// 2️⃣ 解构获取数据
const { 
  workbookData, 
  images, 
  conditionalFormats, 
  filters, 
  sorts, 
  charts, 
  pivotTables 
} = result;

// 3️⃣ 创建 Univer 工作簿
const workbook = univerAPI.createWorkbook(workbookData);

// 4️⃣ 添加附加功能（按需）
await addConditionalFormatsToWorkbook(univerAPI, conditionalFormats);
await addFiltersToWorkbook(univerAPI, filters);
await addSortsToWorkbook(univerAPI, sorts);
await addChartsToWorkbook(univerAPI, charts);
await addPivotTablesToWorkbook(univerAPI, pivotTables);
await addImagesToWorkbook(univerAPI, images);
```

### 使用辅助函数

```typescript
import {
  getDefaultWorkbookData,
  getWorkbookDataBySheets,
  buildSheetFrom2DArray,
} from 'univer-file-import';

// 从二维数组构建工作表
const sheet = buildSheetFrom2DArray('数据', [
  ['姓名', '年龄', '城市'],
  ['张三', 25, '北京'],
  ['李四', 30, '上海'],
]);

// 生成工作簿数据
const workbookData = getWorkbookDataBySheets([sheet], '我的工作簿');
```

---

## 📖 API 文档

### 核心函数

#### `importFile(file, options?)`

统一的文件导入接口，支持 Excel (.xlsx, .xls) 和 CSV (.csv) 文件。

```typescript
const result = await importFile(file, {
  includeImages: true,  // 是否包含图片，默认 true
});
```

**返回值 `FileImportResult`：**

| 属性 | 类型 | 说明 |
|------|------|------|
| `workbookData` | `IWorkbookData` | Univer 工作簿数据 |
| `images` | `ImportedImage[]` | 图片列表 |
| `conditionalFormats` | `Record<string, ImportedConditionalFormat[]>` | 条件格式（按 sheetId） |
| `filters` | `Record<string, ImportedFilter>` | 筛选器（按 sheetId） |
| `sorts` | `Record<string, ImportedSort>` | 排序（按 sheetId） |
| `charts` | `Record<string, ImportedChart[]>` | 图表（按 sheetId） |
| `pivotTables` | `ImportedPivotTable[]` | 透视表列表 |

---

### 添加功能函数

这些函数用于在工作簿创建后添加高级功能：

| 函数 | 说明 |
|------|------|
| `addConditionalFormatsToWorkbook(univerAPI, conditionalFormats)` | 添加条件格式 |
| `addFiltersToWorkbook(univerAPI, filters)` | 添加筛选器 |
| `addSortsToWorkbook(univerAPI, sorts)` | 添加排序 |
| `addChartsToWorkbook(univerAPI, charts)` | 添加图表 |
| `addPivotTablesToWorkbook(univerAPI, pivotTables)` | 添加透视表 |
| `addImagesToWorkbook(univerAPI, images)` | 添加图片 |

#### `addImagesToWorkbook` 选项

> `addImagesToWorkbook` 现在不再支持自定义图片类型、进度回调等参数，所有图片均以浮动图片方式插入，失败自动跳过。

---

### 辅助函数

| 函数 | 说明 |
|------|------|
| `getDefaultWorkbookData(props?)` | 获取默认工作簿数据 |
| `getWorkbookDataBySheets(sheets, name?)` | 从工作表数组生成工作簿 |
| `buildSheetFrom2DArray(name, rows)` | 从二维数组构建工作表 |
| `getWorkbookData({ snapshotData?, name })` | 获取工作簿数据 |

### 常量

| 常量 | 值 | 说明 |
|------|-----|------|
| `DEFAULT_ROW_COUNT` | 1000 | 默认行数 |
| `DEFAULT_COLUMN_COUNT` | 26 | 默认列数 |
| `DEFAULT_COLUMN_WIDTH` | 73 | 默认列宽 |
| `DEFAULT_ROW_HEIGHT` | 24 | 默认行高 |

### 枚举

```typescript
enum ImageType {
  FLOATING = 'floating',  // 浮动图片
  CELL = 'cell',          // 单元格图片
}
```

---

## 📝 类型定义

<details>
<summary><b>ImportedImage</b> - 图片信息</summary>

```typescript
interface ImportedImage {
  id: string;
  type: ImageType;
  source: string;
  sheetId: string;
  sheetName: string;
  position: {
    column: number;
    row: number;
    columnOffset: number;
    rowOffset: number;
  };
  size: {
    width: number;
    height: number;
  };
  endPosition?: { column: number; row: number; columnOffset: number; rowOffset: number };
  title?: string;
  description?: string;
}
```
</details>

<details>
<summary><b>ImportedChart</b> - 图表信息</summary>

```typescript
interface ImportedChart {
  chartId: string;
  sheetId: string;
  sheetName: string;
  chartType: 'column' | 'bar' | 'line' | 'area' | 'pie' | 'doughnut' | 'scatter' | 'radar' | 'bubble' | 'combo' | 'unknown';
  dataRange?: string;
  dataSheetName?: string;
  position: { row: number; column: number; rowOffset: number; columnOffset: number };
  size: { width: number; height: number };
  title?: string;
  rawData?: any;
}
```
</details>

<details>
<summary><b>ImportedPivotTable</b> - 透视表信息</summary>

```typescript
interface ImportedPivotTable {
  pivotTableId: string;
  sheetId: string;
  sheetName: string;
  sourceRange: {
    sheetName: string;
    startRow: number;
    startColumn: number;
    endRow: number;
    endColumn: number;
  };
  anchorCell: { row: number; col: number };
  occupiedRange?: { startRow: number; startColumn: number; endRow: number; endColumn: number };
  fields: {
    rowFields: number[];
    colFields: number[];
    valueFields: number[];
    filterFields: number[];
  };
  name?: string;
}
```
</details>

<details>
<summary><b>其他类型</b></summary>

```typescript
interface ImportedConditionalFormat {
  type: 'dataBar' | 'colorScale' | 'iconSet' | 'highlightCell' | 'other';
  ranges: string[];
  config: any;
  priority?: number;
  stopIfTrue?: boolean;
}

interface ImportedFilter {
  range: string;
}

interface ImportedSort {
  range: string;
  conditions: Array<{ column: number; ascending: boolean }>;
}

interface FileImportOptions {
  includeImages?: boolean;
}

interface ImageInsertOptions {
  defaultType?: ImageType;
  continueOnError?: boolean;
  onProgress?: (current: number, total: number) => void;
}

interface FileImportResult {
  workbookData: IWorkbookData;
  images: ImportedImage[];
  conditionalFormats: Record<string, ImportedConditionalFormat[]>;
  filters: Record<string, ImportedFilter>;
  sorts: Record<string, ImportedSort>;
  charts: Record<string, ImportedChart[]>;
  pivotTables: ImportedPivotTable[];
}
```
</details>

---

## 🌐 浏览器兼容性

此库设计用于浏览器环境，依赖以下 Web API：

- `File` / `Blob`
- `ArrayBuffer`
- `TextDecoder`
- `btoa`

---

## 📄 许可证

MIT

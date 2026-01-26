# univer-file-import

Excel/CSV 文件导入工具，用于将 Excel 和 CSV 文件转换为 [Univer](https://univer.ai/) 工作簿数据格式。

## 功能特性

- ✅ Excel 文件 (.xlsx, .xls)
- ✅ CSV 文件 (.csv)
- ✅ 保留所有工作表（包括空表）
- ✅ 处理工作表名称中的特殊字符
- ✅ 保持工作表顺序
- ✅ 完整样式保留（字体、颜色、边框、对齐）
- ✅ 公式和计算值保留（包括 TRANSPOSE 和数组公式）
- ✅ 共享公式支持
- ✅ 合并单元格支持
- ✅ 条件格式
- ✅ 数据验证
- ✅ 超链接和富文本
- ✅ 图片导入支持（浮动图片和单元格图片）

## 安装

```bash
npm install univer-file-import

# 或者使用 pnpm
pnpm add univer-file-import

# 或者使用 yarn
yarn add univer-file-import
```

### Peer Dependencies

此包需要以下 peer dependencies：

```bash
npm install @univerjs/core @univerjs/presets
```

## 使用方法

### 基础用法 - 导入 Excel 文件

```typescript
import { importFile, ImageType } from 'univer-file-import';

// 从文件输入获取文件
const fileInput = document.querySelector('input[type="file"]');
const file = fileInput.files[0];

// 导入文件
const result = await importFile(file, {
  includeImages: true, // 是否包含图片，默认 true
});

// 获取工作簿数据
const { workbookData, images } = result;

// 使用 workbookData 创建 Univer 工作簿
// ...

// 如果需要插入图片（在工作簿创建后）
if (images.length > 0) {
  const sheetIdMapping = new Map(); // 原始 sheetId -> 实际 sheetId 的映射
  // 填充 sheetIdMapping...
  
  const insertResult = await result.insertImages(univerAPI, sheetIdMapping, {
    defaultType: ImageType.FLOATING,
    continueOnError: true,
    onProgress: (current, total) => {
      console.log(`插入图片进度: ${current}/${total}`);
    },
  });
  
  console.log(`成功插入 ${insertResult.success} 张图片，失败 ${insertResult.failed} 张`);
}
```

### 导入 CSV 文件

```typescript
import { importCsv } from 'univer-file-import';

const workbookData = await importCsv(file, {
  delimiter: ',',        // 分隔符，默认自动检测
  encoding: 'UTF-8',     // 编码，默认 UTF-8
  hasHeader: true,       // 是否有表头，默认 true
  sheetName: 'Sheet1',   // 工作表名称
  skipEmptyLines: true,  // 跳过空行，默认 true
  trimValues: true,      // 修剪单元格值的空白，默认 true
});
```

### 使用辅助函数

```typescript
import {
  getDefaultWorkbookData,
  getWorkbookDataBySheets,
  buildSheetFrom2DArray,
  DEFAULT_ROW_COUNT,
  DEFAULT_COLUMN_COUNT,
} from 'univer-file-import';

// 获取默认工作簿数据
const defaultWorkbook = getDefaultWorkbookData({ name: '我的工作簿' });

// 从二维数组构建工作表
const rows = [
  ['姓名', '年龄', '城市'],
  ['张三', 25, '北京'],
  ['李四', 30, '上海'],
];
const sheet = buildSheetFrom2DArray('数据', rows);

// 从工作表数组生成工作簿数据
const workbookData = getWorkbookDataBySheets([sheet], '我的工作簿');
```

## API 文档

### `importFile(file, options?)`

统一的文件导入接口，支持 Excel 和 CSV 文件。

**参数:**

- `file: File` - 要导入的文件
- `options?: FileImportOptions`
  - `includeImages?: boolean` - 是否包含图片，默认 `true`

**返回值:** `Promise<FileImportResult>`

- `workbookData: IWorkbookData` - Univer 工作簿数据
- `images: ImportedImage[]` - 导入的图片列表
- `insertImages(univerAPI, sheetIdMapping, options?)` - 插入图片的方法

### `importCsv(file, options?)`

CSV 文件导入函数。

**参数:**

- `file: File | Blob` - CSV 文件
- `options?: CsvImportOptions`
  - `delimiter?: string` - 分隔符，默认自动检测
  - `encoding?: string` - 编码，默认 `'UTF-8'`
  - `hasHeader?: boolean` - 是否有表头，默认 `true`
  - `sheetName?: string` - 工作表名称，默认 `'Sheet1'`
  - `skipEmptyLines?: boolean` - 跳过空行，默认 `true`
  - `trimValues?: boolean` - 修剪空白，默认 `true`

**返回值:** `Promise<IWorkbookData>`

### `ImageType`

图片类型枚举：

- `ImageType.FLOATING` - 浮动图片
- `ImageType.CELL` - 单元格图片

### 辅助函数

- `getDefaultWorkbookData(props?)` - 获取默认工作簿数据
- `getWorkbookDataBySheets(sheets, name?)` - 从工作表数组生成工作簿数据
- `buildSheetFrom2DArray(name, rows)` - 从二维数组构建工作表
- `getWorkbookData({ snapshotData?, name })` - 获取工作簿数据

### 常量

- `DEFAULT_ROW_COUNT` - 默认行数 (1000)
- `DEFAULT_COLUMN_COUNT` - 默认列数 (26)
- `DEFAULT_COLUMN_WIDTH` - 默认列宽 (73)
- `DEFAULT_ROW_HEIGHT` - 默认行高 (24)

## 类型定义

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
  endPosition?: {
    column: number;
    row: number;
    columnOffset: number;
    rowOffset: number;
  };
  title?: string;
  description?: string;
}

type SheetIdMapping = Map<string, string>;

interface FileImportOptions {
  includeImages?: boolean;
}

interface ImageInsertOptions {
  defaultType?: ImageType;
  continueOnError?: boolean;
  onProgress?: (current: number, total: number) => void;
}
```

## 浏览器兼容性

此库设计用于浏览器环境，依赖以下 Web API：

- `File` / `Blob`
- `ArrayBuffer`
- `TextDecoder`
- `btoa`

## 许可证

MIT

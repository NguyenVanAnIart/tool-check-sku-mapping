import ExcelJS from 'exceljs';
import {
  canonicalizeHeader,
  excelCellValueToString,
  isFactorySheetHeaderRow,
  isLikelyFactorySheetName,
  toBooleanStock,
  toNumberOrNull,
} from './normalize.js';
import type { ExcelRow } from './types.js';

const HEADER_ROW_DEFAULT = 5;

const STREAM_READER_OPTIONS: Partial<ExcelJS.stream.xlsx.WorkbookStreamReaderOptions> = {
  sharedStrings: 'cache',
  hyperlinks: 'ignore',
  styles: 'ignore',
};

function worksheetName(reader: ExcelJS.stream.xlsx.WorksheetReader): string {
  return String((reader as { name?: string }).name ?? '').trim();
}

function buildColToField(headerCells: unknown[]): Map<number, string> {
  const colToField = new Map<number, string>();
  for (let i = 0; i < headerCells.length; i++) {
    const field = canonicalizeHeader(excelCellValueToString(headerCells[i]));
    if (field) colToField.set(i, field);
  }
  return colToField;
}

function getByField(
  cells: unknown[],
  colToField: Map<number, string>,
  field: string,
): unknown {
  for (const [colIndex, f] of colToField.entries()) {
    if (f === field) return cells[colIndex] ?? '';
  }
  return '';
}

function parseDataRow(
  cells: unknown[],
  colToField: Map<number, string>,
  rowNumber: number,
): ExcelRow | null | 'skip' {
  const str = (field: string) => excelCellValueToString(getByField(cells, colToField, field)).trim();

  const id = str('id');
  const sku = str('sku');
  const type = str('type');
  const printTech = str('print_tech');

  if (!id && !sku) return null;
  if (!sku || !type || !printTech) return 'skip';

  return {
    rowNumber,
    id,
    sku,
    productName: str('product_name') || null,
    color: str('color') || null,
    size: str('size') || null,
    type,
    printTech,
    include: str('include') || '',
    exclude: str('exclude') || '',
    tag: str('tag') || null,
    stock: toBooleanStock(getByField(cells, colToField, 'stock')),
    price: toNumberOrNull(getByField(cells, colToField, 'price')) ?? 0,
    flashSale1: toNumberOrNull(getByField(cells, colToField, 'flash_sale_1')),
    flashSale2: toNumberOrNull(getByField(cells, colToField, 'flash_sale_2')),
    additionalPrice: toNumberOrNull(
      getByField(cells, colToField, 'additional_price'),
    ),
  };
}

function rowToCells(row: ExcelJS.Row): unknown[] {
  const cells: unknown[] = [];
  row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
    while (cells.length < colNumber - 1) cells.push('');
    cells[colNumber - 1] = excelCellValueToString(cell.value);
  });
  return cells;
}

async function drainWorksheetRows(
  worksheetReader: ExcelJS.stream.xlsx.WorksheetReader,
): Promise<void> {
  for await (const _row of worksheetReader) {
    // Bỏ qua nội dung sheet không cần parse để giải phóng stream.
  }
}

export interface ParsedFactorySheet {
  factory: string;
  rows: ExcelRow[];
  skippedRows: number;
  headerRow: number;
}

async function parseWorksheetReader(
  worksheetReader: ExcelJS.stream.xlsx.WorksheetReader,
  preferredHeaderRow: number,
): Promise<ParsedFactorySheet | null> {
  let headerRowIndex: number | null = null;
  let colToField: Map<number, string> | null = null;
  const rows: ExcelRow[] = [];
  let skippedRows = 0;
  let drainOnly = false;

  for await (const row of worksheetReader) {
    if (drainOnly) continue;

    const rowNumber = row.number;
    const cells = rowToCells(row);

    if (headerRowIndex === null) {
      if (rowNumber > 20) {
        drainOnly = true;
        continue;
      }

      if (rowNumber === preferredHeaderRow && isFactorySheetHeaderRow(cells)) {
        headerRowIndex = rowNumber;
        colToField = buildColToField(cells);
        continue;
      }

      if (rowNumber !== preferredHeaderRow && isFactorySheetHeaderRow(cells)) {
        headerRowIndex = rowNumber;
        colToField = buildColToField(cells);
        continue;
      }

      continue;
    }

    if (rowNumber <= headerRowIndex || !colToField) continue;

    const parsedRow = parseDataRow(cells, colToField, rowNumber);
    if (parsedRow === null) continue;
    if (parsedRow === 'skip') {
      skippedRows++;
      continue;
    }
    rows.push(parsedRow);
  }

  if (headerRowIndex === null || colToField === null) return null;

  return {
    factory: worksheetName(worksheetReader).toLowerCase(),
    rows,
    skippedRows,
    headerRow: headerRowIndex,
  };
}

export async function parseFactorySheets(
  filePath: string,
  headerRow = HEADER_ROW_DEFAULT,
  factories?: string[],
): Promise<ParsedFactorySheet[]> {
  const wantedFactories = factories?.length
    ? new Set(factories.map((f) => f.trim().toLowerCase()))
    : null;

  const parsed: ParsedFactorySheet[] = [];
  const reader = new ExcelJS.stream.xlsx.WorkbookReader(filePath, STREAM_READER_OPTIONS);

  for await (const worksheetReader of reader) {
    const sheetName = worksheetName(worksheetReader);
    const factoryKey = sheetName.toLowerCase();
    const shouldParse =
      isLikelyFactorySheetName(sheetName) &&
      (!wantedFactories || wantedFactories.has(factoryKey));

    if (!shouldParse) {
      await drainWorksheetRows(worksheetReader);
      continue;
    }

    const sheet = await parseWorksheetReader(worksheetReader, headerRow);
    if (sheet) parsed.push(sheet);
  }

  return parsed;
}

export async function listSheetNames(filePath: string): Promise<string[]> {
  const names: string[] = [];
  const reader = new ExcelJS.stream.xlsx.WorkbookReader(filePath, STREAM_READER_OPTIONS);

  for await (const worksheetReader of reader) {
    names.push(worksheetName(worksheetReader));
    await drainWorksheetRows(worksheetReader);
  }

  return names;
}

export async function listFactorySheetNames(filePath: string): Promise<string[]> {
  const names: string[] = [];
  const reader = new ExcelJS.stream.xlsx.WorkbookReader(filePath, STREAM_READER_OPTIONS);

  for await (const worksheetReader of reader) {
    const sheetName = worksheetName(worksheetReader);
    if (isLikelyFactorySheetName(sheetName)) names.push(sheetName);
    await drainWorksheetRows(worksheetReader);
  }

  return names;
}

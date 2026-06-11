import ExcelJS from 'exceljs';
import fs from 'node:fs';
import path from 'node:path';
import type { CompareReport, CompareResult, DbRow, ExcelRow, SheetSummary } from './types.js';

const COMMON_HEADERS = [
  'Xưởng',
  'Dòng Excel',
  'ID Excel',
  'SKU',
  'Type',
  'PrintTech',
  'Include',
] as const;

const DETAIL_HEADERS = [
  'Product Name',
  'Color',
  'Size',
  'Price',
  'Stock',
  'Flash Sale 1',
  'Flash Sale 2',
  'Additional Price',
] as const;

const EXCEL_DETAIL_HEADERS = DETAIL_HEADERS.map((h) => `${h} (Excel)`) as readonly string[];
const DB_DETAIL_HEADERS = DETAIL_HEADERS.map((h) => `${h} (DB)`) as readonly string[];

function excelDetailValues(row: ExcelRow): (string | number | boolean)[] {
  return [
    row.productName ?? '',
    row.color ?? '',
    row.size ?? '',
    row.price,
    row.stock,
    row.flashSale1 ?? '',
    row.flashSale2 ?? '',
    row.additionalPrice ?? '',
  ];
}

function dbDetailValues(row: DbRow): (string | number | boolean)[] {
  return [
    row.product_name ?? '',
    row.color ?? '',
    row.size ?? '',
    row.price,
    row.stock,
    row.flash_sale1 ?? '',
    row.flash_sale2 ?? '',
    row.additional_price ?? '',
  ];
}

function normalizeFactory(factory: string): string {
  return factory.trim().toLowerCase();
}

function sumSheetSummaries(sheets: SheetSummary[]): CompareReport['totals'] {
  return sheets.reduce(
    (acc, s) => ({
      totalExcelRows: acc.totalExcelRows + s.totalExcelRows,
      skippedRows: acc.skippedRows + s.skippedRows,
      matched: acc.matched + s.matched,
      mismatched: acc.mismatched + s.mismatched,
      notInDb: acc.notInDb + s.notInDb,
      onlyInDb: acc.onlyInDb + s.onlyInDb,
    }),
    {
      totalExcelRows: 0,
      skippedRows: 0,
      matched: 0,
      mismatched: 0,
      notInDb: 0,
      onlyInDb: 0,
    },
  );
}

export function filterReportByFactory(
  report: CompareReport,
  factory: string,
): CompareReport {
  const wanted = normalizeFactory(factory);
  const sheets = report.sheets.filter((s) => normalizeFactory(s.factory) === wanted);
  const results = report.results.filter((r) => normalizeFactory(r.factory) === wanted);
  const onlyInDb = report.onlyInDb.filter(
    (entry) => normalizeFactory(entry.factory) === wanted,
  );

  return {
    ...report,
    sheets,
    results,
    onlyInDb,
    totals: sumSheetSummaries(sheets),
  };
}

const DEFAULT_HEADER_FILL: ExcelJS.Fill = {
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb: 'FF4472C4' },
};

const EXCEL_HEADER_FILL: ExcelJS.Fill = {
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb: 'FF70AD47' },
};

const DB_HEADER_FILL: ExcelJS.Fill = {
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb: 'FFED7D31' },
};

const HEADER_FONT: Partial<ExcelJS.Font> = {
  bold: true,
  color: { argb: 'FFFFFFFF' },
};

interface HeaderZone {
  from: number;
  to: number;
  fill: ExcelJS.Fill;
}

function styleHeaderRow(
  sheet: ExcelJS.Worksheet,
  rowNumber = 1,
  zones?: HeaderZone[],
): void {
  const row = sheet.getRow(rowNumber);
  row.eachCell((cell, colNumber) => {
    const zone = zones?.find((z) => colNumber >= z.from && colNumber <= z.to);
    cell.fill = zone?.fill ?? DEFAULT_HEADER_FILL;
    cell.font = HEADER_FONT;
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
  });
  row.commit();
}

function autoWidth(sheet: ExcelJS.Worksheet, min = 10, max = 50): void {
  sheet.columns.forEach((col) => {
    let width = min;
    if (col.eachCell) {
      col.eachCell({ includeEmpty: false }, (cell) => {
        const len = String(cell.value ?? '').length;
        width = Math.min(max, Math.max(width, len + 2));
      });
    }
    col.width = width;
  });
}

function addSummarySheet(workbook: ExcelJS.Workbook, report: CompareReport): void {
  const sheet = workbook.addWorksheet('Tổng hợp');
  sheet.addRow(['Báo cáo so sánh SKU Price']);
  sheet.addRow(['File Excel', report.sourceFile]);
  sheet.addRow(['Thời gian', report.comparedAt]);
  sheet.addRow([]);

  sheet.addRow([
    'Xưởng',
    'Dòng Excel',
    'Bỏ qua',
    'Khớp',
    'Sai khác',
    'Không có trong DB',
    'Chỉ có trong DB',
  ]);
  styleHeaderRow(sheet, 6);

  for (const s of report.sheets) {
    sheet.addRow([
      s.factory.toUpperCase(),
      s.totalExcelRows,
      s.skippedRows,
      s.matched,
      s.mismatched,
      s.notInDb,
      s.onlyInDb,
    ]);
  }

  sheet.addRow([]);
  sheet.addRow([
    'TỔNG',
    report.totals.totalExcelRows,
    report.totals.skippedRows,
    report.totals.matched,
    report.totals.mismatched,
    report.totals.notInDb,
    report.totals.onlyInDb,
  ]);

  const totalRow = sheet.lastRow;
  if (totalRow) totalRow.font = { bold: true };

  autoWidth(sheet);
}

function addMatchSheet(workbook: ExcelJS.Workbook, results: CompareResult[]): void {
  const sheet = workbook.addWorksheet('Khớp');
  const headers = [
    ...COMMON_HEADERS,
    'DB ID',
    'sku_ecb',
    ...DETAIL_HEADERS,
  ];
  sheet.addRow(headers);
  styleHeaderRow(sheet);

  for (const r of results) {
    sheet.addRow([
      r.factory.toUpperCase(),
      r.excelRow.rowNumber,
      r.excelRow.id,
      r.excelRow.sku,
      r.excelRow.type,
      r.excelRow.printTech,
      r.excelRow.include,
      r.dbRow?.id ?? '',
      r.dbRow?.sku_ecb ?? '',
      ...excelDetailValues(r.excelRow),
    ]);
  }

  autoWidth(sheet);
}

function addNotInDbSheet(workbook: ExcelJS.Workbook, results: CompareResult[]): void {
  const sheet = workbook.addWorksheet('Không có trong DB');
  const headers = [...COMMON_HEADERS, ...DETAIL_HEADERS];
  sheet.addRow(headers);
  styleHeaderRow(sheet);

  for (const r of results) {
    sheet.addRow([
      r.factory.toUpperCase(),
      r.excelRow.rowNumber,
      r.excelRow.id,
      r.excelRow.sku,
      r.excelRow.type,
      r.excelRow.printTech,
      r.excelRow.include,
      ...excelDetailValues(r.excelRow),
    ]);
  }

  autoWidth(sheet);
}

function addMismatchSheet(workbook: ExcelJS.Workbook, results: CompareResult[]): void {
  const sheet = workbook.addWorksheet('Sai khác');
  const headers = [
    ...COMMON_HEADERS,
    ...EXCEL_DETAIL_HEADERS,
    'DB ID',
    'sku_ecb',
    ...DB_DETAIL_HEADERS,
    'Trường sai',
    'Giá trị Excel',
    'Giá trị DB',
  ];
  sheet.addRow(headers);

  const commonEnd = COMMON_HEADERS.length;
  const excelStart = commonEnd + 1;
  const excelEnd = excelStart + EXCEL_DETAIL_HEADERS.length - 1;
  const dbMetaStart = excelEnd + 1;
  const dbMetaEnd = dbMetaStart + 1;
  const dbDetailStart = dbMetaEnd + 1;
  const dbDetailEnd = dbDetailStart + DB_DETAIL_HEADERS.length - 1;

  styleHeaderRow(sheet, 1, [
    { from: excelStart, to: excelEnd, fill: EXCEL_HEADER_FILL },
    { from: dbMetaStart, to: dbDetailEnd, fill: DB_HEADER_FILL },
  ]);

  const baseColumnCount =
    COMMON_HEADERS.length +
    EXCEL_DETAIL_HEADERS.length +
    2 +
    DB_DETAIL_HEADERS.length;

  for (const r of results) {
    const baseValues = [
      r.factory.toUpperCase(),
      r.excelRow.rowNumber,
      r.excelRow.id,
      r.excelRow.sku,
      r.excelRow.type,
      r.excelRow.printTech,
      r.excelRow.include,
      ...excelDetailValues(r.excelRow),
      r.dbRow?.id ?? '',
      r.dbRow?.sku_ecb ?? '',
      ...dbDetailValues(r.dbRow!),
    ];

    for (let i = 0; i < r.diffs.length; i++) {
      const d = r.diffs[i];
      if (i === 0) {
        sheet.addRow([...baseValues, d.field, d.excelValue, d.dbValue]);
      } else {
        sheet.addRow([
          ...Array(baseColumnCount).fill(''),
          d.field,
          d.excelValue,
          d.dbValue,
        ]);
      }
    }
  }

  autoWidth(sheet);
}

function addOnlyInDbSheet(workbook: ExcelJS.Workbook, report: CompareReport): void {
  const sheet = workbook.addWorksheet('Chỉ có trong DB');
  const headers = [
    'Xưởng',
    'DB ID',
    'sku_ecb',
    'sku_factory',
    'type',
    'print_tech',
    'include',
    ...DETAIL_HEADERS,
  ];
  sheet.addRow(headers);
  styleHeaderRow(sheet, 1, [{ from: 1, to: headers.length, fill: DB_HEADER_FILL }]);

  for (const entry of report.onlyInDb) {
    const r = entry.dbRow;
    sheet.addRow([
      r.factory.toUpperCase(),
      r.id,
      r.sku_ecb,
      r.sku_factory,
      r.type,
      r.print_tech,
      r.include ?? '',
      ...dbDetailValues(r),
    ]);
  }

  autoWidth(sheet);
}

export async function writeReport(
  report: CompareReport,
  outputPath: string,
): Promise<string> {
  const dir = path.dirname(outputPath);
  fs.mkdirSync(dir, { recursive: true });

  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'sku-price-compare-tool';
  workbook.created = new Date();

  addSummarySheet(workbook, report);
  addMatchSheet(workbook, report.results.filter((r) => r.status === 'match'));
  addMismatchSheet(workbook, report.results.filter((r) => r.status === 'mismatch'));
  addNotInDbSheet(workbook, report.results.filter((r) => r.status === 'not_in_db'));
  addOnlyInDbSheet(workbook, report);

  await workbook.xlsx.writeFile(outputPath);
  return outputPath;
}

export function printConsoleSummary(report: CompareReport): void {
  console.log('\n========== BÁO CÁO SO SÁNH SKU PRICE ==========\n');
  console.log(`File:     ${report.sourceFile}`);
  console.log(`Thời gian: ${report.comparedAt}\n`);

  console.log(
    'Xưởng'.padEnd(12) +
      'Excel'.padStart(8) +
      'Bỏ qua'.padStart(8) +
      'Khớp'.padStart(8) +
      'Sai'.padStart(8) +
      'Thiếu DB'.padStart(10) +
      'Thừa DB'.padStart(10),
  );
  console.log('-'.repeat(64));

  for (const s of report.sheets) {
    console.log(
      s.factory.toUpperCase().padEnd(12) +
        String(s.totalExcelRows).padStart(8) +
        String(s.skippedRows).padStart(8) +
        String(s.matched).padStart(8) +
        String(s.mismatched).padStart(8) +
        String(s.notInDb).padStart(10) +
        String(s.onlyInDb).padStart(10),
    );
  }

  console.log('-'.repeat(64));
  const t = report.totals;
  console.log(
    'TỔNG'.padEnd(12) +
      String(t.totalExcelRows).padStart(8) +
      String(t.skippedRows).padStart(8) +
      String(t.matched).padStart(8) +
      String(t.mismatched).padStart(8) +
      String(t.notInDb).padStart(10) +
      String(t.onlyInDb).padStart(10),
  );

  const matchRate =
    t.totalExcelRows > 0
      ? ((t.matched / t.totalExcelRows) * 100).toFixed(1)
      : '0.0';
  console.log(`\nTỷ lệ khớp: ${matchRate}% (${t.matched}/${t.totalExcelRows})`);
}

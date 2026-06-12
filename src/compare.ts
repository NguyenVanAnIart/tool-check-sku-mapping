import {
  buildMatchKey,
  formatDisplayValue,
  normString,
  numbersEqual,
} from './normalize.js';
import type { ParsedFactorySheet } from './excel-parser.js';
import { resolveDtgPrintSkus } from './sku-mapping.js';
import type {
  CompareReport,
  CompareResult,
  DbRow,
  ExcelRow,
  FieldDiff,
  SheetSummary,
  SkuMappingIndex,
} from './types.js';
import { COMPARE_FIELDS } from './types.js';

function getExcelFieldValue(row: ExcelRow, field: string): unknown {
  return (row as unknown as Record<string, unknown>)[field];
}

function getDbFieldValue(row: DbRow, field: string): unknown {
  return (row as unknown as Record<string, unknown>)[field];
}

function compareRow(excelRow: ExcelRow, dbRow: DbRow): FieldDiff[] {
  const diffs: FieldDiff[] = [];

  for (const { excel, db, label } of COMPARE_FIELDS) {
    const excelVal = getExcelFieldValue(excelRow, excel);
    const dbVal = getDbFieldValue(dbRow, db);

    let equal = false;

    if (excel === 'stock') {
      equal = Boolean(excelVal) === Boolean(dbVal);
    } else if (
      excel === 'price' ||
      excel === 'flashSale1' ||
      excel === 'flashSale2' ||
      excel === 'additionalPrice'
    ) {
      const dbNum =
        dbVal === null || dbVal === undefined || dbVal === ''
          ? null
          : Number(dbVal);
      equal = numbersEqual(excelVal as number | null, dbNum);
    } else if (excel === 'include' || excel === 'exclude') {
      equal =
        normString(excelVal, false) === normString(dbVal ?? '', false);
    } else {
      equal =
        normString(excelVal, true) === normString(dbVal, true);
    }

    if (!equal) {
      diffs.push({
        field: label,
        excelValue: formatDisplayValue(excelVal),
        dbValue: formatDisplayValue(dbVal),
      });
    }
  }

  return diffs;
}

export function compareSheetsWithDb(
  sheets: ParsedFactorySheet[],
  dbRows: DbRow[],
  sourceFile: string,
  skuMapping: SkuMappingIndex,
): CompareReport {
  const dbByKey = new Map<string, DbRow>();
  for (const row of dbRows) {
    const key = buildMatchKey(
      row.factory,
      row.sku_factory,
      row.type,
      row.print_tech,
      row.include ?? '',
      row.sku_ecb ?? '',
    );
    dbByKey.set(key, row);
  }

  const results: CompareResult[] = [];
  const matchedKeys = new Set<string>();
  const sheetSummaries: SheetSummary[] = [];

  for (const sheet of sheets) {
    let matched = 0;
    let mismatched = 0;
    let notInDb = 0;

    for (const excelRow of sheet.rows) {
      const dtgPrintSkus = resolveDtgPrintSkus(skuMapping, sheet.factory, excelRow.id);

      if (dtgPrintSkus.length === 0) {
        notInDb++;
        results.push({
          factory: sheet.factory,
          excelRow,
          dbRow: null,
          status: 'not_in_db',
          diffs: [],
          matchKey: buildMatchKey(
            sheet.factory,
            excelRow.sku,
            excelRow.type,
            excelRow.printTech,
            excelRow.include,
            '',
          ),
        });
        continue;
      }

      let dbRow: DbRow | null = null;
      let matchKey = '';

      for (const dtgPrintSku of dtgPrintSkus) {
        const key = buildMatchKey(
          sheet.factory,
          excelRow.sku,
          excelRow.type,
          excelRow.printTech,
          excelRow.include,
          dtgPrintSku,
        );
        const found = dbByKey.get(key);
        if (found) {
          dbRow = found;
          matchKey = key;
          break;
        }
      }

      if (!dbRow) {
        notInDb++;
        results.push({
          factory: sheet.factory,
          excelRow,
          dbRow: null,
          status: 'not_in_db',
          diffs: [],
          matchKey: buildMatchKey(
            sheet.factory,
            excelRow.sku,
            excelRow.type,
            excelRow.printTech,
            excelRow.include,
            dtgPrintSkus[0],
          ),
        });
        continue;
      }

      matchedKeys.add(matchKey);
      const diffs = compareRow(excelRow, dbRow);

      if (diffs.length === 0) {
        matched++;
        results.push({
          factory: sheet.factory,
          excelRow,
          dbRow,
          status: 'match',
          diffs: [],
          matchKey,
        });
      } else {
        mismatched++;
        results.push({
          factory: sheet.factory,
          excelRow,
          dbRow,
          status: 'mismatch',
          diffs,
          matchKey,
        });
      }
    }

    const factoryDbRows = dbRows.filter(
      (r) => r.factory.toLowerCase() === sheet.factory.toLowerCase(),
    );
    const onlyInDb = factoryDbRows.filter((r) => {
      const key = buildMatchKey(
        r.factory,
        r.sku_factory,
        r.type,
        r.print_tech,
        r.include ?? '',
        r.sku_ecb ?? '',
      );
      return !matchedKeys.has(key);
    }).length;

    sheetSummaries.push({
      factory: sheet.factory,
      totalExcelRows: sheet.rows.length,
      skippedRows: sheet.skippedRows,
      matched,
      mismatched,
      notInDb,
      onlyInDb,
    });
  }

  const onlyInDbEntries: CompareReport['onlyInDb'] = [];
  for (const row of dbRows) {
    const key = buildMatchKey(
      row.factory,
      row.sku_factory,
      row.type,
      row.print_tech,
      row.include ?? '',
      row.sku_ecb ?? '',
    );
    if (!matchedKeys.has(key)) {
      const inExcelFactories = sheets.some(
        (s) => s.factory.toLowerCase() === row.factory.toLowerCase(),
      );
      if (inExcelFactories) {
        onlyInDbEntries.push({ factory: row.factory, dbRow: row, matchKey: key });
      }
    }
  }

  const totals = sheetSummaries.reduce(
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

  return {
    sourceFile,
    comparedAt: new Date().toISOString(),
    sheets: sheetSummaries,
    results,
    onlyInDb: onlyInDbEntries,
    totals,
  };
}

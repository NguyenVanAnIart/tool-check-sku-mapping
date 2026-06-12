import fs from 'node:fs';
import path from 'node:path';
import { compareSheetsWithDb } from './compare.js';
import {
  type DbConfig,
  fetchSkuPriceByFactories,
  loadDbConfigFromEnv,
  testConnection,
} from './db.js';
import {
  listFactorySheetNames,
  listSheetNames,
  parseFactorySheets,
} from './excel-parser.js';
import { filterReportByFactory, writeReport } from './report.js';
import { parseSkuMapping } from './sku-mapping.js';
import type { CompareReport } from './types.js';

export interface RunCompareOptions {
  filePath: string;
  outputPath?: string;
  headerRow: number;
  factory?: string;
  sheets?: string[];
  dbConfig?: DbConfig;
}

export interface RunCompareResult {
  report: CompareReport;
  outputPath: string;
}

function filterSheetsByFactories<T extends { factory: string }>(
  sheets: T[],
  factories: Set<string>,
): T[] {
  return sheets.filter((s) => factories.has(s.factory.toLowerCase()));
}

export async function runCompare(options: RunCompareOptions): Promise<RunCompareResult> {
  const filePath = path.resolve(options.filePath);
  if (!fs.existsSync(filePath)) {
    throw new Error(`Không tìm thấy file: ${filePath}`);
  }

  const dbConfig = options.dbConfig ?? loadDbConfigFromEnv();
  await testConnection(dbConfig);

  const wantedFactories = new Set<string>();
  if (options.factory) wantedFactories.add(options.factory.trim().toLowerCase());
  if (options.sheets?.length) {
    for (const sheet of options.sheets) {
      wantedFactories.add(sheet.trim().toLowerCase());
    }
  }
  const factoryFilter =
    wantedFactories.size > 0 ? [...wantedFactories] : undefined;

  const allFactorySheets = await parseFactorySheets(
    filePath,
    options.headerRow,
    factoryFilter,
  );

  let sheets = allFactorySheets;
  if (wantedFactories.size > 0) {
    sheets = filterSheetsByFactories(allFactorySheets, wantedFactories);
    if (sheets.length === 0) {
      const available = (await listFactorySheetNames(filePath))
        .map((name) => name.toUpperCase())
        .join(', ');
      const requested = [...wantedFactories].map((f) => f.toUpperCase()).join(', ');
      throw new Error(
        `Không tìm thấy xưởng "${requested}" trong file Excel.${available ? ` Các xưởng có trong file: ${available}` : ''}`,
      );
    }
  }

  if (sheets.length === 0) {
    throw new Error(
      'Không tìm thấy sheet xưởng hợp lệ (LK401, LK1, ...). Kiểm tra tên sheet và dòng header.',
    );
  }

  const factories = [...new Set(sheets.map((s) => s.factory))];
  const dbRows = await fetchSkuPriceByFactories(dbConfig, factories);
  const skuMapping = await parseSkuMapping(filePath);

  let report = compareSheetsWithDb(sheets, dbRows, filePath, skuMapping);
  if (options.factory) {
    report = filterReportByFactory(report, options.factory);
  }

  const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
  const defaultOutput = path.resolve(`./reports/report-${timestamp}.xlsx`);
  const outputPath = path.resolve(options.outputPath ?? defaultOutput);

  const written = await writeReport(report, outputPath);
  return { report, outputPath: written };
}

export function getDbInfo(): { host: string; port: number; database: string } {
  const cfg = loadDbConfigFromEnv();
  return { host: cfg.host, port: cfg.port, database: cfg.database };
}

export async function testDbConnection(dbConfig?: DbConfig): Promise<void> {
  await testConnection(dbConfig ?? loadDbConfigFromEnv());
}

export async function inspectExcelFile(filePath: string): Promise<{
  allSheets: string[];
  factorySheets: string[];
}> {
  const resolved = path.resolve(filePath);
  if (!fs.existsSync(resolved)) {
    throw new Error(`Không tìm thấy file: ${resolved}`);
  }
  const allSheets = await listSheetNames(resolved);
  const factorySheets = await listFactorySheetNames(resolved);
  return { allSheets, factorySheets };
}

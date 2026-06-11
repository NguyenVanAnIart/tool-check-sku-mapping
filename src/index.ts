import 'dotenv/config';
import path from 'node:path';
import { runCompare } from './compare-runner.js';
import { printConsoleSummary } from './report.js';

function printUsage(): void {
  console.log(`
SKU Price Compare Tool
So sánh file Excel xưởng (LK401, LK1, ...) với bảng sku_price PostgreSQL.

Cách dùng:
  npm run compare -- <file.xlsx> [options]
  npm run ui              Giao diện web (mở trình duyệt)

Options:
  --output, -o <path>     File báo cáo xlsx (mặc định: ./reports/report-<timestamp>.xlsx)
  --header-row <n>        Dòng header (mặc định: 5)
  <xưởng>                 Chỉ so sánh 1 xưởng (tham số thứ 2, khuyến nghị trên Windows)
  --factory, -f <name>    Chỉ so sánh 1 xưởng (dùng khi chạy trực tiếp: npx tsx ...)
  --sheet <name>          Chỉ so sánh sheet cụ thể (có thể lặp lại)
  --help, -h              Hiển thị trợ giúp

Cấu hình DB qua file .env (copy từ .env.example):
  DB_HOST, DB_PORT, DB_USERNAME, DB_PASSWORD, DB_NAME

Ví dụ:
  npm run compare -- ./data/pricing.xlsx
  npm run compare -- ./data/pricing.xlsx -o ./reports/result.xlsx
  npm run compare -- ./data/pricing.xlsx LK401
  npx tsx src/index.ts ./data/pricing.xlsx --factory LK401
  npm run compare -- ./data/pricing.xlsx --sheet LK401 --sheet LK1
`);
}

interface CliOptions {
  filePath: string;
  outputPath?: string;
  headerRow: number;
  factory?: string;
  sheets?: string[];
}

function parseArgs(argv: string[]): CliOptions | 'help' | null {
  if (argv.length === 0 || argv.includes('--help') || argv.includes('-h')) {
    return 'help';
  }

  let outputPath: string | undefined;
  let headerRow = 5;
  let factory: string | undefined;
  const sheets: string[] = [];
  const positionals: string[] = [];

  for (let i = 0; i < argv.length; i++) {
    const arg = argv[i];
    const factoryEq = arg.match(/^--factory=(.+)$/i);
    if (factoryEq) {
      factory = factoryEq[1];
      continue;
    }
    const shortFactoryEq = arg.match(/^-f=(.+)$/);
    if (shortFactoryEq) {
      factory = shortFactoryEq[1];
      continue;
    }
    if (arg === '--output' || arg === '-o') {
      outputPath = argv[++i];
      continue;
    }
    if (arg === '--header-row') {
      headerRow = Number(argv[++i]);
      continue;
    }
    if (arg === '--factory' || arg === '-f') {
      factory = argv[++i];
      continue;
    }
    if (arg === '--sheet') {
      sheets.push(argv[++i]);
      continue;
    }
    if (!arg.startsWith('-')) {
      positionals.push(arg);
    }
  }

  const filePath = positionals[0] ?? '';
  if (!factory && positionals[1]) {
    factory = positionals[1];
  }

  if (!filePath) return null;
  return {
    filePath,
    outputPath,
    headerRow,
    factory,
    sheets: sheets.length ? sheets : undefined,
  };
}

async function main(): Promise<void> {
  const parsed = parseArgs(process.argv.slice(2));
  if (parsed === 'help') {
    printUsage();
    return;
  }
  if (!parsed) {
    console.error('Thiếu file Excel. Chạy với --help để xem hướng dẫn.\n');
    printUsage();
    process.exit(1);
  }

  const filePath = path.resolve(parsed.filePath);
  console.log('Đang kết nối PostgreSQL...');
  try {
    const { report, outputPath } = await runCompare({
      filePath,
      outputPath: parsed.outputPath,
      headerRow: parsed.headerRow,
      factory: parsed.factory,
      sheets: parsed.sheets,
    });

    if (parsed.factory) {
      console.log(`\nBáo cáo chỉ hiển thị xưởng: ${parsed.factory.trim().toUpperCase()}`);
    }
    printConsoleSummary(report);
    console.log(`\nĐã tạo file báo cáo: ${outputPath}`);
  } catch (err) {
    console.error(err instanceof Error ? err.message : err);
    process.exit(1);
  }
}

main().catch((err) => {
  console.error('Lỗi:', err instanceof Error ? err.message : err);
  process.exit(1);
});

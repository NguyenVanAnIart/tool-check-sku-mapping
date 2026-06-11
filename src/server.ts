import 'dotenv/config';
import express from 'express';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import multer from 'multer';
import {
  inspectExcelFile,
  runCompare,
  testDbConnection,
} from './compare-runner.js';
import {
  isDbEnvironmentId,
  listDbEnvironments,
  loadDbConfigForEnvironment,
} from './db-environments.js';
import type { CompareReport } from './types.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(__dirname, '..');
const PUBLIC_DIR = path.join(ROOT, 'public');
const UPLOAD_DIR = path.join(ROOT, 'uploads');
const REPORTS_DIR = path.join(ROOT, 'reports');

for (const dir of [UPLOAD_DIR, REPORTS_DIR, PUBLIC_DIR]) {
  fs.mkdirSync(dir, { recursive: true });
}

const upload = multer({
  dest: UPLOAD_DIR,
  limits: { fileSize: 500 * 1024 * 1024 },
});

const app = express();
const PORT = Number(process.env.UI_PORT ?? 3847);

function reportToJson(report: CompareReport) {
  return {
    sourceFile: report.sourceFile,
    comparedAt: report.comparedAt,
    sheets: report.sheets,
    totals: report.totals,
    matchRate:
      report.totals.totalExcelRows > 0
        ? ((report.totals.matched / report.totals.totalExcelRows) * 100).toFixed(1)
        : '0.0',
  };
}

function readEnvironmentId(body: Record<string, unknown>): string {
  const raw =
    typeof body.environment === 'string' ? body.environment.trim().toLowerCase() : '';
  if (!isDbEnvironmentId(raw)) {
    throw new Error('Chọn môi trường: beta hoặc production.');
  }
  return raw;
}

app.use(express.json());
app.use(express.static(PUBLIC_DIR));

app.get('/api/db-environments', (_req, res) => {
  res.json({ environments: listDbEnvironments(ROOT) });
});

app.post('/api/db-test', async (req, res) => {
  try {
    const environment = readEnvironmentId(req.body);
    const dbConfig = loadDbConfigForEnvironment(environment, ROOT);
    await testDbConnection(dbConfig);
    res.json({
      ok: true,
      environment,
      host: dbConfig.host,
      port: dbConfig.port,
      database: dbConfig.database,
      user: dbConfig.user,
    });
  } catch (err) {
    const environment =
      typeof req.body?.environment === 'string' ? req.body.environment : undefined;
    res.json({
      ok: false,
      environment,
      error: err instanceof Error ? err.message : String(err),
    });
  }
});

app.post('/api/inspect', upload.single('file'), async (req, res) => {
  if (!req.file) {
    res.status(400).json({ error: 'Chưa chọn file Excel' });
    return;
  }
  try {
    const { allSheets, factorySheets } = await inspectExcelFile(req.file.path);
    res.json({
      uploadId: req.file.filename,
      originalName: req.file.originalname,
      allSheets,
      factorySheets,
    });
  } catch (err) {
    fs.unlink(req.file.path, () => {});
    res.status(400).json({
      error: err instanceof Error ? err.message : String(err),
    });
  }
});

app.post('/api/compare', upload.single('file'), async (req, res) => {
  const uploadId = typeof req.body.uploadId === 'string' ? req.body.uploadId : '';
  const headerRow = Number(req.body.headerRow ?? 5);
  const factory =
    typeof req.body.factory === 'string' && req.body.factory.trim()
      ? req.body.factory.trim()
      : undefined;

  let sheets: string[] | undefined;
  if (typeof req.body.sheets === 'string' && req.body.sheets.trim()) {
    try {
      const parsed = JSON.parse(req.body.sheets) as string[];
      if (Array.isArray(parsed) && parsed.length > 0) sheets = parsed;
    } catch {
      sheets = req.body.sheets
        .split(',')
        .map((s: string) => s.trim())
        .filter(Boolean);
    }
  }

  let filePath: string | undefined;
  let tempFile = false;

  if (req.file) {
    filePath = req.file.path;
    tempFile = true;
  } else if (uploadId) {
    filePath = path.join(UPLOAD_DIR, uploadId);
    if (!fs.existsSync(filePath)) {
      res.status(400).json({ error: 'File upload đã hết hạn. Vui lòng chọn file lại.' });
      return;
    }
  }

  if (!filePath) {
    res.status(400).json({ error: 'Chưa chọn file Excel' });
    return;
  }

  try {
    const environment = readEnvironmentId(req.body);
    const dbConfig = loadDbConfigForEnvironment(environment, ROOT);

    const { report, outputPath } = await runCompare({
      filePath,
      headerRow,
      factory,
      sheets,
      dbConfig,
    });

    const reportId = path.basename(outputPath);
    res.json({
      ok: true,
      environment,
      summary: reportToJson(report),
      downloadUrl: `/api/download/${encodeURIComponent(reportId)}`,
      reportId,
    });
  } catch (err) {
    res.status(400).json({
      ok: false,
      error: err instanceof Error ? err.message : String(err),
    });
  } finally {
    if (tempFile && filePath) {
      fs.unlink(filePath, () => {});
    }
  }
});

app.get('/api/download/:reportId', (req, res) => {
  const reportId = path.basename(req.params.reportId);
  const filePath = path.join(REPORTS_DIR, reportId);
  if (!fs.existsSync(filePath)) {
    res.status(404).json({ error: 'Không tìm thấy báo cáo' });
    return;
  }
  res.download(filePath, reportId);
});

const server = app.listen(PORT, () => {
  const url = `http://localhost:${PORT}`;
  console.log(`\n  SKU Price Compare — Giao diện web\n  Mở trình duyệt: ${url}\n`);
  const start =
    process.platform === 'win32'
      ? `start "" "${url}"`
      : process.platform === 'darwin'
        ? `open ${url}`
        : `xdg-open ${url}`;
  import('node:child_process').then(({ exec }) => {
    exec(start, () => {});
  });
});

server.on('error', (err: NodeJS.ErrnoException) => {
  if (err.code === 'EADDRINUSE') {
    const url = `http://localhost:${PORT}`;
    console.error(`\n  Port ${PORT} đã được sử dụng.`);
    console.error(`  Giao diện có thể đã chạy — mở: ${url}`);
    console.error('  Hoặc tắt process cũ rồi chạy lại.\n');
    process.exit(1);
  }
  console.error('Lỗi server:', err.message);
  process.exit(1);
});

import path from 'node:path';
import { loadDbConfigFromEnvFile, type DbConfig } from './db.js';

export type DbEnvironmentId = 'beta' | 'production';

export interface DbEnvironmentMeta {
  label: string;
  file: string;
}

export const DB_ENVIRONMENT_META: Record<DbEnvironmentId, DbEnvironmentMeta> = {
  beta: { label: 'Beta', file: '.env.dev' },
  production: { label: 'Production', file: '.env.production' },
};

export interface DbEnvironmentInfo {
  id: DbEnvironmentId;
  label: string;
  file: string;
  host: string;
  port: number;
  database: string;
  user: string;
  available: boolean;
}

export function isDbEnvironmentId(value: string): value is DbEnvironmentId {
  return value === 'beta' || value === 'production';
}

export function loadDbConfigForEnvironment(
  environmentId: string,
  rootDir: string,
): DbConfig {
  if (!isDbEnvironmentId(environmentId)) {
    throw new Error(`Môi trường không hợp lệ: ${environmentId}. Chọn beta hoặc production.`);
  }

  const meta = DB_ENVIRONMENT_META[environmentId];
  const filePath = path.join(rootDir, meta.file);
  const config = loadDbConfigFromEnvFile(filePath);
  if (!config) {
    throw new Error(`Không tìm thấy file cấu hình ${meta.file}`);
  }
  return config;
}

export function listDbEnvironments(rootDir: string): DbEnvironmentInfo[] {
  return (Object.entries(DB_ENVIRONMENT_META) as [DbEnvironmentId, DbEnvironmentMeta][]).map(
    ([id, meta]) => {
      const config = loadDbConfigFromEnvFile(path.join(rootDir, meta.file));
      return {
        id,
        label: meta.label,
        file: meta.file,
        host: config?.host ?? '',
        port: config?.port ?? 5432,
        database: config?.database ?? '',
        user: config?.user ?? '',
        available: config !== null,
      };
    },
  );
}

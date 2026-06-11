import fs from 'node:fs';
import pg from 'pg';
import type { DbRow } from './types.js';

export interface DbConfig {
  host: string;
  port: number;
  user: string;
  password: string;
  database: string;
}

function configFromVars(vars: Record<string, string>): DbConfig {
  return {
    host: vars.DB_HOST ?? 'localhost',
    port: Number(vars.DB_PORT ?? 5432),
    user: vars.DB_USERNAME ?? vars.DB_USER ?? 'postgres',
    password: vars.DB_PASSWORD ?? '',
    database: vars.DB_NAME ?? 'pod_1',
  };
}

export function loadDbConfigFromEnv(): DbConfig {
  return configFromVars({
    DB_HOST: process.env.DB_HOST ?? '',
    DB_PORT: process.env.DB_PORT ?? '',
    DB_USERNAME: process.env.DB_USERNAME ?? process.env.DB_USER ?? '',
    DB_PASSWORD: process.env.DB_PASSWORD ?? '',
    DB_NAME: process.env.DB_NAME ?? '',
  });
}

export function loadDbConfigFromEnvFile(filePath: string): DbConfig | null {
  if (!fs.existsSync(filePath)) return null;
  const vars: Record<string, string> = {};
  const content = fs.readFileSync(filePath, 'utf8');
  for (const line of content.split(/\r?\n/)) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith('#')) continue;
    const eq = trimmed.indexOf('=');
    if (eq <= 0) continue;
    vars[trimmed.slice(0, eq).trim()] = trimmed.slice(eq + 1).trim();
  }
  return configFromVars(vars);
}

const SELECT_SQL = `
  SELECT
    id,
    sku_ecb,
    sku_factory,
    factory,
    type,
    print_tech,
    include,
    exclude,
    tag,
    stock,
    price,
    flash_sale1,
    flash_sale2,
    additional_price,
    product_name,
    color,
    size,
    ecb_product_name,
    ecb_color,
    ecb_size,
    user_email,
    status
  FROM sku_price
  WHERE factory = ANY($1::text[])
  ORDER BY factory, sku_factory, type, print_tech, include
`;

export async function fetchSkuPriceByFactories(
  config: DbConfig,
  factories: string[],
): Promise<DbRow[]> {
  if (factories.length === 0) return [];

  const client = new pg.Client(config);
  await client.connect();
  try {
    const result = await client.query<DbRow>(SELECT_SQL, [factories]);
    return result.rows;
  } finally {
    await client.end();
  }
}

export async function testConnection(config: DbConfig): Promise<void> {
  const client = new pg.Client(config);
  await client.connect();
  try {
    await client.query('SELECT 1');
  } finally {
    await client.end();
  }
}

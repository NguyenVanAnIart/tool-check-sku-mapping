export interface SkuMappingIndex {
  /** factory (lowercase) → ID xưởng → danh sách DTG Print SKU (có thể nhiều dòng mapping) */
  byFactoryId: Map<string, Map<string, string[]>>;
  /** Các cột xưởng có trong header SKU MAPPING */
  factories: Set<string>;
}

export interface ExcelRow {
  rowNumber: number;
  id: string;
  sku: string;
  productName: string | null;
  color: string | null;
  size: string | null;
  type: string;
  printTech: string;
  include: string;
  exclude: string;
  tag: string | null;
  stock: boolean;
  price: number;
  flashSale1: number | null;
  flashSale2: number | null;
  additionalPrice: number | null;
}

export interface DbRow {
  id: string;
  sku_ecb: string;
  sku_factory: string;
  factory: string;
  type: string;
  print_tech: string;
  include: string | null;
  exclude: string | null;
  tag: string | null;
  stock: boolean;
  price: string | number;
  flash_sale1: string | number | null;
  flash_sale2: string | number | null;
  additional_price: string | number | null;
  product_name: string | null;
  color: string | null;
  size: string | null;
  ecb_product_name: string | null;
  ecb_color: string | null;
  ecb_size: string | null;
  user_email: string | null;
  status: string | null;
}

export interface FieldDiff {
  field: string;
  excelValue: string;
  dbValue: string;
}

export interface CompareResult {
  factory: string;
  excelRow: ExcelRow;
  dbRow: DbRow | null;
  status: 'match' | 'mismatch' | 'not_in_db' | 'skipped';
  diffs: FieldDiff[];
  matchKey: string;
}

export interface SheetSummary {
  factory: string;
  totalExcelRows: number;
  skippedRows: number;
  matched: number;
  mismatched: number;
  notInDb: number;
  onlyInDb: number;
}

export interface CompareReport {
  sourceFile: string;
  comparedAt: string;
  sheets: SheetSummary[];
  results: CompareResult[];
  onlyInDb: Array<{ factory: string; dbRow: DbRow; matchKey: string }>;
  totals: {
    totalExcelRows: number;
    skippedRows: number;
    matched: number;
    mismatched: number;
    notInDb: number;
    onlyInDb: number;
  };
}

export const COMPARE_FIELDS = [
  { excel: 'sku', db: 'sku_factory', label: 'SKU' },
  { excel: 'productName', db: 'product_name', label: 'Product Name' },
  { excel: 'color', db: 'color', label: 'Color' },
  { excel: 'size', db: 'size', label: 'Size' },
  { excel: 'type', db: 'type', label: 'Type' },
  { excel: 'printTech', db: 'print_tech', label: 'PrintTech' },
  { excel: 'include', db: 'include', label: 'Include' },
  { excel: 'exclude', db: 'exclude', label: 'Exclude' },
  { excel: 'tag', db: 'tag', label: 'Tag' },
  { excel: 'stock', db: 'stock', label: 'Stock' },
  { excel: 'price', db: 'price', label: 'Price' },
  { excel: 'flashSale1', db: 'flash_sale1', label: 'Flash sale 1' },
  { excel: 'flashSale2', db: 'flash_sale2', label: 'Flash sale 2' },
  { excel: 'additionalPrice', db: 'additional_price', label: 'Additional Price' },
] as const;

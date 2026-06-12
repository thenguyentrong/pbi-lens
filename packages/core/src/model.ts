import { PowerBiClient } from "./rest";

export interface ModelTable {
  name: string;
  description?: string;
  storageMode?: string;
  dataCategory?: string;
  isHidden?: boolean;
}

export interface ModelColumn {
  table: string;
  name: string;
  dataType?: string;
  isHidden?: boolean;
  expression?: string;
  formatString?: string;
  sortByColumn?: string;
  summarizeBy?: string;
  dataCategory?: string;
}

export interface ModelMeasure {
  table: string;
  name: string;
  dataType?: string;
  expression?: string;
  formatString?: string;
  displayFolder?: string;
  description?: string;
  isHidden?: boolean;
  state?: string;
}

export interface ModelRelationship {
  fromTable: string;
  fromColumn: string;
  fromCardinality?: string;
  toTable: string;
  toColumn: string;
  toCardinality?: string;
  isActive?: boolean;
  crossFilteringBehavior?: string;
  securityFilteringBehavior?: string;
  state?: string;
}

export interface ModelInfo {
  tables: ModelTable[];
  columns: ModelColumn[];
  measures: ModelMeasure[];
  relationships: ModelRelationship[];
  warnings: string[];
}

// INFO.VIEW.* are the user-friendly DMV wrappers (names instead of ids) and
// the only schema surface executeQueries reliably permits — plain INFO.*
// functions are blocked on locked-down tenants.
const Q_TABLES =
  'EVALUATE SELECTCOLUMNS(INFO.VIEW.TABLES(), "name", [Name], "description", [Description], "storageMode", [StorageMode], "dataCategory", [DataCategory], "isHidden", [IsHidden], "isPrivate", [IsPrivate])';
const Q_COLUMNS =
  'EVALUATE SELECTCOLUMNS(INFO.VIEW.COLUMNS(), "table", [Table], "name", [Name], "dataType", [DataType], "colType", [Type], "isHidden", [IsHidden], "expression", [Expression], "formatString", [FormatString], "sortByColumn", [SortByColumn], "summarizeBy", [SummarizeBy], "dataCategory", [DataCategory])';
const Q_MEASURES =
  'EVALUATE SELECTCOLUMNS(INFO.VIEW.MEASURES(), "table", [Table], "name", [Name], "dataType", [DataType], "expression", [Expression], "formatString", [FormatString], "displayFolder", [DisplayFolder], "description", [Description], "isHidden", [IsHidden], "state", [State])';
const Q_RELATIONSHIPS =
  'EVALUATE SELECTCOLUMNS(INFO.VIEW.RELATIONSHIPS(), "fromTable", [FromTable], "fromColumn", [FromColumn], "fromCardinality", [FromCardinality], "toTable", [ToTable], "toColumn", [ToColumn], "toCardinality", [ToCardinality], "isActive", [IsActive], "crossFilteringBehavior", [CrossFilteringBehavior], "securityFilteringBehavior", [SecurityFilteringBehavior], "state", [State])';

/** executeQueries returns keys bracketed ("[name]") — strip them. */
function cleanRows(rows: Record<string, unknown>[]): Record<string, unknown>[] {
  return rows.map((row) => {
    const out: Record<string, unknown> = {};
    for (const [key, value] of Object.entries(row)) {
      out[key.replace(/^\[|\]$/g, "")] = value;
    }
    return out;
  });
}

/** Tables Power BI generates for auto date/time — pure noise for an agent. */
function isAutoDateTable(name: string): boolean {
  return name.startsWith("LocalDateTable_") || name.startsWith("DateTableTemplate_");
}

/**
 * Read the dataset's schema (tables, columns, measures, relationships) via
 * DAX INFO.VIEW.* through executeQueries — Pro-compatible, no XMLA needed.
 * Each section degrades independently into `warnings`; never hard-fails.
 */
export async function getModelInfo(
  client: PowerBiClient,
  datasetId: string,
  opts: { includeHidden?: boolean } = {}
): Promise<ModelInfo> {
  const warnings: string[] = [];

  async function section(label: string, query: string): Promise<Record<string, unknown>[]> {
    try {
      const result = await client.executeDax(datasetId, query);
      return cleanRows(result.rows);
    } catch (e) {
      warnings.push(`${label} unavailable: ${(e as Error).message.split("\n")[0]}`);
      return [];
    }
  }

  const [rawTables, rawColumns, rawMeasures, rawRelationships] = await Promise.all([
    section("INFO.VIEW.TABLES", Q_TABLES),
    section("INFO.VIEW.COLUMNS", Q_COLUMNS),
    section("INFO.VIEW.MEASURES", Q_MEASURES),
    section("INFO.VIEW.RELATIONSHIPS", Q_RELATIONSHIPS),
  ]);

  const noiseTables = new Set(
    rawTables
      .filter((t) => t.isPrivate === true || isAutoDateTable(String(t.name ?? "")))
      .map((t) => String(t.name))
  );
  // Relationships reference auto-date tables even when INFO.VIEW.TABLES fails.
  const isNoise = (table: unknown) => noiseTables.has(String(table)) || isAutoDateTable(String(table ?? ""));
  const keepHidden = opts.includeHidden === true;

  const tables: ModelTable[] = rawTables
    .filter((t) => !isNoise(t.name) && (keepHidden || t.isHidden !== true))
    .map((t) => ({
      name: String(t.name),
      description: (t.description as string) || undefined,
      storageMode: (t.storageMode as string) || undefined,
      dataCategory: (t.dataCategory as string) || undefined,
      isHidden: t.isHidden === true || undefined,
    }));

  const columns: ModelColumn[] = rawColumns
    .filter(
      (c) => !isNoise(c.table) && c.colType !== "RowNumber" && (keepHidden || c.isHidden !== true)
    )
    .map((c) => ({
      table: String(c.table),
      name: String(c.name),
      dataType: (c.dataType as string) || undefined,
      isHidden: c.isHidden === true || undefined,
      expression: (c.expression as string) || undefined,
      formatString: (c.formatString as string) || undefined,
      sortByColumn: (c.sortByColumn as string) || undefined,
      summarizeBy: (c.summarizeBy as string) || undefined,
      dataCategory: (c.dataCategory as string) || undefined,
    }));

  const measures: ModelMeasure[] = rawMeasures
    .filter((m) => !isNoise(m.table) && (keepHidden || m.isHidden !== true))
    .map((m) => ({
      table: String(m.table),
      name: String(m.name),
      dataType: (m.dataType as string) || undefined,
      expression: (m.expression as string) || undefined,
      formatString: (m.formatString as string) || undefined,
      displayFolder: (m.displayFolder as string) || undefined,
      description: (m.description as string) || undefined,
      isHidden: m.isHidden === true || undefined,
      state: (m.state as string) || undefined,
    }));

  const relationships: ModelRelationship[] = rawRelationships
    .filter((r) => !isNoise(r.fromTable) && !isNoise(r.toTable))
    .map((r) => ({
      fromTable: String(r.fromTable),
      fromColumn: String(r.fromColumn),
      fromCardinality: (r.fromCardinality as string) || undefined,
      toTable: String(r.toTable),
      toColumn: String(r.toColumn),
      toCardinality: (r.toCardinality as string) || undefined,
      isActive: r.isActive === true,
      crossFilteringBehavior: (r.crossFilteringBehavior as string) || undefined,
      securityFilteringBehavior: (r.securityFilteringBehavior as string) || undefined,
      state: (r.state as string) || undefined,
    }));

  if (measures.length > 0 && measures.every((m) => !m.expression)) {
    warnings.push(
      "This tenant does not expose measure DAX via executeQueries (INFO.VIEW.MEASURES returned blank " +
        "Expression and INFO.MEASURES is blocked). Read the local .pbip source " +
        "(definition\\**\\*.tmdl or model.bim) for measure definitions."
    );
  }

  return { tables, columns, measures, relationships, warnings };
}

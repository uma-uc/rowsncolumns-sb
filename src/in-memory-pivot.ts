import type { CellFormat, SheetRange } from '@rowsncolumns/spreadsheet'
import type { Store } from 'tinybase'

export type PivotPrimitive = string | number | boolean
export type PivotSourceRow = Record<string, PivotPrimitive>
export type PivotField<T extends PivotSourceRow> = Extract<keyof T, string>
export type AggregationFunction = 'sum' | 'count' | 'avg' | 'min' | 'max'

export type PivotValueConfig<T extends PivotSourceRow> = {
  field: PivotField<T>
  label?: string
  aggFunc: AggregationFunction
}

export type InMemoryPivotConfig<T extends PivotSourceRow> = {
  sourceRange: SheetRange
  targetPosition: {
    sheetId: number
    rowIndex: number
    columnIndex: number
  }
  rows: PivotField<T>[]
  columns: PivotField<T>[]
  values: PivotValueConfig<T>[]
  filters?: Partial<Record<PivotField<T>, readonly PivotPrimitive[]>>
}

export type PivotCellKind = 'header' | 'row-header' | 'value' | 'total-header' | 'total-value'

export type PivotCellMetadata<T extends PivotSourceRow> = {
  kind: PivotCellKind
  rowPath: PivotPrimitive[]
  columnPath: PivotPrimitive[]
  valueField?: PivotField<T>
  sourceRowCount: number
}

export type PivotGroup = {
  key: string
  path: PivotPrimitive[]
  label: string
}

export type InMemoryPivotResult<T extends PivotSourceRow> = {
  values: (string | number | boolean | null)[][]
  metadata: PivotCellMetadata<T>[][]
  rowGroups: PivotGroup[]
  columnGroups: PivotGroup[]
  sourceRows: T[]
}

export type PivotBatch<T extends PivotSourceRow> = {
  range: SheetRange
  values: InMemoryPivotResult<T>['values']
  formats: (CellFormat | undefined)[][]
}

const GROUP_SEPARATOR = '\u001f'

export function createInMemoryPivot<T extends PivotSourceRow>(
  sourceRows: readonly T[],
  config: InMemoryPivotConfig<T>,
): InMemoryPivotResult<T> {
  const filteredRows = applyFilters(sourceRows, config.filters)
  const rowGroups = getGroups(filteredRows, config.rows)
  const columnGroups = getGroups(filteredRows, config.columns)
  const effectiveRowGroups = rowGroups.length ? rowGroups : [{ key: 'grand-total', path: [], label: 'Grand Total' }]
  const effectiveColumnGroups = columnGroups.length ? columnGroups : [{ key: 'all-columns', path: [], label: '' }]
  const values = config.values.length ? config.values : []
  const header = [
    ...config.rows.map(formatFieldName),
    ...effectiveColumnGroups.flatMap(group => values.map(value => formatColumnHeader(group, value))),
    ...(config.columns.length ? values.map(value => `Total ${getValueLabel(value)}`) : []),
  ]
  const headerMetadata = header.map<PivotCellMetadata<T>>((_, index) => ({
    kind: index < config.rows.length ? 'row-header' : 'header',
    rowPath: [],
    columnPath: [],
    sourceRowCount: filteredRows.length,
  }))

  const body = effectiveRowGroups.map(group => {
    const groupRows = matchRowsByPath(filteredRows, config.rows, group.path)
    const rowValues = [
      ...group.path,
      ...effectiveColumnGroups.flatMap(columnGroup =>
        values.map(value => aggregateRows(matchRowsByPath(groupRows, config.columns, columnGroup.path), value)),
      ),
      ...(config.columns.length ? values.map(value => aggregateRows(groupRows, value)) : []),
    ]
    const rowMetadata = rowValues.map<PivotCellMetadata<T>>((_, index) => {
      if (index < config.rows.length) {
        return {
          kind: 'row-header',
          rowPath: group.path,
          columnPath: [],
          sourceRowCount: groupRows.length,
        }
      }
      const valueOffset = index - config.rows.length
      const totalOffset = effectiveColumnGroups.length * values.length
      if (config.columns.length && valueOffset >= totalOffset) {
        const value = values[valueOffset - totalOffset]
        return {
          kind: 'total-value',
          rowPath: group.path,
          columnPath: [],
          valueField: value?.field,
          sourceRowCount: groupRows.length,
        }
      }
      const columnIndex = Math.floor(valueOffset / values.length)
      const valueIndex = valueOffset % values.length
      const columnGroup = effectiveColumnGroups[columnIndex]
      const value = values[valueIndex]
      return {
        kind: 'value',
        rowPath: group.path,
        columnPath: columnGroup?.path ?? [],
        valueField: value?.field,
        sourceRowCount: matchRowsByPath(groupRows, config.columns, columnGroup?.path ?? []).length,
      }
    })

    return { rowValues, rowMetadata }
  })

  const totalRow =
    config.rows.length > 0 ? createGrandTotalRow(filteredRows, config, effectiveColumnGroups, values) : null

  return {
    values: [header, ...body.map(row => row.rowValues), ...(totalRow ? [totalRow.rowValues] : [])],
    metadata: [headerMetadata, ...body.map(row => row.rowMetadata), ...(totalRow ? [totalRow.rowMetadata] : [])],
    rowGroups: effectiveRowGroups,
    columnGroups: effectiveColumnGroups,
    sourceRows: filteredRows,
  }
}

export function getPivotRowsFromTinyBaseTable<T extends PivotSourceRow>(
  store: Store,
  tableId: string,
  fields: readonly PivotField<T>[],
): T[] {
  return Object.values(store.getTable(tableId)).map(row => {
    const pivotRow: Partial<T> = {}
    for (const field of fields) {
      const cell = row[field]
      pivotRow[field] = (isPivotPrimitive(cell) ? cell : '') as T[typeof field]
    }
    return pivotRow as T
  })
}

export function createInMemoryPivotBatch<T extends PivotSourceRow>(
  config: InMemoryPivotConfig<T>,
  result: InMemoryPivotResult<T>,
): PivotBatch<T> {
  return {
    range: getTargetRange(config.targetPosition, result.values),
    values: result.values,
    formats: createInMemoryPivotFormats(result),
  }
}

export function createInMemoryPivotFormats<T extends PivotSourceRow>(
  result: InMemoryPivotResult<T>,
  options: {
    headerBackgroundColor?: CellFormat['backgroundColor']
    headerTextColor?: NonNullable<CellFormat['textFormat']>['color']
  } = {},
): (CellFormat | undefined)[][] {
  const headerBackgroundColor = options.headerBackgroundColor ?? '#dbeafe'
  const headerTextColor = options.headerTextColor ?? '#172554'
  const headerFormat: CellFormat = {
    backgroundColor: headerBackgroundColor,
    textFormat: { bold: true, color: headerTextColor },
    horizontalAlignment: 'center',
  }
  const rowHeaderFormat: CellFormat = {
    backgroundColor: '#f8fafc',
    textFormat: { bold: true, color: '#0f172a' },
  }
  const totalFormat: CellFormat = {
    backgroundColor: '#ecfdf5',
    textFormat: { bold: true, color: '#064e3b' },
    horizontalAlignment: 'right',
  }
  const valueFormat: CellFormat = {
    numberFormat: { type: 'CURRENCY', pattern: '$#,##0' },
    horizontalAlignment: 'right',
  }

  return result.metadata.map(row =>
    row.map(cell => {
      if (cell.kind === 'header') return headerFormat
      if (cell.kind === 'row-header') return rowHeaderFormat
      if (cell.kind === 'total-header' || cell.kind === 'total-value') return totalFormat
      if (cell.kind === 'value') return valueFormat
      return undefined
    }),
  )
}

export function expandInMemoryPivotGroup<T extends PivotSourceRow>(
  sourceRows: readonly T[],
  config: InMemoryPivotConfig<T>,
  rowPath: readonly PivotPrimitive[],
  columnPath: readonly PivotPrimitive[] = [],
): T[] {
  const filteredRows = applyFilters(sourceRows, config.filters)
  const rowMatches = matchRowsByPath(filteredRows, config.rows, [...rowPath])
  return matchRowsByPath(rowMatches, config.columns, [...columnPath])
}

function createGrandTotalRow<T extends PivotSourceRow>(
  rows: T[],
  config: InMemoryPivotConfig<T>,
  columnGroups: PivotGroup[],
  values: PivotValueConfig<T>[],
): { rowValues: (string | number | boolean | null)[]; rowMetadata: PivotCellMetadata<T>[] } {
  const rowHeader = config.rows.map((_, index) => (index === 0 ? 'Grand Total' : ''))
  const rowValues = [
    ...rowHeader,
    ...columnGroups.flatMap(columnGroup =>
      values.map(value => aggregateRows(matchRowsByPath(rows, config.columns, columnGroup.path), value)),
    ),
    ...(config.columns.length ? values.map(value => aggregateRows(rows, value)) : []),
  ]
  const rowMetadata = rowValues.map<PivotCellMetadata<T>>((_, index) => {
    if (index < config.rows.length) {
      return { kind: 'total-header', rowPath: [], columnPath: [], sourceRowCount: rows.length }
    }
    const valueOffset = index - config.rows.length
    const totalOffset = columnGroups.length * values.length
    if (config.columns.length && valueOffset >= totalOffset) {
      const value = values[valueOffset - totalOffset]
      return {
        kind: 'total-value',
        rowPath: [],
        columnPath: [],
        valueField: value?.field,
        sourceRowCount: rows.length,
      }
    }
    const columnIndex = Math.floor(valueOffset / values.length)
    const valueIndex = valueOffset % values.length
    const columnGroup = columnGroups[columnIndex]
    const value = values[valueIndex]
    return {
      kind: 'total-value',
      rowPath: [],
      columnPath: columnGroup?.path ?? [],
      valueField: value?.field,
      sourceRowCount: matchRowsByPath(rows, config.columns, columnGroup?.path ?? []).length,
    }
  })

  return { rowValues, rowMetadata }
}

function applyFilters<T extends PivotSourceRow>(rows: readonly T[], filters: InMemoryPivotConfig<T>['filters']): T[] {
  if (!filters) return [...rows]
  const entries = Object.entries(filters) as [PivotField<T>, readonly PivotPrimitive[]][]
  if (!entries.length) return [...rows]
  return rows.filter(row =>
    entries.every(([field, allowedValues]) => !allowedValues.length || allowedValues.includes(row[field])),
  )
}

function isPivotPrimitive(value: unknown): value is PivotPrimitive {
  return typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean'
}

function getGroups<T extends PivotSourceRow>(rows: readonly T[], fields: PivotField<T>[]): PivotGroup[] {
  if (!fields.length) return []
  const groups = new Map<string, PivotGroup>()
  for (const row of rows) {
    const path = fields.map(field => row[field])
    const key = getGroupKey(path)
    if (!groups.has(key)) {
      groups.set(key, { key, path, label: path.map(String).join(' / ') })
    }
  }
  return [...groups.values()].sort((a, b) => a.label.localeCompare(b.label))
}

function matchRowsByPath<T extends PivotSourceRow>(
  rows: readonly T[],
  fields: PivotField<T>[],
  path: PivotPrimitive[],
): T[] {
  if (!fields.length || !path.length) return [...rows]
  return rows.filter(row => fields.every((field, index) => row[field] === path[index]))
}

function aggregateRows<T extends PivotSourceRow>(rows: readonly T[], value: PivotValueConfig<T> | undefined): number {
  if (!value) return 0
  const numericValues: number[] = []
  for (const row of rows) {
    const cell = row[value.field] as PivotPrimitive
    if (typeof cell === 'number' && Number.isFinite(cell)) {
      numericValues.push(cell)
    }
  }

  if (value.aggFunc === 'count') return rows.filter(row => row[value.field] !== '').length
  if (!numericValues.length) return 0
  if (value.aggFunc === 'avg') return numericValues.reduce((sum, cell) => sum + cell, 0) / numericValues.length
  if (value.aggFunc === 'min') return Math.min(...numericValues)
  if (value.aggFunc === 'max') return Math.max(...numericValues)
  return numericValues.reduce((sum, cell) => sum + cell, 0)
}

function getTargetRange(
  targetPosition: InMemoryPivotConfig<PivotSourceRow>['targetPosition'],
  values: InMemoryPivotResult<PivotSourceRow>['values'],
): SheetRange {
  return {
    sheetId: targetPosition.sheetId,
    startRowIndex: targetPosition.rowIndex,
    endRowIndex: targetPosition.rowIndex + Math.max(values.length - 1, 0),
    startColumnIndex: targetPosition.columnIndex,
    endColumnIndex: targetPosition.columnIndex + Math.max((values[0]?.length ?? 1) - 1, 0),
  }
}

function getGroupKey(path: readonly PivotPrimitive[]): string {
  return path.map(String).join(GROUP_SEPARATOR)
}

function formatColumnHeader<T extends PivotSourceRow>(group: PivotGroup, value: PivotValueConfig<T>): string {
  const valueLabel = getValueLabel(value)
  return group.label ? `${group.label} ${valueLabel}` : valueLabel
}

function getValueLabel<T extends PivotSourceRow>(value: PivotValueConfig<T>): string {
  return value.label ?? `${value.aggFunc.toUpperCase()} ${formatFieldName(value.field)}`
}

function formatFieldName(field: string): string {
  return field
    .replace(/([a-z0-9])([A-Z])/g, '$1 $2')
    .replace(/[-_]/g, ' ')
    .replace(/\b\w/g, char => char.toUpperCase())
}

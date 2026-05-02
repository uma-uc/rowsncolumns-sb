import {
  CanvasGrid,
  SpreadsheetProvider,
  type CanvasGridProps,
  type CellData,
  type CellFormat,
  type ExtendedValue,
} from '@rowsncolumns/spreadsheet'
import '@rowsncolumns/spreadsheet/dist/spreadsheet.min.css'
import type { Meta, StoryObj } from '@storybook/react-vite'
import * as React from 'react'
import { createStore, type Store } from 'tinybase'
import {
  createInMemoryPivot,
  createInMemoryPivotBatch,
  expandInMemoryPivotGroup,
  getPivotRowsFromTinyBaseTable,
  type AggregationFunction,
  type InMemoryPivotConfig,
  type PivotCellMetadata,
  type PivotField,
  type PivotPrimitive,
  type PivotSourceRow,
} from './in-memory-pivot'

const salesTableId = 'sales'
const sheetId = 1

type SalesRow = PivotSourceRow & {
  orderId: string
  region: string
  product: string
  channel: string
  quarter: string
  sales: number
  quantity: number
}

type SalesField = PivotField<SalesRow>
type PreviewCellValue = string | number | boolean | null | undefined

const salesFields = ['orderId', 'region', 'product', 'channel', 'quarter', 'sales', 'quantity'] as const
const groupFields = ['region', 'product', 'channel', 'quarter'] as const

const initialSalesRows: SalesRow[] = [
  {
    orderId: 'north-laptop-direct-q1',
    region: 'North',
    product: 'Laptop',
    channel: 'Direct',
    quarter: 'Q1',
    sales: 132000,
    quantity: 44,
  },
  {
    orderId: 'north-laptop-partner-q1',
    region: 'North',
    product: 'Laptop',
    channel: 'Partner',
    quarter: 'Q1',
    sales: 91000,
    quantity: 31,
  },
  {
    orderId: 'north-monitor-direct-q1',
    region: 'North',
    product: 'Monitor',
    channel: 'Direct',
    quarter: 'Q1',
    sales: 62000,
    quantity: 55,
  },
  {
    orderId: 'north-monitor-partner-q2',
    region: 'North',
    product: 'Monitor',
    channel: 'Partner',
    quarter: 'Q2',
    sales: 74000,
    quantity: 63,
  },
  {
    orderId: 'south-laptop-direct-q2',
    region: 'South',
    product: 'Laptop',
    channel: 'Direct',
    quarter: 'Q2',
    sales: 118000,
    quantity: 39,
  },
  {
    orderId: 'south-laptop-partner-q2',
    region: 'South',
    product: 'Laptop',
    channel: 'Partner',
    quarter: 'Q2',
    sales: 86000,
    quantity: 28,
  },
  {
    orderId: 'south-monitor-direct-q3',
    region: 'South',
    product: 'Monitor',
    channel: 'Direct',
    quarter: 'Q3',
    sales: 68000,
    quantity: 60,
  },
  {
    orderId: 'east-laptop-direct-q3',
    region: 'East',
    product: 'Laptop',
    channel: 'Direct',
    quarter: 'Q3',
    sales: 143000,
    quantity: 47,
  },
  {
    orderId: 'east-monitor-partner-q3',
    region: 'East',
    product: 'Monitor',
    channel: 'Partner',
    quarter: 'Q3',
    sales: 71000,
    quantity: 59,
  },
  {
    orderId: 'west-laptop-partner-q4',
    region: 'West',
    product: 'Laptop',
    channel: 'Partner',
    quarter: 'Q4',
    sales: 96000,
    quantity: 32,
  },
  {
    orderId: 'west-monitor-direct-q4',
    region: 'West',
    product: 'Monitor',
    channel: 'Direct',
    quarter: 'Q4',
    sales: 83000,
    quantity: 71,
  },
  {
    orderId: 'west-tablet-direct-q4',
    region: 'West',
    product: 'Tablet',
    channel: 'Direct',
    quarter: 'Q4',
    sales: 54000,
    quantity: 48,
  },
]

const meta: Meta = {
  title: 'RowsnColumns/In Memory Pivot',
  parameters: {
    layout: 'fullscreen',
  },
}

export default meta

type Story = StoryObj

export const CompleteExample: Story = {
  render: () => <TinyBasePivotStory />,
}

export const FormattingPivotResults: Story = {
  render: () => <TinyBasePivotStory initialMode="formatting" />,
}

export const DrillDownSupport: Story = {
  render: () => <TinyBasePivotStory initialMode="drilldown" />,
}

export const TinyBaseInitialization: Story = {
  render: () => <TinyBasePivotStory initialMode="initialization" />,
}

function TinyBasePivotStory({
  initialMode = 'complete',
}: {
  initialMode?: 'complete' | 'formatting' | 'drilldown' | 'initialization'
}) {
  const store = React.useMemo(createSalesStore, [])
  const { rows: sourceRows, refresh: refreshSourceRows } = useTinyBaseSalesRows(store)
  const [rowField, setRowField] = React.useState<SalesField>('region')
  const [columnField, setColumnField] = React.useState<SalesField>('product')
  const [aggregation, setAggregation] = React.useState<AggregationFunction>('sum')
  const [regionFilter, setRegionFilter] = React.useState('all')
  const [selectedCell, setSelectedCell] = React.useState<PivotCellMetadata<SalesRow> | null>(() =>
    initialMode === 'drilldown'
      ? { kind: 'value', rowPath: ['North'], columnPath: ['Laptop'], valueField: 'sales', sourceRowCount: 2 }
      : null,
  )

  const config = React.useMemo<InMemoryPivotConfig<SalesRow>>(
    () => ({
      sourceRange: {
        sheetId,
        startRowIndex: 0,
        endRowIndex: sourceRows.length,
        startColumnIndex: 0,
        endColumnIndex: salesFields.length - 1,
      },
      targetPosition: { sheetId, rowIndex: 1, columnIndex: 9 },
      rows: [rowField],
      columns: [columnField],
      values: [{ field: 'sales', label: aggregation === 'count' ? 'Orders' : 'Revenue', aggFunc: aggregation }],
      filters: regionFilter === 'all' ? undefined : { region: [regionFilter] },
    }),
    [aggregation, columnField, regionFilter, rowField, sourceRows.length],
  )
  const result = React.useMemo(() => createInMemoryPivot(sourceRows, config), [config, sourceRows])
  const batch = React.useMemo(() => createInMemoryPivotBatch(config, result), [config, result])
  const drillRows = selectedCell
    ? expandInMemoryPivotGroup(sourceRows, config, selectedCell.rowPath, selectedCell.columnPath)
    : []
  const handleAddTinyBaseRow = React.useCallback(() => {
    addTinyBaseSaleRow(store)
    refreshSourceRows()
  }, [refreshSourceRows, store])

  return (
    <div className="page">
      <header className="topbar">
        <div>
          <h1 className="title">Rows n Columns Pivot</h1>
          <div className="badge-row">
            <Badge tone="info">TinyBase source</Badge>
            <Badge tone="success">In-memory transform</Badge>
            <Badge tone="neutral">DuckDB unused</Badge>
          </div>
        </div>
        <button type="button" className="button" onClick={handleAddTinyBaseRow}>
          Add TinyBase row
        </button>
      </header>

      <main className="app-grid">
        <aside className="sidebar">
          <PivotControls
            aggregation={aggregation}
            columnField={columnField}
            regionFilter={regionFilter}
            rowField={rowField}
            onChangeAggregation={setAggregation}
            onChangeColumnField={setColumnField}
            onChangeRegionFilter={setRegionFilter}
            onChangeRowField={setRowField}
          />
          <InitializationPanel mode={initialMode} sourceRowCount={sourceRows.length} />
        </aside>

        <section className="content">
          <div className="story-grid">
            <SourceRowsTable rows={sourceRows} />
            <PivotResultTable
              aggregation={aggregation}
              result={result}
              selectedCell={selectedCell}
              onSelectCell={setSelectedCell}
            />
            <BatchPreview batch={batch} visible={initialMode === 'formatting'} />
            <DrillDownTable
              rows={drillRows}
              selectedCell={selectedCell}
              visible={initialMode === 'drilldown' || Boolean(selectedCell)}
            />
            <RowsNColumnsCanvasPreview
              batch={batch}
              sourceRows={sourceRows}
              store={store}
              onAddTinyBaseRow={handleAddTinyBaseRow}
            />
          </div>
        </section>
      </main>
    </div>
  )
}

function Badge({ children, tone }: { children: React.ReactNode; tone: 'info' | 'success' | 'neutral' }) {
  return <span className={`badge ${tone}`}>{children}</span>
}

function PivotControls({
  aggregation,
  columnField,
  regionFilter,
  rowField,
  onChangeAggregation,
  onChangeColumnField,
  onChangeRegionFilter,
  onChangeRowField,
}: {
  aggregation: AggregationFunction
  columnField: SalesField
  regionFilter: string
  rowField: SalesField
  onChangeAggregation: (value: AggregationFunction) => void
  onChangeColumnField: (value: SalesField) => void
  onChangeRegionFilter: (value: string) => void
  onChangeRowField: (value: SalesField) => void
}) {
  return (
    <div className="control-stack">
      <ControlSelect
        label="Rows"
        value={rowField}
        options={groupFields}
        onChange={value => onChangeRowField(value as SalesField)}
      />
      <ControlSelect
        label="Columns"
        value={columnField}
        options={groupFields}
        onChange={value => onChangeColumnField(value as SalesField)}
      />
      <ControlSelect
        label="Value"
        value={aggregation}
        options={['sum', 'avg', 'count', 'min', 'max']}
        onChange={value => onChangeAggregation(value as AggregationFunction)}
      />
      <ControlSelect
        label="Region"
        value={regionFilter}
        options={['all', 'North', 'South', 'East', 'West']}
        onChange={onChangeRegionFilter}
      />
    </div>
  )
}

function ControlSelect({
  label,
  value,
  options,
  onChange,
}: {
  label: string
  value: string
  options: readonly string[]
  onChange: (value: string) => void
}) {
  return (
    <label className="control">
      <span>{label}</span>
      <select value={value} onChange={event => onChange(event.target.value)}>
        {options.map(option => (
          <option key={option} value={option}>
            {formatFieldLabel(option)}
          </option>
        ))}
      </select>
    </label>
  )
}

function InitializationPanel({ mode, sourceRowCount }: { mode: string; sourceRowCount: number }) {
  if (mode !== 'initialization') return null
  return (
    <div className="runtime">
      <div className="runtime-title">Runtime</div>
      <div className="status-stack">
        <StatusLine label="TinyBase store" value="ready" tone="success" />
        <StatusLine label="Source rows" value={String(sourceRowCount)} tone="info" />
        <StatusLine label="Pivot engine" value="memory" tone="neutral" />
      </div>
    </div>
  )
}

function StatusLine({ label, value, tone }: { label: string; value: string; tone: 'success' | 'info' | 'neutral' }) {
  return (
    <div className="status-line">
      <span>{label}</span>
      <Badge tone={tone}>{value}</Badge>
    </div>
  )
}

function SourceRowsTable({ rows }: { rows: SalesRow[] }) {
  return (
    <section className="panel">
      <SectionHeader title="TinyBase Sales Table" meta={`${rows.length} rows`} />
      <div className="table-wrap">
        <table className="table">
          <thead>
            <tr>
              <th>Region</th>
              <th>Product</th>
              <th>Channel</th>
              <th>Quarter</th>
              <th className="right">Sales</th>
              <th className="right">Qty</th>
            </tr>
          </thead>
          <tbody>
            {rows.map(row => (
              <tr key={row.orderId}>
                <td>{row.region}</td>
                <td>{row.product}</td>
                <td>{row.channel}</td>
                <td>{row.quarter}</td>
                <td className="right">{formatCurrency(row.sales)}</td>
                <td className="right">{row.quantity}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </section>
  )
}

function PivotResultTable({
  aggregation,
  result,
  selectedCell,
  onSelectCell,
}: {
  aggregation: AggregationFunction
  result: ReturnType<typeof createInMemoryPivot<SalesRow>>
  selectedCell: PivotCellMetadata<SalesRow> | null
  onSelectCell: (cell: PivotCellMetadata<SalesRow>) => void
}) {
  const [header = [], ...rows] = result.values
  return (
    <section className="panel">
      <SectionHeader title="Pivot Result" meta={`${result.sourceRows.length} source rows`} />
      <div className="table-wrap">
        <table className="table">
          <thead>
            <tr>
              {header.map((cell, index) => (
                <th key={`${String(cell)}-${index}`} className={index === 0 ? undefined : 'right'}>
                  {cell}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {rows.map((row, rowIndex) => (
              <tr key={`${row.join('-')}-${rowIndex}`}>
                {row.map((cell, columnIndex) => {
                  const metadata = result.metadata[rowIndex + 1]?.[columnIndex]
                  const isInteractive = metadata?.kind === 'value' || metadata?.kind === 'total-value'
                  const isSelected = isSamePivotCell(selectedCell, metadata)
                  return (
                    <td
                      key={`${String(cell)}-${columnIndex}`}
                      className={[
                        columnIndex === 0 ? undefined : 'right',
                        isInteractive ? 'interactive' : undefined,
                        isSelected ? 'selected' : undefined,
                      ]
                        .filter(Boolean)
                        .join(' ')}
                      onClick={() => {
                        if (metadata && isInteractive) onSelectCell(metadata)
                      }}
                    >
                      {formatPivotCell(cell, aggregation, metadata)}
                    </td>
                  )
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </section>
  )
}

function BatchPreview({
  batch,
  visible,
}: {
  batch: ReturnType<typeof createInMemoryPivotBatch<SalesRow>>
  visible: boolean
}) {
  if (!visible) return null
  return (
    <section className="panel">
      <SectionHeader title="Rows n Columns Batch" meta={formatRange(batch.range)} />
      <div className="batch-box">
        <BatchLine label="values" value={`${batch.values.length} rows x ${batch.values[0]?.length ?? 0} columns`} />
        <BatchLine label="formats" value={`${batch.formats.flat().filter(Boolean).length} formatted cells`} />
        <BatchLine label="target" value={formatRange(batch.range)} />
      </div>
    </section>
  )
}

function BatchLine({ label, value }: { label: string; value: string }) {
  return (
    <div className="batch-line">
      <span className="mono">{label}</span>
      <span>{value}</span>
    </div>
  )
}

function DrillDownTable({
  rows,
  selectedCell,
  visible,
}: {
  rows: SalesRow[]
  selectedCell: PivotCellMetadata<SalesRow> | null
  visible: boolean
}) {
  if (!visible) return null
  return (
    <section className="panel">
      <SectionHeader
        title="Drill Down Rows"
        meta={selectedCell ? selectedCell.rowPath.concat(selectedCell.columnPath).join(' / ') : 'Select a value'}
      />
      <div className="table-wrap">
        <table className="table">
          <thead>
            <tr>
              <th>Order</th>
              <th>Region</th>
              <th>Product</th>
              <th>Channel</th>
              <th className="right">Sales</th>
            </tr>
          </thead>
          <tbody>
            {rows.map(row => (
              <tr key={row.orderId}>
                <td>{row.orderId}</td>
                <td>{row.region}</td>
                <td>{row.product}</td>
                <td>{row.channel}</td>
                <td className="right">{formatCurrency(row.sales)}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </section>
  )
}

function RowsNColumnsCanvasPreview({
  batch,
  onAddTinyBaseRow,
  sourceRows,
  store,
}: {
  batch: ReturnType<typeof createInMemoryPivotBatch<SalesRow>>
  onAddTinyBaseRow: () => void
  sourceRows: SalesRow[]
  store: Store
}) {
  const [editRevision, bumpEditRevision] = React.useReducer((value: number) => value + 1, 0)
  const preview = React.useMemo(() => createRowsNColumnsPreview(sourceRows, batch), [batch, sourceRows])
  const sheetRowCount = preview.values.length + 8
  const sheetColumnCount = preview.columnCount + 4
  const getPreviewValue = React.useCallback(
    (rowIndex: number, columnIndex: number) => {
      if (isSourceCell(rowIndex, columnIndex, sourceRows)) {
        return getSourceCellValue(store, sourceRows, rowIndex, columnIndex)
      }
      return preview.values[rowIndex]?.[columnIndex]
    },
    [editRevision, preview, sourceRows, store],
  )
  const getEffectiveValue = React.useCallback(
    (_sheetId: number, rowIndex: number, columnIndex: number) => getPreviewValue(rowIndex, columnIndex),
    [getPreviewValue],
  )
  const getExtendedValue = React.useCallback(
    (_sheetId: number, rowIndex: number, columnIndex: number) =>
      toExtendedValue(getPreviewValue(rowIndex, columnIndex)),
    [getPreviewValue],
  )
  const getFormattedValue = React.useCallback(
    (_sheetId: number, rowIndex: number, columnIndex: number) => formatPreviewValue(getPreviewValue(rowIndex, columnIndex)),
    [getPreviewValue],
  )
  const getCellData = React.useCallback<NonNullable<CanvasGridProps<CellData>['getCellData']>>(
    (_sheetId, rowIndex, columnIndex) => {
      const value = getPreviewValue(rowIndex, columnIndex)
      const extendedValue = toExtendedValue(value)
      if (!extendedValue) return undefined
      return {
        ue: extendedValue,
        ev: extendedValue,
        fv: formatPreviewValue(value),
      }
    },
    [getPreviewValue],
  )
  const getEffectiveFormat = React.useCallback(
    (_sheetId: number, rowIndex: number, columnIndex: number) => preview.formats[rowIndex]?.[columnIndex],
    [preview],
  )
  const onChange = React.useCallback<NonNullable<CanvasGridProps['onChange']>>(
    (_sheetId, cell, value) => {
      const field = salesFields[cell.columnIndex]
      const row = sourceRows[cell.rowIndex - 1]
      if (!field || !row || field === 'orderId') return
      store.setCell(salesTableId, row.orderId, field, parseSalesCellValue(field, value))
      bumpEditRevision()
    },
    [sourceRows, store],
  )

  return (
    <section className="panel full-span">
      <div className="canvas-header">
        <SectionHeader
          title="Rows n Columns CanvasGrid"
          meta={`${sourceRows.length} TinyBase rows, ${sheetRowCount} sheet rows`}
        />
        <button type="button" className="button small" onClick={onAddTinyBaseRow}>
          Add source row
        </button>
      </div>
      <div className="canvas-shell">
        <SpreadsheetProvider>
          <CanvasGrid
            sheetId={sheetId}
            rowCount={sheetRowCount}
            columnCount={sheetColumnCount}
            showGridLines
            showHeaders
            defaultColumnWidth={132}
            defaultRowHeight={28}
            columnHeaderHeight={28}
            rowHeaderWidth={44}
            getCellData={getCellData}
            getEffectiveValue={getEffectiveValue}
            getEffectiveExtendedValue={getExtendedValue}
            getUserEnteredExtendedValue={getExtendedValue}
            getFormattedValue={getFormattedValue}
            getEffectiveFormat={getEffectiveFormat}
            getSheetRowCount={() => sheetRowCount}
            getSheetColumnCount={() => sheetColumnCount}
            getDataRowCount={() => sourceRows.length + 1}
            getDataColumnCount={() => preview.columnCount}
            onChange={onChange}
            getSheetName={() => 'Sales'}
            getSheetId={() => sheetId}
            style={{ height: '100%', width: '100%' }}
          />
        </SpreadsheetProvider>
      </div>
    </section>
  )
}

function SectionHeader({ title, meta }: { title: string; meta: string }) {
  return (
    <div className="section-header">
      <h2>{title}</h2>
      <span>{meta}</span>
    </div>
  )
}

function useTinyBaseSalesRows(store: Store): { rows: SalesRow[]; refresh: () => void } {
  const readRows = React.useCallback(
    () => getPivotRowsFromTinyBaseTable<SalesRow>(store, salesTableId, salesFields),
    [store],
  )
  const [rows, setRows] = React.useState<SalesRow[]>(readRows)
  const refresh = React.useCallback(() => setRows(readRows()), [readRows])

  React.useEffect(() => {
    refresh()
    const tableListenerId = store.addTableListener(salesTableId, refresh)
    const rowIdsListenerId = store.addRowIdsListener(salesTableId, refresh)
    return () => {
      store.delListener(tableListenerId)
      store.delListener(rowIdsListenerId)
    }
  }, [refresh, store])

  return React.useMemo(() => ({ rows, refresh }), [refresh, rows])
}

function createSalesStore(): Store {
  return createStore().setTables({
    [salesTableId]: Object.fromEntries(initialSalesRows.map(row => [row.orderId, row])),
  })
}

function addTinyBaseSaleRow(store: Store) {
  const nextIndex = store.getRowIds(salesTableId).length + 1
  const region = ['North', 'South', 'East', 'West'][nextIndex % 4] ?? 'North'
  const product = ['Laptop', 'Monitor', 'Tablet'][nextIndex % 3] ?? 'Laptop'
  const channel = nextIndex % 2 === 0 ? 'Direct' : 'Partner'
  const quarter = `Q${(nextIndex % 4) + 1}`
  const orderId = `${region.toLowerCase()}-${product.toLowerCase()}-${channel.toLowerCase()}-${quarter.toLowerCase()}-${nextIndex}`

  store.setRow(salesTableId, orderId, {
    orderId,
    region,
    product,
    channel,
    quarter,
    sales: 54000 + nextIndex * 3500,
    quantity: 20 + nextIndex,
  })
}

function parseSalesCellValue(field: (typeof salesFields)[number], value: string | boolean): string | number | boolean {
  if (field === 'sales' || field === 'quantity') {
    const numericValue = Number(String(value).replace(/[$,]/g, ''))
    return Number.isFinite(numericValue) ? numericValue : 0
  }
  return value
}

function isSourceCell(rowIndex: number, columnIndex: number, sourceRows: SalesRow[]): boolean {
  return columnIndex < salesFields.length && rowIndex >= 0 && rowIndex <= sourceRows.length
}

function getSourceCellValue(
  store: Store,
  sourceRows: SalesRow[],
  rowIndex: number,
  columnIndex: number,
): PreviewCellValue {
  const field = salesFields[columnIndex]
  if (!field) return undefined
  if (rowIndex === 0) return formatFieldLabel(field)
  const row = sourceRows[rowIndex - 1]
  if (!row) return undefined
  const value = store.getCell(salesTableId, row.orderId, field)
  return isPivotPrimitiveValue(value) ? value : row[field]
}

function toExtendedValue(value: PreviewCellValue): ExtendedValue | undefined {
  if (value == null) return undefined
  if (typeof value === 'number') return { numberValue: value, nv: value }
  if (typeof value === 'boolean') return { boolValue: value, bv: value }
  return { stringValue: value, sv: value }
}

function formatPreviewValue(value: PreviewCellValue): string | undefined {
  if (typeof value === 'number') return formatCurrency(value)
  return value == null ? undefined : String(value)
}

function isPivotPrimitiveValue(value: unknown): value is PivotPrimitive {
  return typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean'
}

function createRowsNColumnsPreview(
  sourceRows: SalesRow[],
  batch: ReturnType<typeof createInMemoryPivotBatch<SalesRow>>,
): {
  columnCount: number
  formats: (CellFormat | undefined)[][]
  values: PreviewCellValue[][]
} {
  const rowCount = Math.max(sourceRows.length + 2, batch.range.endRowIndex + 2)
  const columnCount = Math.max(salesFields.length + 2, batch.range.endColumnIndex + 2)
  const values = Array.from({ length: rowCount }, () => Array.from<PreviewCellValue>({ length: columnCount }))
  const formats = Array.from({ length: rowCount }, () => Array.from<CellFormat | undefined>({ length: columnCount }))
  const sourceHeaderFormat: CellFormat = {
    backgroundColor: '#f1f5f9',
    textFormat: { bold: true, color: '#0f172a' },
  }
  const moneyFormat: CellFormat = {
    numberFormat: { type: 'CURRENCY', pattern: '$#,##0' },
    horizontalAlignment: 'right',
  }

  salesFields.forEach((field, columnIndex) => {
    values[0][columnIndex] = formatFieldLabel(field)
    formats[0][columnIndex] = sourceHeaderFormat
  })
  sourceRows.forEach((row, rowIndex) => {
    salesFields.forEach((field, columnIndex) => {
      values[rowIndex + 1][columnIndex] = row[field]
      if (field === 'sales') {
        formats[rowIndex + 1][columnIndex] = moneyFormat
      }
    })
  })
  batch.values.forEach((row, rowOffset) => {
    row.forEach((cell, columnOffset) => {
      const rowIndex = batch.range.startRowIndex + rowOffset
      const columnIndex = batch.range.startColumnIndex + columnOffset
      values[rowIndex][columnIndex] = cell
      formats[rowIndex][columnIndex] = batch.formats[rowOffset]?.[columnOffset]
    })
  })

  return { columnCount, formats, values }
}

function formatPivotCell(
  cell: string | number | boolean | null,
  aggregation: AggregationFunction,
  metadata: PivotCellMetadata<SalesRow> | undefined,
): string {
  if (typeof cell !== 'number') return cell === null ? '' : String(cell)
  if (aggregation === 'count') return String(cell)
  if (metadata?.valueField === 'sales') return formatCurrency(cell)
  return cell.toLocaleString('en-US')
}

function formatCurrency(value: number): string {
  return value.toLocaleString('en-US', {
    style: 'currency',
    currency: 'USD',
    maximumFractionDigits: 0,
  })
}

function formatFieldLabel(value: string): string {
  if (value === 'all') return 'All'
  return value
    .replace(/([a-z0-9])([A-Z])/g, '$1 $2')
    .replace(/[-_]/g, ' ')
    .replace(/\b\w/g, char => char.toUpperCase())
}

function formatRange(range: {
  sheetId: number
  startRowIndex: number
  endRowIndex: number
  startColumnIndex: number
  endColumnIndex: number
}): string {
  return `S${range.sheetId} R${range.startRowIndex}:${range.endRowIndex} C${range.startColumnIndex}:${range.endColumnIndex}`
}

function isSamePivotCell(
  left: PivotCellMetadata<SalesRow> | null,
  right: PivotCellMetadata<SalesRow> | undefined,
): boolean {
  if (!left || !right) return false
  return (
    samePath(left.rowPath, right.rowPath) && samePath(left.columnPath, right.columnPath) && left.kind === right.kind
  )
}

function samePath(left: PivotPrimitive[], right: PivotPrimitive[]): boolean {
  return left.length === right.length && left.every((value, index) => value === right[index])
}

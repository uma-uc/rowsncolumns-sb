import {
  CanvasGrid,
  CellEditor as DefaultCellEditor,
  SpreadsheetProvider,
  type CanvasGridProps,
  type CellData,
  type CellEditorProps,
  type CellFormat,
  type DataValidationRule,
  type ExtendedValue,
} from '@rowsncolumns/spreadsheet'
import '@rowsncolumns/spreadsheet/dist/spreadsheet.min.css'
import type { Meta, StoryObj } from '@storybook/react-vite'
import * as React from 'react'
import { createStore, type Store } from 'tinybase'

const taskTableId = 'tasks'
const sheetId = 2
const headerDepth = 2
const sheetHeaderOffset = 1
const groupHeaderRowIndex = sheetHeaderOffset
const fieldHeaderRowIndex = sheetHeaderOffset + 1
const dataStartRowIndex = sheetHeaderOffset + headerDepth
const dataStartColumnIndex = sheetHeaderOffset
const rowsncolumnsLicenseKey =
  (import.meta as ImportMeta & { env?: Record<string, string | undefined> }).env?.VITE_ROWSNCOLUMNS_LICENSE_KEY ?? ''

type TaskField = 'title' | 'owner' | 'status' | 'priority' | 'due' | 'budget' | 'approved'
type TaskCellValue = string | number | boolean | null | undefined

type TaskRow = {
  id: string
  parentId: string
  title: string
  owner: string
  status: string
  priority: string
  due: string
  budget: number
  approved: boolean
  sort: number
}

type TaskColumn = {
  field: TaskField
  label: string
  group: string
  width: number
  align?: 'left' | 'right' | 'center'
  options?: readonly string[]
}

type TreeRow = {
  row: TaskRow
  depth: number
  hidden: boolean
  hasChildren: boolean
  expanded: boolean
  childrenCount: number
}

const statusOptions = ['Planned', 'Active', 'Blocked', 'Done'] as const
const priorityOptions = ['Low', 'Medium', 'High', 'Critical'] as const
const ownerOptions = ['Avery', 'Blake', 'Casey', 'Drew', 'Morgan', 'Riley', 'Taylor'] as const

const taskColumns: TaskColumn[] = [
  { field: 'title', label: 'Task', group: 'Work item', width: 220 },
  { field: 'owner', label: 'Owner', group: 'Work item', width: 118, options: ownerOptions },
  { field: 'status', label: 'Status', group: 'Planning', width: 120, options: statusOptions },
  { field: 'priority', label: 'Priority', group: 'Planning', width: 112, options: priorityOptions },
  { field: 'due', label: 'Due', group: 'Planning', width: 112 },
  { field: 'budget', label: 'Budget', group: 'Controls', width: 118, align: 'right' },
  { field: 'approved', label: 'Approved', group: 'Controls', width: 108, align: 'center' },
]

const taskColumnByField = Object.fromEntries(taskColumns.map(column => [column.field, column])) as Record<
  TaskField,
  TaskColumn
>

const initialTaskRows: TaskRow[] = [
  {
    id: 'crm',
    parentId: '',
    title: 'CRM migration',
    owner: 'Avery',
    status: 'Active',
    priority: 'High',
    due: '2026-05-15',
    budget: 185000,
    approved: true,
    sort: 10,
  },
  {
    id: 'crm-discovery',
    parentId: 'crm',
    title: 'Map current account fields',
    owner: 'Casey',
    status: 'Done',
    priority: 'Medium',
    due: '2026-05-03',
    budget: 18000,
    approved: true,
    sort: 10,
  },
  {
    id: 'crm-import',
    parentId: 'crm',
    title: 'Build import validation rules',
    owner: 'Drew',
    status: 'Active',
    priority: 'High',
    due: '2026-05-10',
    budget: 42000,
    approved: true,
    sort: 20,
  },
  {
    id: 'crm-training',
    parentId: 'crm',
    title: 'Draft sales training guide',
    owner: 'Morgan',
    status: 'Planned',
    priority: 'Medium',
    due: '2026-05-14',
    budget: 15000,
    approved: false,
    sort: 30,
  },
  {
    id: 'billing',
    parentId: '',
    title: 'Billing portal',
    owner: 'Blake',
    status: 'Blocked',
    priority: 'Critical',
    due: '2026-05-22',
    budget: 240000,
    approved: false,
    sort: 20,
  },
  {
    id: 'billing-api',
    parentId: 'billing',
    title: 'Finalize payment API contract',
    owner: 'Riley',
    status: 'Blocked',
    priority: 'Critical',
    due: '2026-05-08',
    budget: 68000,
    approved: false,
    sort: 10,
  },
  {
    id: 'billing-invoice',
    parentId: 'billing',
    title: 'Invoice preview editor',
    owner: 'Taylor',
    status: 'Active',
    priority: 'High',
    due: '2026-05-17',
    budget: 54000,
    approved: true,
    sort: 20,
  },
  {
    id: 'analytics',
    parentId: '',
    title: 'Analytics launch',
    owner: 'Morgan',
    status: 'Planned',
    priority: 'Medium',
    due: '2026-05-29',
    budget: 98000,
    approved: true,
    sort: 30,
  },
  {
    id: 'analytics-model',
    parentId: 'analytics',
    title: 'Revenue attribution model',
    owner: 'Casey',
    status: 'Planned',
    priority: 'Medium',
    due: '2026-05-18',
    budget: 36000,
    approved: true,
    sort: 10,
  },
  {
    id: 'analytics-dashboard',
    parentId: 'analytics',
    title: 'Executive dashboard QA',
    owner: 'Avery',
    status: 'Planned',
    priority: 'Low',
    due: '2026-05-24',
    budget: 22000,
    approved: false,
    sort: 20,
  },
]

const meta: Meta = {
  title: 'RowsnColumns/Data Grid Parity',
  parameters: {
    layout: 'fullscreen',
  },
}

export default meta

type Story = StoryObj

export const TreeGridGroupedHeadersAndEditors: Story = {
  render: () => <DataGridParityStory />,
}

export const ProfessionalLicenseEnabled: Story = {
  render: () => <DataGridParityStory licenseKey={rowsncolumnsLicenseKey} showLicenseRuntime />,
}

function DataGridParityStory({
  licenseKey,
  showLicenseRuntime = false,
}: {
  licenseKey?: string
  showLicenseRuntime?: boolean
}) {
  const store = React.useMemo(createTaskStore, [])
  const { refresh: refreshTaskRows, rows: taskRows } = useTinyBaseTaskRows(store)
  const [expandedIds, setExpandedIds] = React.useState<Set<string>>(() => new Set(['crm', 'billing']))
  const [columnOrder, setColumnOrder] = React.useState<TaskField[]>(() => taskColumns.map(column => column.field))
  const [selectedColumn, setSelectedColumn] = React.useState<TaskField>('status')
  const [selectedParentId, setSelectedParentId] = React.useState('crm')
  const columns = React.useMemo(() => columnOrder.map(field => taskColumnByField[field]), [columnOrder])
  const treeRows = React.useMemo(() => buildTreeRows(taskRows, expandedIds), [expandedIds, taskRows])
  const visibleTreeRows = React.useMemo(() => treeRows.filter(entry => !entry.hidden), [treeRows])
  const sheetRowCount = headerDepth + treeRows.length + 10
  const sheetColumnCount = columns.length + 2

  const handleAddChildTask = React.useCallback(() => {
    addTinyBaseTaskRow(store, selectedParentId)
    setExpandedIds(previous => new Set(previous).add(selectedParentId))
    refreshTaskRows()
  }, [refreshTaskRows, selectedParentId, store])

  const handleMoveSelectedColumn = React.useCallback(
    (direction: -1 | 1) => {
      setColumnOrder(previous => moveColumnField(previous, selectedColumn, direction))
    },
    [selectedColumn],
  )

  const handleNativeColumnMove = React.useCallback<NonNullable<CanvasGridProps['onMoveColumns']>>(
    (_sheetId, dims, toColumn) => {
      const dataDims = dims.map(index => index - dataStartColumnIndex)
      setColumnOrder(previous => moveColumnIndices(previous, dataDims, toColumn - dataStartColumnIndex))
    },
    [],
  )

  const handleExpandCollapse = React.useCallback<NonNullable<CanvasGridProps['onExpandCollapse']>>(
    (_sheetId, _cell, groupKeys, newExpandedState) => {
      const rowId = groupKeys[0]
      if (!rowId) return
      setExpandedIds(previous => {
        const next = new Set(previous)
        if (newExpandedState) {
          next.add(rowId)
        } else {
          next.delete(rowId)
        }
        return next
      })
    },
    [],
  )

  return (
    <div className="page">
      <header className="topbar">
        <div>
          <h1 className="title">Rows n Columns Data Grid</h1>
          <div className="badge-row">
            <Badge tone="info">TinyBase source</Badge>
            <Badge tone="success">Tree rows</Badge>
            <Badge tone="neutral">Grouped headers</Badge>
            <Badge tone="success">Dropdown editors</Badge>
            {licenseKey ? <Badge tone="success">Professional license</Badge> : null}
          </div>
        </div>
        <div className="topbar-actions">
          <button type="button" className="button" onClick={() => setExpandedIds(new Set(taskRows.map(row => row.id)))}>
            Expand all
          </button>
          <button type="button" className="button" onClick={() => setExpandedIds(new Set())}>
            Collapse all
          </button>
        </div>
      </header>

      <main className="app-grid">
        <aside className="sidebar">
          <div className="control-stack">
            <ControlSelect
              label="Add child under"
              value={selectedParentId}
              options={taskRows.map(row => ({ label: row.title, value: row.id }))}
              onChange={setSelectedParentId}
            />
            <button type="button" className="button" onClick={handleAddChildTask}>
              Add TinyBase row
            </button>
            <ControlSelect
              label="Column"
              value={selectedColumn}
              options={columns.map(column => ({ label: column.label, value: column.field }))}
              onChange={value => setSelectedColumn(value as TaskField)}
            />
            <div className="split-buttons">
              <button type="button" className="button small" onClick={() => handleMoveSelectedColumn(-1)}>
                Move left
              </button>
              <button type="button" className="button small" onClick={() => handleMoveSelectedColumn(1)}>
                Move right
              </button>
            </div>
          </div>

          <div className="runtime">
            <div className="runtime-title">Grid Runtime</div>
            <div className="status-stack">
              <StatusLine label="TinyBase rows" value={String(taskRows.length)} tone="info" />
              <StatusLine label="Visible rows" value={String(visibleTreeRows.length)} tone="success" />
              <StatusLine label="Canvas rows" value={String(sheetRowCount)} tone="neutral" />
              {showLicenseRuntime ? (
                <StatusLine
                  label="License key"
                  value={licenseKey ? 'env loaded' : 'missing'}
                  tone={licenseKey ? 'success' : 'neutral'}
                />
              ) : null}
            </div>
          </div>

          <div className="runtime">
            <div className="runtime-title">Column Order</div>
            <div className="column-pill-list">
              {columns.map(column => (
                <button
                  key={column.field}
                  type="button"
                  className={column.field === selectedColumn ? 'column-pill selected' : 'column-pill'}
                  onClick={() => setSelectedColumn(column.field)}
                >
                  {column.label}
                </button>
              ))}
            </div>
          </div>
        </aside>

        <section className="content">
          <div className="parity-grid">
            <section className="panel full-span">
              <div className="canvas-header">
                <SectionHeader
                  title="CanvasGrid with TinyBase-backed Rows"
                  meta={`${taskRows.length} TinyBase rows, ${visibleTreeRows.length} visible tree rows`}
                />
              </div>
              <TaskCanvas
                columns={columns}
                store={store}
                treeRows={treeRows}
                licenseKey={licenseKey}
                onChangeRows={refreshTaskRows}
                onExpandCollapse={handleExpandCollapse}
                onMoveColumns={handleNativeColumnMove}
              />
            </section>
            <TinyBaseTaskTable rows={taskRows} />
            <FeatureNotes />
          </div>
        </section>
      </main>
    </div>
  )
}

function TaskCanvas({
  columns,
  onChangeRows,
  onExpandCollapse,
  onMoveColumns,
  store,
  treeRows,
  licenseKey,
}: {
  columns: TaskColumn[]
  licenseKey?: string
  onChangeRows: () => void
  onExpandCollapse: NonNullable<CanvasGridProps['onExpandCollapse']>
  onMoveColumns: NonNullable<CanvasGridProps['onMoveColumns']>
  store: Store
  treeRows: TreeRow[]
}) {
  const columnMetadata = React.useMemo<NonNullable<CanvasGridProps['columnMetadata']>>(
    () => {
      const metadata: NonNullable<CanvasGridProps['columnMetadata']> = []
      columns.forEach((column, index) => {
        metadata[dataStartColumnIndex + index] = { size: column.width }
      })
      return metadata
    },
    [columns],
  )
  const rowMetadata = React.useMemo<NonNullable<CanvasGridProps['rowMetadata']>>(() => {
    const metadata: NonNullable<CanvasGridProps['rowMetadata']> = []
    metadata[groupHeaderRowIndex] = { size: 34 }
    metadata[fieldHeaderRowIndex] = { size: 34 }
    treeRows.forEach((entry, index) => {
      metadata[dataStartRowIndex + index] = entry.hidden ? { hiddenByUser: true } : { size: entry.hasChildren ? 34 : 30 }
    })
    return metadata
  }, [treeRows])
  const merges = React.useMemo(() => createHeaderMerges(columns), [columns])
  const sheetRowCount = dataStartRowIndex + treeRows.length + 10
  const sheetColumnCount = dataStartColumnIndex + columns.length + 2

  const getGridValue = React.useCallback(
    (rowIndex: number, columnIndex: number): TaskCellValue => {
      const column = columns[columnIndex - dataStartColumnIndex]
      if (!column) return undefined
      if (rowIndex === groupHeaderRowIndex) return column.group
      if (rowIndex === fieldHeaderRowIndex) return column.label
      const entry = treeRows[rowIndex - dataStartRowIndex]
      if (!entry) return undefined
      return entry.row[column.field]
    },
    [columns, treeRows],
  )
  const getEffectiveValue = React.useCallback(
    (_sheetId: number, rowIndex: number, columnIndex: number) => getGridValue(rowIndex, columnIndex),
    [getGridValue],
  )
  const getExtendedValue = React.useCallback(
    (_sheetId: number, rowIndex: number, columnIndex: number) => toExtendedValue(getGridValue(rowIndex, columnIndex)),
    [getGridValue],
  )
  const getFormattedValue = React.useCallback(
    (_sheetId: number, rowIndex: number, columnIndex: number) => formatTaskValue(getGridValue(rowIndex, columnIndex)),
    [getGridValue],
  )
  const getEffectiveFormat = React.useCallback<NonNullable<CanvasGridProps['getEffectiveFormat']>>(
    (_sheetId, rowIndex, columnIndex) => {
      const column = columns[columnIndex - dataStartColumnIndex]
      if (!column) return undefined
      if (rowIndex === groupHeaderRowIndex) return groupHeaderFormat
      if (rowIndex === fieldHeaderRowIndex) return fieldHeaderFormat
      const entry = treeRows[rowIndex - dataStartRowIndex]
      if (!entry) return undefined
      return getTaskCellFormat(entry, column)
    },
    [columns, treeRows],
  )
  const getDataValidation = React.useCallback<NonNullable<CanvasGridProps['getDataValidation']>>(
    (_sheetId, rowIndex, columnIndex) => {
      if (rowIndex < dataStartRowIndex) return undefined
      const column = columns[columnIndex - dataStartColumnIndex]
      if (!column) return undefined
      if (column.field === 'approved') return booleanRule
      if (column.options) return createListRule(column.options)
      return undefined
    },
    [columns],
  )
  const getCellData = React.useCallback<NonNullable<CanvasGridProps<CellData>['getCellData']>>(
    (_sheetId, rowIndex, columnIndex) => {
      const value = getGridValue(rowIndex, columnIndex)
      const extendedValue = toExtendedValue(value)
      if (!extendedValue) return undefined
      const column = columns[columnIndex - dataStartColumnIndex]
      const entry = treeRows[rowIndex - dataStartRowIndex]
      return {
        ue: extendedValue,
        ev: extendedValue,
        fv: formatTaskValue(value),
        dataValidation: rowIndex >= dataStartRowIndex ? getDataValidation(sheetId, rowIndex, columnIndex) : undefined,
        expandable: column?.field === 'title' && entry?.hasChildren ? true : undefined,
        expanded: column?.field === 'title' && entry?.hasChildren ? entry.expanded : undefined,
        groupKeys: column?.field === 'title' && entry?.hasChildren ? [entry.row.id] : undefined,
        childrenCount: column?.field === 'title' && entry?.hasChildren ? entry.childrenCount : undefined,
      }
    },
    [columns, getDataValidation, getGridValue, treeRows],
  )
  const onChange = React.useCallback<NonNullable<CanvasGridProps['onChange']>>(
    (_sheetId, cell, value) => {
      if (cell.rowIndex < dataStartRowIndex || cell.columnIndex < dataStartColumnIndex) return
      const entry = treeRows[cell.rowIndex - dataStartRowIndex]
      const column = columns[cell.columnIndex - dataStartColumnIndex]
      if (!entry || !column) return
      store.setCell(taskTableId, entry.row.id, column.field, parseTaskCellValue(column.field, value))
      onChangeRows()
    },
    [columns, onChangeRows, store, treeRows],
  )
  const TaskGridCellEditor = React.useCallback(
    (props: CellEditorProps) => {
      const entry = treeRows[props.cell.rowIndex - dataStartRowIndex]
      const column = columns[props.cell.columnIndex - dataStartColumnIndex]
      if (!entry || !column || props.cell.rowIndex < dataStartRowIndex) {
        return <DefaultCellEditor {...props} />
      }
      if (column.options) {
        return (
          <InlineCellEditor props={props}>
            <select
              autoFocus
              value={String(entry.row[column.field])}
              onChange={event => {
                store.setCell(taskTableId, entry.row.id, column.field, event.target.value)
                onChangeRows()
                props.onCancel?.()
              }}
              onKeyDown={event => {
                if (event.key === 'Escape') props.onCancel?.(event)
              }}
            >
              {column.options.map(option => (
                <option key={option} value={option}>
                  {option}
                </option>
              ))}
            </select>
          </InlineCellEditor>
        )
      }
      if (column.field === 'approved') {
        return (
          <InlineCellEditor props={props}>
            <label>
              <input
                autoFocus
                checked={entry.row.approved}
                type="checkbox"
                onChange={event => {
                  store.setCell(taskTableId, entry.row.id, column.field, event.target.checked)
                  onChangeRows()
                  props.onCancel?.()
                }}
                onKeyDown={event => {
                  if (event.key === 'Escape') props.onCancel?.(event)
                }}
              />
              <span>{entry.row.approved ? 'Approved' : 'Not approved'}</span>
            </label>
          </InlineCellEditor>
        )
      }
      return <DefaultCellEditor {...props} />
    },
    [columns, onChangeRows, store, treeRows],
  )

  return (
    <div className="canvas-shell parity-canvas">
      <SpreadsheetProvider>
        <CanvasGrid
          licenseKey={licenseKey}
          sheetId={sheetId}
          rowCount={sheetRowCount}
          columnCount={sheetColumnCount}
          activeCell={{ rowIndex: dataStartRowIndex, columnIndex: dataStartColumnIndex }}
          showGridLines
          showHeaders
          frozenRowCount={headerDepth}
          defaultColumnWidth={132}
          defaultRowHeight={30}
          columnHeaderHeight={28}
          rowHeaderWidth={48}
          columnMetadata={columnMetadata}
          rowMetadata={rowMetadata}
          merges={merges}
          getCellData={getCellData}
          getEffectiveValue={getEffectiveValue}
          getEffectiveExtendedValue={getExtendedValue}
          getUserEnteredExtendedValue={getExtendedValue}
          getFormattedValue={getFormattedValue}
          getEffectiveFormat={getEffectiveFormat}
          getDataValidation={getDataValidation}
          getSheetRowCount={() => sheetRowCount}
          getSheetColumnCount={() => sheetColumnCount}
          getDataRowCount={() => headerDepth + treeRows.filter(entry => !entry.hidden).length}
          getDataColumnCount={() => columns.length}
          onChange={onChange}
          onExpandCollapse={onExpandCollapse}
          onMoveColumns={onMoveColumns}
          CellEditor={TaskGridCellEditor}
          getSheetName={() => 'Tasks'}
          getSheetId={() => sheetId}
          style={{ height: '100%', width: '100%' }}
        />
      </SpreadsheetProvider>
    </div>
  )
}

function InlineCellEditor({ children, props }: { children: React.ReactNode; props: CellEditorProps }) {
  return (
    <div
      style={{
        alignItems: 'center',
        background: '#ffffff',
        border: '2px solid #2563eb',
        boxShadow: '0 8px 24px rgba(15, 23, 42, 0.18)',
        display: 'flex',
        gap: 8,
        height: Math.max(props.position.height, 34),
        left: props.position.x,
        padding: '0 8px',
        position: 'absolute',
        top: props.position.y,
        width: Math.max(props.position.width, 180),
        zIndex: 20,
      }}
      onMouseDown={event => event.stopPropagation()}
    >
      {children}
    </div>
  )
}

function TinyBaseTaskTable({ rows }: { rows: TaskRow[] }) {
  return (
    <section className="panel">
      <SectionHeader title="TinyBase Task Table" meta={`${rows.length} rows`} />
      <div className="table-wrap">
        <table className="table">
          <thead>
            <tr>
              <th>Task</th>
              <th>Parent</th>
              <th>Status</th>
              <th>Priority</th>
              <th className="right">Budget</th>
            </tr>
          </thead>
          <tbody>
            {rows.map(row => (
              <tr key={row.id}>
                <td>{row.title}</td>
                <td className="mono">{row.parentId || 'root'}</td>
                <td>{row.status}</td>
                <td>{row.priority}</td>
                <td className="right">{formatCurrency(row.budget)}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </section>
  )
}

function FeatureNotes() {
  return (
    <section className="panel">
      <SectionHeader title="Mapped Capabilities" meta="Rows n Columns primitives" />
      <div className="capability-list">
        <Capability title="Tree rows" value="CellData.expandable plus rowMetadata.hiddenByUser" />
        <Capability title="Grouped headers" value="Frozen rows plus merged header ranges" />
        <Capability title="Column order" value="Controlled column array plus onMoveColumns" />
        <Capability title="Editors" value="ONE_OF_LIST and BOOLEAN data validation" />
      </div>
    </section>
  )
}

function Capability({ title, value }: { title: string; value: string }) {
  return (
    <div className="capability">
      <span>{title}</span>
      <strong>{value}</strong>
    </div>
  )
}

function Badge({ children, tone }: { children: React.ReactNode; tone: 'info' | 'success' | 'neutral' }) {
  return <span className={`badge ${tone}`}>{children}</span>
}

function ControlSelect({
  label,
  value,
  options,
  onChange,
}: {
  label: string
  value: string
  options: readonly { label: string; value: string }[]
  onChange: (value: string) => void
}) {
  return (
    <label className="control">
      <span>{label}</span>
      <select value={value} onChange={event => onChange(event.target.value)}>
        {options.map(option => (
          <option key={option.value} value={option.value}>
            {option.label}
          </option>
        ))}
      </select>
    </label>
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

function StatusLine({ label, value, tone }: { label: string; value: string; tone: 'success' | 'info' | 'neutral' }) {
  return (
    <div className="status-line">
      <span>{label}</span>
      <Badge tone={tone}>{value}</Badge>
    </div>
  )
}

function useTinyBaseTaskRows(store: Store): { rows: TaskRow[]; refresh: () => void } {
  const readRows = React.useCallback(() => getTaskRowsFromTinyBase(store), [store])
  const [rows, setRows] = React.useState<TaskRow[]>(readRows)
  const refresh = React.useCallback(() => setRows(readRows()), [readRows])

  React.useEffect(() => {
    refresh()
    const tableListenerId = store.addTableListener(taskTableId, refresh)
    const rowIdsListenerId = store.addRowIdsListener(taskTableId, refresh)
    return () => {
      store.delListener(tableListenerId)
      store.delListener(rowIdsListenerId)
    }
  }, [refresh, store])

  return React.useMemo(() => ({ refresh, rows }), [refresh, rows])
}

function createTaskStore(): Store {
  return createStore().setTables({
    [taskTableId]: Object.fromEntries(initialTaskRows.map(row => [row.id, row])),
  })
}

function getTaskRowsFromTinyBase(store: Store): TaskRow[] {
  return store
    .getRowIds(taskTableId)
    .map(rowId => {
      const row = store.getRow(taskTableId, rowId)
      return {
        id: String(row.id ?? rowId),
        parentId: String(row.parentId ?? ''),
        title: String(row.title ?? ''),
        owner: String(row.owner ?? ''),
        status: String(row.status ?? 'Planned'),
        priority: String(row.priority ?? 'Medium'),
        due: String(row.due ?? ''),
        budget: Number(row.budget ?? 0),
        approved: row.approved === true,
        sort: Number(row.sort ?? 0),
      }
    })
    .sort((left, right) => {
      if (left.parentId !== right.parentId) return left.parentId.localeCompare(right.parentId)
      return left.sort - right.sort || left.title.localeCompare(right.title)
    })
}

function addTinyBaseTaskRow(store: Store, parentId: string) {
  const parent = store.getRow(taskTableId, parentId)
  const nextIndex = store.getRowIds(taskTableId).length + 1
  const rowId = `task-${nextIndex}`
  const siblingCount = store.getRowIds(taskTableId).filter(id => store.getCell(taskTableId, id, 'parentId') === parentId).length
  store.setRow(taskTableId, rowId, {
    id: rowId,
    parentId,
    title: `Follow-up ${nextIndex}`,
    owner: ownerOptions[nextIndex % ownerOptions.length],
    status: statusOptions[nextIndex % statusOptions.length],
    priority: priorityOptions[nextIndex % priorityOptions.length],
    due: '2026-05-31',
    budget: 12000 + nextIndex * 1500,
    approved: false,
    sort: (siblingCount + 1) * 10,
  })
  if (parent.id === undefined) {
    store.setCell(taskTableId, parentId, 'id', parentId)
  }
}

function buildTreeRows(rows: TaskRow[], expandedIds: Set<string>): TreeRow[] {
  const childrenByParent = new Map<string, TaskRow[]>()
  rows.forEach(row => {
    const parentId = row.parentId || ''
    const children = childrenByParent.get(parentId) ?? []
    children.push(row)
    childrenByParent.set(parentId, children)
  })
  childrenByParent.forEach(children => {
    children.sort((left, right) => left.sort - right.sort || left.title.localeCompare(right.title))
  })

  const result: TreeRow[] = []
  const walk = (parentId: string, depth: number, parentHidden: boolean) => {
    const children = childrenByParent.get(parentId) ?? []
    children.forEach(row => {
      const childRows = childrenByParent.get(row.id) ?? []
      const hasChildren = childRows.length > 0
      const expanded = expandedIds.has(row.id)
      const hidden = parentHidden
      result.push({
        row,
        depth,
        hidden,
        hasChildren,
        expanded,
        childrenCount: childRows.length,
      })
      walk(row.id, depth + 1, hidden || (hasChildren && !expanded))
    })
  }
  walk('', 0, false)
  return result
}

function createHeaderMerges(columns: TaskColumn[]): NonNullable<CanvasGridProps['merges']> {
  const ranges: NonNullable<CanvasGridProps['merges']> = []
  let startColumnIndex = 0
  while (startColumnIndex < columns.length) {
    const group = columns[startColumnIndex]?.group
    let endColumnIndex = startColumnIndex
    while (endColumnIndex + 1 < columns.length && columns[endColumnIndex + 1]?.group === group) {
      endColumnIndex += 1
    }
    if (endColumnIndex > startColumnIndex) {
      ranges.push({
        startRowIndex: groupHeaderRowIndex,
        endRowIndex: groupHeaderRowIndex,
        startColumnIndex: dataStartColumnIndex + startColumnIndex,
        endColumnIndex: dataStartColumnIndex + endColumnIndex,
      })
    }
    startColumnIndex = endColumnIndex + 1
  }
  return ranges
}

function moveColumnField(columns: TaskField[], field: TaskField, direction: -1 | 1): TaskField[] {
  const fromIndex = columns.indexOf(field)
  const toIndex = fromIndex + direction
  if (fromIndex < 0 || toIndex < 0 || toIndex >= columns.length) return columns
  const next = [...columns]
  next.splice(fromIndex, 1)
  next.splice(toIndex, 0, field)
  return next
}

function moveColumnIndices(columns: TaskField[], dims: number[], toColumn: number): TaskField[] {
  const movingIndices = [...new Set(dims)]
    .filter(index => index >= 0 && index < columns.length)
    .sort((left, right) => left - right)
  if (!movingIndices.length) return columns

  const movingSet = new Set(movingIndices)
  const movingFields = movingIndices.map(index => columns[index])
  const remainingFields = columns.filter((_field, index) => !movingSet.has(index))
  const movedBeforeTarget = movingIndices.filter(index => index < toColumn).length
  const insertIndex = Math.max(0, Math.min(remainingFields.length, toColumn - movedBeforeTarget))
  return [
    ...remainingFields.slice(0, insertIndex),
    ...movingFields,
    ...remainingFields.slice(insertIndex),
  ] as TaskField[]
}

function toExtendedValue(value: TaskCellValue): ExtendedValue | undefined {
  if (value == null) return undefined
  if (typeof value === 'number') return { numberValue: value, nv: value }
  if (typeof value === 'boolean') return { boolValue: value, bv: value }
  return { stringValue: value, sv: value }
}

function formatTaskValue(value: TaskCellValue): string | undefined {
  if (value == null) return undefined
  if (typeof value === 'number') return formatCurrency(value)
  return String(value)
}

function parseTaskCellValue(field: TaskField, value: string | boolean): string | number | boolean {
  if (field === 'approved') return value === true || String(value).toLowerCase() === 'true'
  if (field === 'budget') {
    const numericValue = Number(String(value).replace(/[$,]/g, ''))
    return Number.isFinite(numericValue) ? numericValue : 0
  }
  return String(value)
}

function createListRule(options: readonly string[]): DataValidationRule {
  return {
    condition: {
      type: 'ONE_OF_LIST',
      values: options.map(option => ({ userEnteredValue: option })),
    },
    displayStyle: 'arrow',
    allowBlank: false,
  }
}

const booleanRule: DataValidationRule = {
  condition: {
    type: 'BOOLEAN',
  },
}

const groupHeaderFormat: CellFormat = {
  backgroundColor: '#dbeafe',
  horizontalAlignment: 'center',
  verticalAlignment: 'middle',
  textFormat: { bold: true, color: '#1e3a8a' },
}

const fieldHeaderFormat: CellFormat = {
  backgroundColor: '#f1f5f9',
  verticalAlignment: 'middle',
  textFormat: { bold: true, color: '#0f172a' },
}

function getTaskCellFormat(entry: TreeRow, column: TaskColumn): CellFormat {
  const backgroundColor = entry.hasChildren ? '#ecfdf5' : undefined
  const textFormat = entry.hasChildren ? { bold: true, color: '#14532d' } : undefined
  if (column.field === 'title') {
    return {
      backgroundColor,
      indent: Math.min(entry.depth, 6),
      textFormat,
      verticalAlignment: 'middle',
    }
  }
  if (column.field === 'budget') {
    return {
      backgroundColor,
      horizontalAlignment: 'right',
      numberFormat: { type: 'CURRENCY', pattern: '$#,##0' },
      textFormat,
    }
  }
  if (column.field === 'approved') {
    return {
      backgroundColor,
      horizontalAlignment: 'center',
      textFormat,
    }
  }
  if (column.options) {
    return {
      backgroundColor,
      textFormat,
    }
  }
  return {
    backgroundColor,
    horizontalAlignment: column.align,
    textFormat,
  }
}

function formatCurrency(value: number): string {
  return value.toLocaleString('en-US', {
    style: 'currency',
    currency: 'USD',
    maximumFractionDigits: 0,
  })
}

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

const sheetId = 4
const rowsncolumnsLicenseKey =
  (import.meta as ImportMeta & { env?: Record<string, string | undefined> }).env?.VITE_ROWSNCOLUMNS_LICENSE_KEY ?? ''

type LicenseCellValue = string | number | boolean | null | undefined

const values: LicenseCellValue[][] = [
  ['Professional License Smoke Test', ''],
  ['License key', rowsncolumnsLicenseKey ? 'env loaded' : 'missing'],
  ['CanvasGrid prop', rowsncolumnsLicenseKey ? 'licenseKey passed' : 'not passed'],
  ['Expected tier', 'professional'],
  ['Notes', 'This story is independent of the tree/grid parity prototype.'],
]

const labelFormat: CellFormat = {
  backgroundColor: '#f8fafc',
  textFormat: { bold: true, color: '#334155' },
}

const valueFormat: CellFormat = {
  textFormat: { color: '#0f172a' },
}

const successFormat: CellFormat = {
  backgroundColor: '#dcfce7',
  textFormat: { bold: true, color: '#166534' },
}

const neutralFormat: CellFormat = {
  backgroundColor: '#e2e8f0',
  textFormat: { bold: true, color: '#334155' },
}

const formats: (CellFormat | undefined)[][] = [
  [{ backgroundColor: '#dbeafe', textFormat: { bold: true, color: '#1e3a8a' } }],
  [labelFormat, rowsncolumnsLicenseKey ? successFormat : neutralFormat],
  [labelFormat, rowsncolumnsLicenseKey ? successFormat : neutralFormat],
  [labelFormat, valueFormat],
  [labelFormat, valueFormat],
]

const meta: Meta = {
  title: 'RowsnColumns/Professional License',
  parameters: {
    layout: 'fullscreen',
  },
}

export default meta

type Story = StoryObj

export const LicenseKeyFromEnv: Story = {
  render: () => <ProfessionalLicenseStory />,
}

function ProfessionalLicenseStory() {
  const rowCount = 12
  const columnCount = 5
  const getValue = React.useCallback((rowIndex: number, columnIndex: number) => values[rowIndex]?.[columnIndex], [])
  const getEffectiveValue = React.useCallback(
    (_sheetId: number, rowIndex: number, columnIndex: number) => getValue(rowIndex, columnIndex),
    [getValue],
  )
  const getExtendedValue = React.useCallback(
    (_sheetId: number, rowIndex: number, columnIndex: number) => toExtendedValue(getValue(rowIndex, columnIndex)),
    [getValue],
  )
  const getFormattedValue = React.useCallback(
    (_sheetId: number, rowIndex: number, columnIndex: number) => formatValue(getValue(rowIndex, columnIndex)),
    [getValue],
  )
  const getCellData = React.useCallback<NonNullable<CanvasGridProps<CellData>['getCellData']>>(
    (_sheetId, rowIndex, columnIndex) => {
      const value = getValue(rowIndex, columnIndex)
      const extendedValue = toExtendedValue(value)
      if (!extendedValue) return undefined
      return {
        ue: extendedValue,
        ev: extendedValue,
        fv: formatValue(value),
      }
    },
    [getValue],
  )

  return (
    <div className="page">
      <header className="topbar">
        <div>
          <h1 className="title">Rows n Columns Professional License</h1>
          <div className="badge-row">
            <Badge tone={rowsncolumnsLicenseKey ? 'success' : 'neutral'}>
              {rowsncolumnsLicenseKey ? 'env loaded' : 'env missing'}
            </Badge>
            <Badge tone="info">independent story</Badge>
          </div>
        </div>
      </header>

      <main className="content">
        <section className="panel">
          <div className="canvas-header">
            <SectionHeader title="CanvasGrid License Smoke Test" meta="VITE_ROWSNCOLUMNS_LICENSE_KEY" />
          </div>
          <div className="canvas-shell">
            <SpreadsheetProvider>
              <CanvasGrid
                licenseKey={rowsncolumnsLicenseKey}
                sheetId={sheetId}
                rowCount={rowCount}
                columnCount={columnCount}
                showGridLines
                showHeaders
                defaultColumnWidth={180}
                defaultRowHeight={34}
                columnHeaderHeight={28}
                rowHeaderWidth={48}
                merges={[{ startRowIndex: 0, endRowIndex: 0, startColumnIndex: 0, endColumnIndex: 1 }]}
                getCellData={getCellData}
                getEffectiveValue={getEffectiveValue}
                getEffectiveExtendedValue={getExtendedValue}
                getUserEnteredExtendedValue={getExtendedValue}
                getFormattedValue={getFormattedValue}
                getEffectiveFormat={(_sheetId, rowIndex, columnIndex) => formats[rowIndex]?.[columnIndex]}
                getSheetRowCount={() => rowCount}
                getSheetColumnCount={() => columnCount}
                getDataRowCount={() => values.length}
                getDataColumnCount={() => 2}
                getSheetName={() => 'License'}
                getSheetId={() => sheetId}
                style={{ height: '100%', width: '100%' }}
              />
            </SpreadsheetProvider>
          </div>
        </section>
      </main>
    </div>
  )
}

function Badge({ children, tone }: { children: React.ReactNode; tone: 'info' | 'success' | 'neutral' }) {
  return <span className={`badge ${tone}`}>{children}</span>
}

function SectionHeader({ title, meta }: { title: string; meta: string }) {
  return (
    <div className="section-header">
      <h2>{title}</h2>
      <span>{meta}</span>
    </div>
  )
}

function toExtendedValue(value: LicenseCellValue): ExtendedValue | undefined {
  if (value == null) return undefined
  if (typeof value === 'number') return { numberValue: value, nv: value }
  if (typeof value === 'boolean') return { boolValue: value, bv: value }
  return { stringValue: value, sv: value }
}

function formatValue(value: LicenseCellValue): string | undefined {
  return value == null ? undefined : String(value)
}

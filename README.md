# Rows n Columns Storybook

Standalone Storybook project for discussing Rows n Columns spreadsheet and pivot behavior.

The examples use:

- `@rowsncolumns/spreadsheet` for `CanvasGrid.`
- TinyBase as the mutable source table
- an in-memory pivot transform, without DuckDB

## Run

```sh
pnpm install
pnpm storybook
```

Storybook runs on port `6009` by default.

## Notes

Pivot output is derived from TinyBase rows and is read-only in these examples. Edit source rows in the `CanvasGrid` or add rows with the provided buttons; the in-memory pivot result recomputes from TinyBase.

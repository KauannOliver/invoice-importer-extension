# Invoice Importer Extension

A browser extension that validates spreadsheet rows and fills an invoice form after manual login.

**JavaScript, Chrome Extensions Manifest V3, SheetJS**

## What it does

- Reads Excel/CSV data using bundled SheetJS.
- Validates rows before filling form fields.
- Reports progress and interrupts on errors for operator review.

## Workflow

```mermaid
flowchart LR
  N0["Excel or CSV"] --> N1
  N1["Validation"] --> N2
  N2["Authenticated form"] --> N3
  N3["Progress or stop on error"]
```

## Local setup and use

Open `chrome://extensions`, enable Developer mode and load the `extension/` folder unpacked. Review `extension/manifest.json` and restrict its host patterns to your authorized test portal before enabling it. Log in manually and use synthetic input for the first run.

## Boundaries and limitations

DOM selectors and form rules are portal-specific. The extension is separate from the Python `spreadsheet-to-portal` application. SheetJS retains its upstream license notice.

This is a source snapshot, not a claim of production readiness. Business results, dates and unmeasured performance are intentionally not asserted. Application UI language is preserved.

## License

No new distribution license is granted for this corporate source edition. Public visibility does not grant reuse rights; obtain permission from the relevant rights holder. Existing third-party notices remain applicable.

## Verification record

See [snapshot checks and remaining limitations](docs/VERIFICATION.md).

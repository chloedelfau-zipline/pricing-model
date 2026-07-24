# BOM Supplier Optimizer

A local web app for turning a CSV, TSV, XLSX, or XLSM bill of materials into a supplier recommendation report. It calculates required buy quantity as:

```text
required quantity = BOM quantity per build x demand x (1 + buffer %)
```

The app uses the official Nexar/Octopart GraphQL API when `NEXAR_CLIENT_ID` and `NEXAR_CLIENT_SECRET` are set. Without credentials, it runs in deterministic demo mode so the workflow can still be tested.

## Run

```bash
npm start
```

Open `http://localhost:4173`.

## Octopart API Setup

1. Create an application in the Altium/Nexar developer portal.
2. Copy `.env.example` to `.env` or export the variables in your shell.
3. Start the app with those environment variables present:

```bash
export NEXAR_CLIENT_ID="..."
export NEXAR_CLIENT_SECRET="..."
npm start
```

The backend exchanges the credentials for an OAuth token using the `supply.domain` scope and sends BOM lookups to the Nexar GraphQL endpoint. Secrets stay server-side.

This app uses Nexar's client credentials flow for supply data, so it does not need a browser redirect URL. If you use Nexar's GraphQL IDE or add an interactive login flow later, configure redirect URLs there separately; the BOM app itself only needs the client ID and client secret.

## BOM Columns

The importer recognizes common names for these columns:

| Field | Common column names |
| --- | --- |
| Zipline PN | `ID`, `Item ID`, `Item Number`, `Zipline PN`, `ZL PN` |
| Line | `Line`, `Item`, `Reference` |
| Description | `Description`, `Desc`, `Comment`, `Value`, `Revision Name` |
| Manufacturer | `Manufacturer`, `Mfr`, `Make`, `Manufacturer Name` |
| MPN | `MPN`, `Manufacturer Part Number`, `Mfr Part Number`, `Part Number` |
| Alternatives | `Alternatives`, `Alternate MPN`, `Approved Alternatives`, `Substitutes` |
| Qty | `Qty`, `Quantity`, `BOM Qty`, `Qty Per Assembly` |
| Part Type | `Part Type`, `Type`, `Item Type` |
| Procurement Intent | `Procurement Intent`, `Procurement`, `Sourcing Intent`, `Buy/Build`, `Make/Buy` |
Alternatives should be separated with semicolons or pipes. Build demand is entered in the app and applied to every BOM line.
The BOM table can be filtered by MPN, part type, and procurement intent before generating a report.

## Recommendation Logic

For each BOM line, the app searches the primary MPN plus any approved alternatives. It evaluates seller offers by:

1. Matching the best price break for the required quantity.
2. Prioritizing sufficient inventory, then unknown inventory, then short inventory.
3. Selecting the lowest extended cost.

The report includes recommended supplier, next-best supplier comparison, quoted MPN, unit price, extended cost, seller stock, market availability, lead time, status, and offer link.

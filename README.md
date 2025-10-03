# create-qa-reports

Node + TypeScript script that pulls Bugs/PBIs completed within each sprint and associates matching test suites.

## Configuration

Set the following environment variables before running:

- `AZDO_ORG_URL` – Azure DevOps organization URL, e.g. `https://dev.azure.com/contoso`.
- `AZDO_PROJECT` – Project name.
- `AZDO_TEAM` – Team name used for sprint iterations.
- `AZDO_PAT` – Personal Access Token with **Work Items (Read)** and **Test Management (Read)** scopes.

## Install & Build

```bash
npm install
npm run build
```

## Run

After building, execute:

```bash
npm start
```

or run without compiling for quick checks:

```bash
npm run dev
```

The script prints a table with columns:

- `ID`
- `Assigned To`
- `Test Suite Link`

Rows are sorted by assignee. If no matching test suite exists, the link column shows `Not Found`.

## Notes

- The script gathers past, current, and future iterations for the configured team and deduplicates work items.
- Test suite lookup scans all test plans in the project; heavy projects may take a moment on first run.

# Cash Disbursement Audit Sampler

A 100% client-side static web application for generating reproducible cash-disbursement audit selections from Yardi, AppFolio, and OneSite Excel reports.

## Privacy architecture

The application has no API routes, backend, database, Azure service, App Service, or Node server after deployment.

When a user chooses an Excel file:

1. The browser reads the file into memory.
2. SheetJS parses the workbook in the browser.
3. The browser applies grouping, exclusions, annualization, and sampling.
4. SheetJS creates the multi-sheet output workbook in browser memory.
5. The browser creates a temporary object URL and triggers the Excel download.
6. The temporary URL is revoked.

The selected source report and generated workbook never leave the user's device.

## Preserved functionality

- Yardi Expense Distribution (Paid Only)
- AppFolio Check Register Detail (Enhanced)
- OneSite Check Register Detail
- Standard 10% sampling, rounded up and capped at 25
- HUD Handbook 2000.04 profiles
- Partial-year population annualization
- Carter & Company fee exclusion
- Utilities exclusion
- Mortgage and debt-service exclusion
- Tenant utility-reimbursement exclusion
- Optional dollar threshold
- Deterministic SHA-256 selection ranking
- Selection Summary, Selections, Selected Detail Lines, and Excluded Disbursements worksheets
- Uniform Selections worksheet columns across all three source systems
- Unavailable source fields remain in place and may be hidden so downstream workpapers can use one import layout
- Exception Status column

## Requirements

- Node.js 22
- npm

Node.js is required only to develop and build the static files. It is not required by the deployed website.

## Local development

```bash
npm ci
npm run dev
```

Open `http://localhost:3000`.

## Verification and static build

```bash
npm ci
npm run typecheck
npm run build
```

The static website is written to `out/`.

## Static hosting

Upload the contents of `out/` to any static host, including:

- GitHub Pages
- Cloudflare Pages
- Netlify

For GitHub Pages, use the included `.github/workflows/deploy.yml`. See `DEPLOY-TO-GITHUB-PAGES.md`.

For Cloudflare Pages or Netlify:

- Build command: `npm run build`
- Output directory: `out`
- Node version: `22`
- Environment variable: leave `PAGES_BASE_PATH` empty

## Final project structure

```text
cash-disbursement-audit-sampler/
├── .github/
│   └── workflows/
│       └── deploy.yml
├── app/
│   ├── globals.css
│   ├── layout.tsx
│   └── page.tsx
├── lib/
│   └── audit-sampler.ts
├── public/
│   ├── favicon.svg
│   ├── file.svg
│   ├── globe.svg
│   └── window.svg
├── .gitignore
├── DEPLOY-TO-GITHUB-PAGES.md
├── README.md
├── next.config.ts
├── package-lock.json
├── package.json
└── tsconfig.json
```

## Architecture changes

### Removed

- `app/api/generate/route.ts`
- The `fetch("/api/generate")` request
- Multipart form submission
- Server response headers used to return counts and filenames
- The Node.js runtime requirement after deployment
- Azure/App Service deployment materials

### Moved into the browser

All former API logic now lives in `lib/audit-sampler.ts`, including:

- Workbook parsing
- Yardi, AppFolio, and OneSite normalization
- Disbursement grouping
- Exclusion classification
- Annualization
- Standard and HUD sample sizing
- SHA-256 deterministic ranking
- Workbook creation and formatting

### Excel generation and download

`generateAuditWorkbook()` returns a browser `Blob`, output filename, and population statistics. `app/page.tsx` creates a temporary object URL for the Blob, clicks a temporary download link, removes the link, and revokes the object URL.

## Making future changes

- Interface and wording: `app/page.tsx`
- Styling: `app/globals.css`
- Parsing, sampling, exclusions, and workbook output: `lib/audit-sampler.ts`
- Static-host path configuration: `next.config.ts`
- GitHub deployment: `.github/workflows/deploy.yml`

Before publishing a methodology change, test all three source systems and compare the output against an approved workbook.

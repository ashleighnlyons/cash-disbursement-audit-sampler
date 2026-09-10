import * as XLSX from "xlsx";

export type SourceSystem = "yardi" | "appfolio";
export type AuditType = "standard" | "hud";
export type HudProfile = "low-5" | "low-10" | "high-5" | "high-10";

export type AuditSamplerOptions = {
  sourceSystem: SourceSystem;
  auditType: AuditType;
  hudProfile: HudProfile;
  minimumAmount: number;
  reportMonths: number;
};

export type AuditSamplerResult = {
  blob: Blob;
  filename: string;
  totalPopulation: number;
  excludedPopulation: number;
  eligiblePopulation: number;
  annualizedPopulation: number;
  sampleCount: number;
};

type Row = (string | number | boolean | Date | null)[];

type Detail = {
  sourceRow: number;
  accountCode: string;
  accountName: string;
  payeeCode: string;
  payeeName: string;
  payableControl: string;
  batch: string;
  property: string;
  invoiceNo: string;
  invoiceDate: unknown;
  period: unknown;
  paymentMethod: string;
  amount: number;
  checkControl: string;
  checkNo: string;
  checkDate: unknown;
  notes: string;
};

type Payment = {
  checkControl: string;
  checkNo: string;
  checkDate: unknown;
  paymentMethod: string;
  payees: string;
  payableControls: string;
  invoices: string;
  accounts: string;
  amount: number;
  lineCount: number;
  sourceRows: string;
  lines: Detail[];
  score: string;
  exclusion: string;
};

const stringValue = (value: unknown) => value == null ? "" : String(value).trim();
const uniqueValues = (items: string[]) => [...new Set(items.filter(Boolean))].join("; ");
const money = (value: number) => Math.round(value * 100) / 100;

export async function generateAuditWorkbook(
  file: File,
  options: AuditSamplerOptions,
): Promise<AuditSamplerResult> {
  if (!file.name.toLowerCase().endsWith(".xlsx")) {
    throw new Error("Upload a valid .xlsx report.");
  }
  if (file.size > 20 * 1024 * 1024) {
    throw new Error("The report exceeds the 20 MB limit.");
  }

  const sourceSystem = options.sourceSystem;
  const auditType = options.auditType;
  const hudProfile = options.hudProfile;
  const minimumAmount = Math.max(0, Number(options.minimumAmount) || 0);
  const reportMonths = Math.min(12, Math.max(1, Math.round(Number(options.reportMonths) || 12)));

  let rows: Row[];
  try {
    const source = XLSX.read(await file.arrayBuffer(), { type: "array", cellDates: true });
    const sheet = source.Sheets[source.SheetNames[0]];
    rows = XLSX.utils.sheet_to_json<Row>(sheet, { header: 1, raw: true, defval: null });
  } catch {
    throw new Error("The Excel workbook could not be read. Confirm it is a valid .xlsx file and try again.");
  }

  let details: Detail[];
  try {
    details = sourceSystem === "appfolio" ? parseAppFolio(rows) : parseYardi(rows);
  } catch {
    throw new Error("The workbook could not be processed. Confirm the selected accounting system matches the unmodified report and try again.");
  }

  if (!details.length) {
    throw new Error("No paid disbursement detail rows were found.");
  }

  const grouped = new Map<string, Detail[]>();
  for (const detail of details) {
    if (!grouped.has(detail.checkControl)) grouped.set(detail.checkControl, []);
    grouped.get(detail.checkControl)!.push(detail);
  }

  const population: Payment[] = [];
  for (const [checkControl, lines] of grouped) {
    const amount = money(lines.reduce((total, line) => total + line.amount, 0));
    const categories = new Set(lines.map(exclusionFor).filter(Boolean));
    if (minimumAmount > 0 && amount < minimumAmount) {
      categories.add(`Below $${minimumAmount.toFixed(2)} threshold`);
    }
    population.push({
      checkControl,
      checkNo: lines[0].checkNo,
      checkDate: lines[0].checkDate,
      paymentMethod: lines[0].paymentMethod,
      payees: uniqueValues(lines.map((line) => line.payeeName)),
      payableControls: uniqueValues(lines.map((line) => line.payableControl)),
      invoices: uniqueValues(lines.map((line) => line.invoiceNo)),
      accounts: uniqueValues(lines.map((line) => `${line.accountCode} ${line.accountName}`)),
      amount,
      lineCount: lines.length,
      sourceRows: lines.map((line) => line.sourceRow).join(", "),
      lines,
      score: await hash(`${file.name}|${checkControl}`),
      exclusion: [...categories].join("; "),
    });
  }

  const excluded = population.filter((payment) => payment.exclusion);
  const eligible = population.filter((payment) => !payment.exclusion);
  const annualizedPopulation = Math.ceil(eligible.length * 12 / reportMonths);
  const calculatedSample = auditType === "hud"
    ? hudSampleSize(annualizedPopulation, hudProfile)
    : Math.min(Math.ceil(annualizedPopulation * 0.10), 25);
  const sampleCount = Math.min(calculatedSample, eligible.length);
  const method = auditType === "hud"
    ? "HUD Handbook 2000.04 REV-2, Appendix A"
    : "Standard 10%, rounded up, capped at 25";
  const profile = auditType === "hud" ? hudProfileLabel(hudProfile) : "Not applicable";

  const selected = [...eligible]
    .sort((left, right) => left.score.localeCompare(right.score))
    .slice(0, sampleCount)
    .sort((left, right) => dateValue(left.checkDate) - dateValue(right.checkDate)
      || left.checkControl.localeCompare(right.checkControl));

  const workbook = XLSX.utils.book_new();
  addSummarySheet(workbook, [
    ["Cash Disbursement Audit Selections"],
    [sourceSystem === "appfolio" ? "AppFolio · Check Register Detail (Enhanced)" : "Yardi · Expense Distribution (Paid Only)"],
    [],
    ["Population Reconciliation", "Count", "Amount"],
    ["Total unique disbursements", population.length, money(sum(population))],
    ["Excluded disbursements", excluded.length, money(sum(excluded))],
    ["Actual eligible population", eligible.length, money(sum(eligible))],
    ["Annualized eligible population", annualizedPopulation],
    ["Final selections", sampleCount],
    [],
    ["Methodology"],
    ["Source system", sourceSystem === "appfolio" ? "AppFolio" : "Yardi"],
    ["Client type", auditType === "hud" ? "HUD client" : "Standard client"],
    ["Sample-size rule", method],
    ["HUD sampling profile", profile],
    ["Months included in report", reportMonths],
    ["Annualization factor", 12 / reportMonths],
    ["Annualization formula", `Actual eligible population × 12 ÷ ${reportMonths}`],
    ["Sampling unit", sourceSystem === "appfolio" ? "One AppFolio payment header" : "Unique Check Control"],
    ["Selection method", "Deterministic hash-based random selection"],
    ["Minimum amount threshold", minimumAmount > 0 ? minimumAmount : "Not applied"],
    ["Exclusions", "Carter & Company fees; utilities; mortgage and debt service; tenant utility reimbursements"],
    ["HUD source", "HUD Handbook 2000.04 REV-2, Appendix A - Attribute Sampling"],
    ["Source file", file.name],
    ["Generated", new Date()],
  ]);

  const selectionHeaders = sourceSystem === "yardi"
    ? ["Selection #", "Check Control", "Check #", "Check Date", "Payee(s)", "Disbursement Amount", "Payable Control(s)", "Invoice #(s)", "Account(s)", "Exception Status"]
    : ["Selection #", "Check Control", "Check #", "Check Date", "Payee(s)", "Disbursement Amount", "Invoice #(s)", "Account(s)", "Exception Status"];
  const selectionRows: Row[] = sourceSystem === "yardi"
    ? selected.map((payment, index) => [index + 1, payment.checkControl, payment.checkNo, payment.checkDate, payment.payees, payment.amount, payment.payableControls, payment.invoices, payment.accounts, ""])
    : selected.map((payment, index) => [index + 1, payment.checkControl, payment.checkNo, payment.checkDate, payment.payees, payment.amount, payment.invoices, payment.accounts, ""]);
  addSheet(
    workbook,
    "Selections",
    [selectionHeaders, ...selectionRows],
    sourceSystem === "yardi" ? [12, 15, 12, 14, 28, 18, 32, 32, 42, 20] : [12, 15, 12, 14, 28, 18, 32, 42, 20],
  );

  const detailHeaders = ["Selection #", "Source Row", "Account Code", "Account Name", "Payee Code", "Payee Name", "Payable Control", "Batch", "Property", "Invoice #", "Invoice Date", "Period", "Payment Method", "Amount", "Check Control", "Check #", "Check Date", "Notes"];
  const detailRows: Row[] = [];
  selected.forEach((payment, index) => payment.lines.forEach((detail) => detailRows.push([
    index + 1,
    detail.sourceRow,
    detail.accountCode,
    detail.accountName,
    detail.payeeCode,
    detail.payeeName,
    detail.payableControl,
    detail.batch,
    detail.property,
    detail.invoiceNo,
    detail.invoiceDate,
    detail.period,
    detail.paymentMethod,
    detail.amount,
    detail.checkControl,
    detail.checkNo,
    detail.checkDate,
    detail.notes,
  ])));
  addSheet(workbook, "Selected Detail Lines", [detailHeaders, ...detailRows], [12, 11, 14, 28, 15, 28, 16, 11, 11, 22, 14, 12, 15, 16, 15, 12, 14, 42]);

  const excludedHeaders = ["Check Control", "Check #", "Check Date", "Payment Method", "Payee(s)", "Disbursement Amount", "Expense Lines", "Exclusion Category", "Account(s)", "Source Row(s)"];
  const excludedRows: Row[] = excluded
    .sort((left, right) => dateValue(left.checkDate) - dateValue(right.checkDate))
    .map((payment) => [payment.checkControl, payment.checkNo, payment.checkDate, payment.paymentMethod, payment.payees, payment.amount, payment.lineCount, payment.exclusion, payment.accounts, payment.sourceRows]);
  addSheet(workbook, "Excluded Disbursements", [excludedHeaders, ...excludedRows], [15, 12, 14, 15, 28, 18, 13, 28, 44, 24]);

  const clientName = sourceSystem === "appfolio"
    ? appFolioClientName(rows)
    : stringValue(rows[1]?.[0]).replace(/[^a-z0-9_-]+/gi, "_").replace(/^_+|_+$/g, "");
  const filename = `Cash_Disbursement_Audit_Selections${clientName ? `_${clientName}` : ""}.xlsx`;
  const output = XLSX.write(workbook, { type: "array", bookType: "xlsx", compression: true }) as ArrayBuffer;
  const blob = new Blob([output], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });

  return {
    blob,
    filename,
    totalPopulation: population.length,
    excludedPopulation: excluded.length,
    eligiblePopulation: eligible.length,
    annualizedPopulation,
    sampleCount,
  };
}

function parseYardi(rows: Row[]): Detail[] {
  const headerIndex = rows.findIndex((row) => stringValue(row[0]) === "Account Code" && stringValue(row[12]) === "Check Control");
  if (headerIndex < 0) throw new Error("Not a Yardi Expense Distribution report");

  const details: Detail[] = [];
  let accountCode = "";
  let accountName = "";
  for (let index = headerIndex + 1; index < rows.length; index += 1) {
    const row = rows[index];
    if (row[0] && !stringValue(row[0]).startsWith("Total") && !row[4]) {
      accountCode = stringValue(row[0]);
      accountName = stringValue(row[1]);
    }
    if (row[4] && row[12]) {
      details.push({
        sourceRow: index + 1,
        accountCode,
        accountName,
        payeeCode: stringValue(row[2]),
        payeeName: stringValue(row[3]),
        payableControl: stringValue(row[4]),
        batch: stringValue(row[5]),
        property: stringValue(row[6]),
        invoiceNo: stringValue(row[7]),
        invoiceDate: row[8],
        period: row[9],
        paymentMethod: stringValue(row[10]),
        amount: Number(row[11] || 0),
        checkControl: stringValue(row[12]),
        checkNo: stringValue(row[13]),
        checkDate: row[14],
        notes: stringValue(row[15]),
      });
    }
  }
  return details;
}

function parseAppFolio(rows: Row[]): Detail[] {
  const headerIndex = rows.findIndex((row) => stringValue(row[0]) === "Bank Account"
    && stringValue(row[1]) === "Payee Name"
    && stringValue(row[2]) === "Check #"
    && stringValue(row[6]) === "Payment Amount"
    && stringValue(row[12]) === "Bill Reference #");
  if (headerIndex < 0) throw new Error("Not an AppFolio Check Register Detail report");

  const details: Detail[] = [];
  for (let index = headerIndex + 1; index < rows.length;) {
    const payment = rows[index];
    if (!payment[0]
      || stringValue(payment[0]).startsWith("Total")
      || !Number.isFinite(Number(payment[6]))
      || Number(payment[6]) === 0) {
      index += 1;
      continue;
    }

    const sourceRow = index + 1;
    const checkNo = stringValue(payment[2]);
    const checkControl = `AF-${sourceRow}`;
    const paymentMethod = /ach/i.test(checkNo) ? "ACH" : "Check";
    const payeeName = stringValue(payment[1]);
    const checkDate = payment[5];
    const checkMemo = stringValue(payment[4]);
    const bank = stringValue(payment[0]);
    let detailIndex = index + 1;
    let lineCount = 0;

    for (; detailIndex < rows.length && !rows[detailIndex][0]; detailIndex += 1) {
      const row = rows[detailIndex];
      if (!row[7] && !row[8] && !row[9] && !row[10] && !row[11] && !row[12]) continue;
      const billReference = stringValue(row[12]);
      details.push({
        sourceRow: detailIndex + 1,
        accountCode: stringValue(row[8]),
        accountName: stringValue(row[9]),
        payeeCode: "",
        payeeName,
        payableControl: billReference,
        batch: "",
        property: stringValue(row[7]),
        invoiceNo: billReference,
        invoiceDate: null,
        period: null,
        paymentMethod,
        amount: Number(row[10] || 0),
        checkControl,
        checkNo,
        checkDate,
        notes: uniqueValues([stringValue(row[11]), checkMemo, bank]),
      });
      lineCount += 1;
    }

    if (!lineCount) {
      details.push({
        sourceRow,
        accountCode: "",
        accountName: "",
        payeeCode: "",
        payeeName,
        payableControl: "",
        batch: "",
        property: "",
        invoiceNo: "",
        invoiceDate: null,
        period: null,
        paymentMethod,
        amount: Number(payment[6]),
        checkControl,
        checkNo,
        checkDate,
        notes: uniqueValues([checkMemo, bank]),
      });
    }
    index = detailIndex;
  }
  return details;
}

function appFolioClientName(rows: Row[]): string {
  const rawName = stringValue(rows.find((row) => stringValue(row[0]).startsWith("Properties:"))?.[0])
    .replace(/^Properties:\s*/, "")
    .split(" - ")[0];
  return rawName.replace(/[^a-z0-9_-]+/gi, "_").replace(/^_+|_+$/g, "");
}

function exclusionFor(detail: Detail): string {
  const payee = detail.payeeName.toLowerCase().replace(/[^a-z0-9]+/g, " ");
  const account = `${detail.accountCode} ${detail.accountName}`.toLowerCase();
  const context = `${detail.payeeName} ${detail.invoiceNo} ${detail.accountName} ${detail.notes}`.toLowerCase();
  if (/\bcarter\b/.test(payee) && /\b(company|co|cpas?|accounting)\b/.test(payee)) return "Carter & Company fees";
  if (/\btenant\b.*\butility\b.*\breimb|\butility\b.*\breimb|\bresident\b.*\butility\b.*\breimb/.test(context)) return "Tenant utility reimbursement";
  if (/\b(electric|electricity|gas|water|sewer|utility|utilities)\b/.test(account)) return "Utilities";
  if (/\b(mortgage|loan note|debt service|interest payable mortgage|tax escrow|replacement reserve held by others)\b/.test(account)) return "Mortgage & debt service";
  return "";
}

function hudSampleSize(population: number, profile: HudProfile): number {
  if (population <= 0) return 0;
  if (population < 20) return Math.min(population, 5);
  if (population < 50) return 5;
  if (population < 100) return 10;
  if (population <= 200) return 20;
  return { "low-5": 50, "low-10": 25, "high-5": 65, "high-10": 35 }[profile] || 65;
}

function hudProfileLabel(profile: HudProfile): string {
  return {
    "low-5": "Low importance; 90% confidence; 5% tolerable exception rate",
    "low-10": "Low importance; 90% confidence; 10% tolerable exception rate",
    "high-5": "High importance; 95% confidence; 5% tolerable exception rate",
    "high-10": "High importance; 95% confidence; 10% tolerable exception rate",
  }[profile] || "High importance; 95% confidence; 5% tolerable exception rate";
}

async function hash(value: string): Promise<string> {
  const bytes = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(value));
  return [...new Uint8Array(bytes)].map((byte) => byte.toString(16).padStart(2, "0")).join("");
}

function dateValue(value: unknown): number {
  if (value instanceof Date) return value.getTime();
  if (typeof value === "number") return value;
  return Date.parse(stringValue(value)) || 0;
}

function sum(items: Payment[]): number {
  return items.reduce((total, payment) => total + payment.amount, 0);
}

function addSheet(workbook: XLSX.WorkBook, name: string, data: Row[], widths: number[]): void {
  const worksheet = XLSX.utils.aoa_to_sheet(data, { cellDates: true });
  worksheet["!cols"] = widths.map((width) => ({ wch: width }));
  if (data.length > 1) {
    worksheet["!autofilter"] = { ref: `A1:${XLSX.utils.encode_col(Math.max(0, widths.length - 1))}${data.length}` };
  }
  XLSX.utils.book_append_sheet(workbook, worksheet, name);
}

function addSummarySheet(workbook: XLSX.WorkBook, data: Row[]): void {
  const worksheet = XLSX.utils.aoa_to_sheet(data, { cellDates: true });
  worksheet["!cols"] = [{ wch: 39 }, { wch: 58 }, { wch: 19 }];
  worksheet["!rows"] = [{ hpt: 27 }, { hpt: 20 }, { hpt: 8 }, { hpt: 22 }, { hpt: 20 }, { hpt: 20 }, { hpt: 20 }, { hpt: 20 }, { hpt: 20 }, { hpt: 8 }, { hpt: 22 }];
  worksheet["!merges"] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: 2 } },
    { s: { r: 1, c: 0 }, e: { r: 1, c: 2 } },
    { s: { r: 10, c: 0 }, e: { r: 10, c: 2 } },
  ];
  worksheet["!sheetViews"] = [{ showGridLines: false }];

  const title = worksheet.A1;
  const subtitle = worksheet.A2;
  const reconciliation = worksheet.A4;
  const methodology = worksheet.A11;
  if (title) title.s = { font: { bold: true, sz: 18, color: { rgb: "FFFFFF" } }, fill: { fgColor: { rgb: "17365D" } }, alignment: { horizontal: "left", vertical: "center" } };
  if (subtitle) subtitle.s = { font: { italic: true, sz: 11, color: { rgb: "44546A" } }, alignment: { horizontal: "left" } };
  for (const cell of [reconciliation, worksheet.B4, worksheet.C4, methodology]) {
    if (cell) cell.s = { font: { bold: true, color: { rgb: "FFFFFF" } }, fill: { fgColor: { rgb: "2F75B5" } }, alignment: { vertical: "center" } };
  }
  for (let row = 4; row <= 8; row += 1) {
    const count = worksheet[`B${row + 1}`];
    const amount = worksheet[`C${row + 1}`];
    if (count) count.z = "#,##0";
    if (amount) amount.z = '"$"#,##0.00;[Red]("$"#,##0.00)';
  }
  if (worksheet.B17) worksheet.B17.z = "0.00x";
  if (worksheet.B25) worksheet.B25.z = "mmm d, yyyy h:mm AM/PM";
  for (let row = 12; row <= 25; row += 1) {
    const label = worksheet[`A${row}`];
    if (label) label.s = { font: { bold: true, color: { rgb: "44546A" } }, fill: { fgColor: { rgb: row % 2 ? "F7F9FC" : "EAF2F8" } };
  }
  XLSX.utils.book_append_sheet(workbook, worksheet, "Selection Summary");
}

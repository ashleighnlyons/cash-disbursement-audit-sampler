"use client";

import { useRef, useState } from "react";
import { generateAuditWorkbook, type AuditType, type HudProfile, type SourceSystem } from "@/lib/audit-sampler";

type Stage = "idle" | "ready" | "processing" | "done" | "error";

const sourceDetails: Record<SourceSystem, { name: string; report: string; icon: string; unit: string; grouping: string }> = {
  yardi: { name: "Yardi", report: "Expense Distribution (Paid Only)", icon: "Y", unit: "Yardi Check Control", grouping: "Expense lines are grouped into one payment" },
  appfolio: { name: "AppFolio", report: "Check Register Detail (Enhanced)", icon: "A", unit: "AppFolio payment", grouping: "Payment header and its GL lines stay together" },
  onesite: { name: "OneSite", report: "Check Register Detail", icon: "O", unit: "OneSite payment", grouping: "Payment header and its applied lines stay together" },
};

export default function Home() {
  const inputRef = useRef<HTMLInputElement>(null);
  const [file, setFile] = useState<File | null>(null);
  const [stage, setStage] = useState<Stage>("idle");
  const [message, setMessage] = useState("");
  const [sourceSystem, setSourceSystem] = useState<SourceSystem>("yardi");
  const [auditType, setAuditType] = useState<AuditType>("standard");
  const [hudProfile, setHudProfile] = useState<HudProfile>("high-5");
  const [useThreshold, setUseThreshold] = useState(false);
  const [threshold, setThreshold] = useState("");
  const [reportMonths, setReportMonths] = useState("12");
  const selectedSource = sourceDetails[sourceSystem];

  function choose(next: File | null) {
    if (!next) return;
    if (!next.name.toLowerCase().endsWith(".xlsx")) {
      setFile(null); setStage("error"); setMessage("Please upload an Excel .xlsx report."); return;
    }
    if (next.size > 20 * 1024 * 1024) {
      setFile(null); setStage("error"); setMessage("The report exceeds the 20 MB limit."); return;
    }
    setFile(next); setStage("ready"); setMessage("");
  }

  function selectSourceSystem(next: SourceSystem) {
    setSourceSystem(next); setFile(null); setStage("idle"); setMessage("");
    if (inputRef.current) inputRef.current.value = "";
  }

  async function generate() {
    if (!file) return;
    setStage("processing"); setMessage("");
    try {
      const result = await generateAuditWorkbook(file, {
        sourceSystem,
        auditType,
        hudProfile,
        minimumAmount: useThreshold ? Number(threshold || 0) : 0,
        reportMonths: Number(reportMonths),
      });
      const url = URL.createObjectURL(result.blob);
      const anchor = document.createElement("a");
      anchor.href = url; anchor.download = result.filename;
      document.body.appendChild(anchor); anchor.click(); anchor.remove();
      window.setTimeout(() => URL.revokeObjectURL(url), 1_000);
      const annualized = result.annualizedPopulation !== result.eligiblePopulation ? `, ${result.annualizedPopulation} annualized` : "";
      setStage("done");
      setMessage(`Complete: ${result.totalPopulation} total, ${result.excludedPopulation} excluded, ${result.eligiblePopulation} eligible${annualized}, ${result.sampleCount} selected.`);
    } catch (error) {
      setStage("error");
      setMessage(error instanceof Error ? error.message : "Unable to process this report.");
    }
  }

  return <main>
    <header className="topbar"><div className="brand"><span className="brandMark">CC</span><span>Carter &amp; Company</span></div><div className="secure"><span className="dot" /> Browser-only audit tool</div></header>
    <section className="hero"><div className="eyebrow">AUDIT OPERATIONS</div><h1>Cash Disbursement<br/><em>Audit Sampler</em></h1><p>Turn a cash-disbursement report into a documented, reproducible audit sample without uploading it anywhere.</p></section>
    <section className="workspace">
      <div className="steps" aria-label="Workflow steps"><div className="step active"><span>1</span><b>Choose file</b></div><div className="line"/><div className={`step ${stage === "processing" || stage === "done" ? "active" : ""}`}><span>2</span><b>Process in browser</b></div><div className="line"/><div className={`step ${stage === "done" ? "active" : ""}`}><span>3</span><b>Download workbook</b></div></div>
      <div className="panelGrid">
        <section className="uploadCard">
          <div className="cardHeading"><span className="number">01</span><div><h2>Choose the accounting system</h2><p>Select the system that produced the report.</p></div></div>
          <div className="methodPicker systemPicker" role="radiogroup" aria-label="Accounting system">
            <button type="button" className={sourceSystem === "yardi" ? "selected" : ""} onClick={() => selectSourceSystem("yardi")}><b>Yardi</b><small>Expense Distribution (Paid Only)</small></button>
            <button type="button" className={sourceSystem === "appfolio" ? "selected" : ""} onClick={() => selectSourceSystem("appfolio")}><b>AppFolio</b><small>Check Register Detail (Enhanced)</small></button>
            <button type="button" className={sourceSystem === "onesite" ? "selected" : ""} onClick={() => selectSourceSystem("onesite")}><b>OneSite</b><small>Check Register Detail</small></button>
          </div>
          <div className="sectionLabel">Choose the audit method</div>
          <div className="methodPicker" role="radiogroup" aria-label="Audit method">
            <button type="button" className={auditType === "standard" ? "selected" : ""} onClick={() => setAuditType("standard")}><b>Standard client</b><small>10%, rounded up · maximum 25</small></button>
            <button type="button" className={auditType === "hud" ? "selected" : ""} onClick={() => setAuditType("hud")}><b>HUD client</b><small>Handbook 2000.04 · Appendix A</small></button>
          </div>
          {auditType === "hud" && <label className="field"><span>HUD sampling profile</span><select value={hudProfile} onChange={(event) => setHudProfile(event.target.value as HudProfile)}><option value="high-5">High importance · 95% confidence · 5% tolerable rate</option><option value="high-10">High importance · 95% confidence · 10% tolerable rate</option><option value="low-5">Low importance · 90% confidence · 5% tolerable rate</option><option value="low-10">Low importance · 90% confidence · 10% tolerable rate</option></select><small>Used for eligible populations over 200; HUD small-population minimums apply otherwise.</small></label>}
          <label className="field"><span>Months included in report</span><select value={reportMonths} onChange={(event) => setReportMonths(event.target.value)}>{Array.from({ length: 12 }, (_, index) => index + 1).map((months) => <option key={months} value={months}>{months} month{months === 1 ? "" : "s"}{months === 12 ? " · Full year" : " · Annualize population"}</option>)}</select><small>{reportMonths === "12" ? "Full-year population; no adjustment needed." : `Population will be annualized using 12 ÷ ${reportMonths}.`}</small></label>
          <label className="threshold"><input type="checkbox" checked={useThreshold} onChange={(event) => setUseThreshold(event.target.checked)}/><span>Exclude disbursements under</span><div className="moneyInput"><b>$</b><input aria-label="Minimum disbursement amount" type="number" min="0" step="0.01" placeholder="0.00" value={threshold} disabled={!useThreshold} onChange={(event) => setThreshold(event.target.value)}/></div></label>
          <div className="sectionLabel">Choose {selectedSource.name} report</div>
          <button type="button" className={`dropzone ${file ? "hasFile" : ""}`} onClick={() => inputRef.current?.click()} onDragOver={(event) => event.preventDefault()} onDrop={(event) => { event.preventDefault(); choose(event.dataTransfer.files[0]); }}><input ref={inputRef} type="file" accept=".xlsx" hidden onChange={(event) => choose(event.target.files?.[0] || null)}/><span className="uploadIcon">↥</span>{file ? <><strong>{file.name}</strong><small>{(file.size / 1024).toFixed(0)} KB · Ready to process</small></> : <><strong>Drop your {selectedSource.name} report here</strong><small>{selectedSource.report} · .xlsx only</small></>}</button>
          <button type="button" className="generate" disabled={!file || stage === "processing"} onClick={generate}>{stage === "processing" ? "Processing in browser…" : "Generate audit sample"}<span>→</span></button>
          {message && <div className={`notice ${stage}`} role="status">{message}</div>}
          <p className="privacy"><span>⌁</span> Your Excel file is processed entirely in this browser and never leaves this device.</p>
        </section>
        <aside className="rulesCard">
          <div className="cardHeading compact"><span className="number">02</span><div><h2>Rules applied automatically</h2><p>Locked methodology for consistent testing.</p></div></div>
          <div className="rule"><span className="ruleIcon">{selectedSource.icon}</span><div><b>{selectedSource.unit}</b><small>{selectedSource.grouping}</small></div></div>
          <div className="rule"><span className="ruleIcon">%</span><div><b>{auditType === "hud" ? "HUD minimum sample" : "10% of eligible disbursements"}</b><small>{auditType === "hud" ? "Handbook 2000.04 Appendix A" : "Rounded up and capped at 25 selections"}</small></div></div>
          <div className="rule"><span className="ruleIcon">#</span><div><b>One disbursement per selection</b><small>Related expense lines remain together</small></div></div>
          <div className="rule"><span className="ruleIcon">12</span><div><b>{reportMonths === "12" ? "Full-year population" : "Annualized population"}</b><small>{reportMonths === "12" ? "No period adjustment" : `Actual eligible count × 12 ÷ ${reportMonths}`}</small></div></div>
          <div className="rule"><span className="ruleIcon">✓</span><div><b>Reproducible selection</b><small>Deterministic hash-based random sample</small></div></div>
          <div className="exclusions"><h3>Excluded before sampling</h3><span>Carter &amp; Company fees</span><span>Utilities</span><span>Mortgage &amp; debt service</span><span>Tenant utility reimbursements</span>{useThreshold && <span>Disbursements below ${Number(threshold || 0).toLocaleString(undefined, { minimumFractionDigits: 2 })}</span>}</div>
        </aside>
      </div>
    </section>
    <footer><span>Cash Disbursement Audit Sampler</span><span>Local browser processing · No file upload</span></footer>
  </main>;
}

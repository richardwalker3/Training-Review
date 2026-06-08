import { useState, useRef } from "react";
import * as XLSX from "xlsx";

const emptyLO = {
  loName: "",
  yearsInIndustry: "",
  encompassExperience: "No",
  attendance: "",
  notableStrengths: "",
  notableWeaknesses: "",
  uniqueFact: "",
  nafDetailsQuiz: "",
  nafLinkQuiz: "",
  encompassQuiz: "",
  finalExam: "",
};

const MULTI_LO_PROMPT = `
These are training grades/survey files for a group of Loan Officers (LOs). There may be multiple files combined below — merge data by LO name.

Extract EVERY LO found. For each LO return a block using EXACTLY this XML format:

<lo>
<loName>full name</loName>
<yearsInIndustry>e.g. 1.5 yrs</yearsInIndustry>
<encompassExperience>Yes or No</encompassExperience>
<attendance>e.g. 100%</attendance>
<uniqueFact>fun fact or leave blank</uniqueFact>
<nafDetailsQuiz>list EVERY attempt score, comma separated e.g. 60%, 90%</nafDetailsQuiz>
<nafLinkQuiz>list EVERY attempt score, comma separated or blank</nafLinkQuiz>
<encompassQuiz>list EVERY attempt score, comma separated or blank</encompassQuiz>
<finalExam>list EVERY attempt score, comma separated or blank</finalExam>
</lo>

IMPORTANT: For quiz/exam scores, include ALL attempts in order — do not only show the highest or most recent score. Separate each attempt with a comma.
Output ONE <lo>...</lo> block per LO. Use blank (empty) tags for missing fields.
Output ONLY the <lo> blocks, nothing else.`;

// ── helpers ───────────────────────────────────────────────────────────────────
function toBase64(file) {
  return new Promise((res, rej) => {
    const r = new FileReader();
    r.onload = () => res(r.result.split(",")[1]);
    r.onerror = () => rej(new Error("Failed to read file"));
    r.readAsDataURL(file);
  });
}

function extractTag(text, tag) {
  const m = text.match(new RegExp(`<${tag}>(.*?)</${tag}>`, "s"));
  return m ? m[1].trim() : "";
}

function parseAllLOs(text) {
  const blocks = [...text.matchAll(/<lo>(.*?)<\/lo>/gs)].map(m => m[1]);
  return blocks.map(block => ({
    loName:              extractTag(block, "loName"),
    yearsInIndustry:     extractTag(block, "yearsInIndustry"),
    encompassExperience: extractTag(block, "encompassExperience") || "No",
    attendance:          extractTag(block, "attendance"),
    uniqueFact:          extractTag(block, "uniqueFact"),
    nafDetailsQuiz:      extractTag(block, "nafDetailsQuiz"),
    nafLinkQuiz:         extractTag(block, "nafLinkQuiz"),
    encompassQuiz:       extractTag(block, "encompassQuiz"),
    finalExam:           extractTag(block, "finalExam"),
    notableStrengths:    "",
    notableWeaknesses:   "",
  })).filter(lo => lo.loName);
}

async function callClaude(messages) {
  const response = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ model: "claude-sonnet-4-20250514", max_tokens: 4000, messages }),
  });
  const result = await response.json();
  if (result.error) throw new Error(result.error.message);
  return result.content.map(b => b.text || "").join("");
}

// Convert a single file to a text chunk for the AI
async function fileToTextChunk(file) {
  const ext = file.name.split(".").pop().toLowerCase();
  const isExcel = ["xlsx","xls","xlsm","ods"].includes(ext);
  const isImage = file.type.startsWith("image/");
  const isPdf   = file.type === "application/pdf";

  if (isExcel) {
    const ab = await file.arrayBuffer();
    const wb = XLSX.read(ab, { type: "array" });
    const csv = wb.SheetNames.map(name => {
      const sheet = wb.Sheets[name];
      return `=== Sheet: ${name} ===\n${XLSX.utils.sheet_to_csv(sheet, { blankrows: false })}`;
    }).join("\n\n");
    return { type: "text", text: `=== FILE: ${file.name} ===\n${csv}` };
  } else if (isImage) {
    const b64 = await toBase64(file);
    return [
      { type: "text", text: `=== FILE: ${file.name} (image) ===` },
      { type: "image", source: { type: "base64", media_type: file.type, data: b64 } },
    ];
  } else if (isPdf) {
    const b64 = await toBase64(file);
    return [
      { type: "text", text: `=== FILE: ${file.name} (PDF) ===` },
      { type: "document", source: { type: "base64", media_type: "application/pdf", data: b64 } },
    ];
  } else {
    const text = await file.text();
    return { type: "text", text: `=== FILE: ${file.name} ===\n${text}` };
  }
}

// ── main component ────────────────────────────────────────────────────────────
export default function TrainingReviewTemplate() {
  const [roster, setRoster]         = useState([]);
  const [selected, setSelected]     = useState([]);
  const [currentIdx, setCurrentIdx] = useState(0);
  const [view, setView]             = useState("upload");
  const [uploadedFiles, setUploadedFiles] = useState([]); // names of files loaded

  const [data, setData]             = useState({ ...emptyLO });
  const [copied, setCopied]         = useState(false);
  const [parsing, setParsing]       = useState(false);
  const [parseMsg, setParseMsg]     = useState("");
  const [parseError, setParseError] = useState("");
  const tableRef    = useRef(null);
  const fileInputRef = useRef(null);

  const update = (field, value) => setData(d => ({ ...d, [field]: value }));

  const saveCurrentToRoster = (currentRoster, selectedIdx, currentData) => {
    const rosterIdx = selected[selectedIdx];
    const updated = [...currentRoster];
    updated[rosterIdx] = { ...currentData };
    return updated;
  };

  const goTo = (newIdx) => {
    const updatedRoster = saveCurrentToRoster(roster, currentIdx, data);
    setRoster(updatedRoster);
    setCurrentIdx(newIdx);
    setData({ ...updatedRoster[selected[newIdx]] });
    setCopied(false);
  };

  // ── multi-file upload ──
  const handleFileUpload = async (e) => {
    const files = Array.from(e.target.files);
    if (!files.length) return;

    setParsing(true);
    setParseError("");
    setParseMsg(`Reading ${files.length} file${files.length > 1 ? "s" : ""}…`);

    try {
      // Convert all files to content blocks
      const contentBlocks = [];
      for (let i = 0; i < files.length; i++) {
        setParseMsg(`Reading file ${i + 1} of ${files.length}: ${files[i].name}…`);
        const chunk = await fileToTextChunk(files[i]);
        if (Array.isArray(chunk)) contentBlocks.push(...chunk);
        else contentBlocks.push(chunk);
      }

      // Add the instruction prompt at the end
      contentBlocks.push({ type: "text", text: MULTI_LO_PROMPT });

      setParseMsg(`Parsing ${files.length > 1 ? "all files" : "file"} with AI…`);
      const rawText = await callClaude([{ role: "user", content: contentBlocks }]);

      const los = parseAllLOs(rawText);
      if (!los.length) throw new Error("No LOs found in the uploaded file(s).");

      setRoster(los);
      setSelected([]);
      setUploadedFiles(files.map(f => f.name));
      setView("select");
      setParseMsg("");
    } catch (err) {
      console.error(err);
      setParseError(`Could not parse: ${err.message}`);
      setParseMsg("");
    } finally {
      setParsing(false);
      if (fileInputRef.current) fileInputRef.current.value = "";
    }
  };

  const confirmSelection = () => {
    if (!selected.length) return;
    setData({ ...roster[selected[0]] });
    setCurrentIdx(0);
    setView("review");
  };

  const toggleSelect = (idx) => {
    setSelected(prev =>
      prev.includes(idx) ? prev.filter(i => i !== idx) : [...prev, idx].sort((a,b)=>a-b)
    );
  };

  const copyTable = () => {
    const el = tableRef.current;
    if (!el) return;
    const range = document.createRange();
    range.selectNode(el);
    const sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(range);
    document.execCommand("copy");
    sel.removeAllRanges();
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  const selectedLOs  = selected.map(i => roster[i]);
  const totalSelected = selectedLOs.length;

  return (
    <div style={{ fontFamily:"Calibri, Arial, sans-serif", padding:24, background:"#f0f4fa", minHeight:"100vh" }}>
      <div style={{ maxWidth:860, margin:"0 auto" }}>

        <h1 style={{ fontSize:20, fontWeight:700, color:"#1e3a5f", marginBottom:4 }}>
          Week 1 Training Review Generator
        </h1>
        <p style={{ fontSize:13, color:"#555", marginBottom:16 }}>
          Upload your grades and survey files together, select the LOs you trained, then fill in Strengths &amp; Weaknesses for each.
        </p>

        {/* ── STEP 1: UPLOAD ── */}
        <div style={{
          background:"#fff", border:`2px dashed ${view==="upload"?"#4472c4":"#bcd0ea"}`,
          borderRadius:8, padding:"16px 20px", marginBottom:20,
          display:"flex", alignItems:"flex-start", gap:16, flexWrap:"wrap",
        }}>
          <div style={{ flex:1, minWidth:200 }}>
            <div style={{ fontWeight:700, color:"#1e3a5f", fontSize:14, marginBottom:3 }}>
              📎 {view === "upload" ? "Step 1 — Upload Grades & Survey Files" : `✓ ${uploadedFiles.length} file${uploadedFiles.length>1?"s":""} loaded — ${roster.length} LOs found`}
            </div>
            {view === "upload" ? (
              <div style={{ fontSize:12, color:"#666" }}>
                Select multiple files at once (Excel, CSV, PDF, image, text). Hold <strong>Ctrl / ⌘</strong> to pick more than one.
              </div>
            ) : (
              <>
                <div style={{ fontSize:12, color:"#16a34a", fontWeight:600, marginBottom:4 }}>
                  Select the LOs you trained below, then click Confirm.
                </div>
                <div style={{ display:"flex", flexWrap:"wrap", gap:4 }}>
                  {uploadedFiles.map((name, i) => (
                    <span key={i} style={{
                      fontSize:11, background:"#eef3fb", border:"1px solid #bcd0ea",
                      borderRadius:4, padding:"2px 7px", color:"#1e3a5f",
                    }}>📄 {name}</span>
                  ))}
                </div>
              </>
            )}
          </div>
          <label style={{
            background: parsing ? "#aaa" : "#4472c4", color:"#fff",
            padding:"8px 18px", borderRadius:6, fontSize:13, fontWeight:600,
            cursor: parsing ? "not-allowed" : "pointer", whiteSpace:"nowrap", marginTop:2,
          }}>
            {parsing ? "Parsing…" : view === "upload" ? "Choose Files" : "Re-upload"}
            <input ref={fileInputRef} type="file" multiple
              accept=".csv,.pdf,.xlsx,.xls,.xlsm,.ods,.txt,.png,.jpg,.jpeg,.webp"
              style={{ display:"none" }} disabled={parsing} onChange={handleFileUpload} />
          </label>
          {parseMsg && !parseError && (
            <div style={{ width:"100%", fontSize:12, fontWeight:600, color:"#1e3a5f" }}>{parseMsg}</div>
          )}
          {parseError && (
            <div style={{ width:"100%", fontSize:12, fontWeight:600, color:"#dc2626" }}>⚠️ {parseError}</div>
          )}
        </div>

        {/* ── STEP 2: SELECT LOs ── */}
        {view === "select" && (
          <div style={{ background:"#fff", borderRadius:8, padding:20, boxShadow:"0 2px 8px rgba(0,0,0,0.08)", marginBottom:20 }}>
            <SectionHeader>Step 2 — Select the LOs You Trained</SectionHeader>
            <p style={{ fontSize:12, color:"#555", marginBottom:12 }}>
              Check each LO you personally trained this week. Uncheck any you didn't train or who didn't complete training.
            </p>
            <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill, minmax(220px,1fr))", gap:8, marginBottom:16 }}>
              {roster.map((lo, idx) => (
                <label key={idx} style={{
                  display:"flex", alignItems:"center", gap:8,
                  padding:"8px 12px", borderRadius:6, cursor:"pointer",
                  border:`2px solid ${selected.includes(idx)?"#4472c4":"#dce6f1"}`,
                  background: selected.includes(idx)?"#eef3fb":"#fafafa",
                  fontSize:13, fontWeight: selected.includes(idx)?700:400,
                  color:"#1e3a5f", transition:"all 0.15s",
                }}>
                  <input type="checkbox" checked={selected.includes(idx)}
                    onChange={() => toggleSelect(idx)}
                    style={{ accentColor:"#4472c4", width:15, height:15 }} />
                  {lo.loName || `LO #${idx+1}`}
                </label>
              ))}
            </div>
            <button onClick={confirmSelection} disabled={!selected.length} style={{
              background: selected.length ? "#1e3a5f" : "#aaa", color:"#fff",
              border:"none", borderRadius:6, padding:"9px 24px",
              fontSize:14, fontWeight:600, cursor: selected.length ? "pointer" : "not-allowed",
            }}>
              Confirm — Review {selected.length} LO{selected.length!==1?"s":""} →
            </button>
          </div>
        )}

        {/* ── STEP 3: REVIEW ── */}
        {view === "review" && (
          <>
            {/* Navigator */}
            <div style={{
              background:"#fff", borderRadius:8, padding:"12px 20px",
              marginBottom:16, boxShadow:"0 2px 8px rgba(0,0,0,0.08)",
              display:"flex", alignItems:"center", gap:12, flexWrap:"wrap",
            }}>
              <button onClick={() => goTo(currentIdx - 1)} disabled={currentIdx===0}
                style={navBtnStyle(currentIdx===0)}>← Prev</button>

              <div style={{ flex:1, textAlign:"center" }}>
                {totalSelected > 1 && <>
                  <span style={{ fontSize:13, color:"#555" }}>LO </span>
                  <span style={{ fontSize:15, fontWeight:700, color:"#1e3a5f" }}>
                    {currentIdx+1} of {totalSelected}
                  </span>
                  <span style={{ fontSize:13, color:"#888", marginLeft:8 }}>
                    — {selectedLOs[currentIdx]?.loName || ""}
                  </span>
                </>}
                {totalSelected === 1 && (
                  <span style={{ fontSize:14, fontWeight:700, color:"#1e3a5f" }}>
                    {selectedLOs[0]?.loName || ""}
                  </span>
                )}
              </div>

              <button onClick={() => goTo(currentIdx + 1)} disabled={currentIdx===totalSelected-1}
                style={navBtnStyle(currentIdx===totalSelected-1)}>Next →</button>

              <button onClick={() => setView("select")}
                style={{ ...navBtnStyle(false), background:"transparent", color:"#4472c4", border:"1px solid #4472c4" }}>
                ← List
              </button>
            </div>

            {/* Form */}
            <div style={{ background:"#fff", borderRadius:8, padding:20, boxShadow:"0 2px 8px rgba(0,0,0,0.08)", marginBottom:28 }}>
              <SectionHeader>LO Information</SectionHeader>
              <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:"12px 20px" }}>
                <Field label="LO Name" placeholder="e.g. Robert Fisher" value={data.loName} onChange={v=>update("loName",v)} />
                <Field label="Years in the Industry" placeholder="e.g. 1.5 yrs" value={data.yearsInIndustry} onChange={v=>update("yearsInIndustry",v)} />
                <div>
                  <label style={labelStyle}>Encompass Experience</label>
                  <select style={{ ...inputStyle, width:"100%" }} value={data.encompassExperience} onChange={e=>update("encompassExperience",e.target.value)}>
                    <option>No</option>
                    <option>Yes</option>
                  </select>
                </div>
                <Field label="Attendance" placeholder="e.g. 100%" value={data.attendance} onChange={v=>update("attendance",v)} />
                <div style={{ gridColumn:"1 / -1" }}>
                  <label style={{ ...labelStyle, color:"#b45309" }}>⭐ Notable Strengths <span style={{ fontWeight:400, color:"#999" }}>(fill in manually)</span></label>
                  <textarea style={{ ...inputStyle, width:"100%", minHeight:60, resize:"vertical" }}
                    placeholder="Describe strengths observed during Week 1 training…"
                    value={data.notableStrengths} onChange={e=>update("notableStrengths",e.target.value)} />
                </div>
                <div style={{ gridColumn:"1 / -1" }}>
                  <label style={{ ...labelStyle, color:"#b45309" }}>⭐ Notable Weaknesses <span style={{ fontWeight:400, color:"#999" }}>(fill in manually)</span></label>
                  <textarea style={{ ...inputStyle, width:"100%", minHeight:60, resize:"vertical" }}
                    placeholder="Describe weaknesses or areas for improvement…"
                    value={data.notableWeaknesses} onChange={e=>update("notableWeaknesses",e.target.value)} />
                </div>
                <div style={{ gridColumn:"1 / -1" }}>
                  <Field label="Unique Fact" placeholder="Fun fact shared by the LO" value={data.uniqueFact} onChange={v=>update("uniqueFact",v)} />
                </div>
              </div>

              <SectionHeader style={{ marginTop:16 }}>Quiz &amp; Exam Scores</SectionHeader>
              <p style={{ fontSize:11, color:"#888", marginTop:-6, marginBottom:10 }}>
                All attempts shown, comma separated. Edit if needed.
              </p>
              <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:"12px 20px" }}>
                <Field label="NAF Details Quiz" placeholder="e.g. 100%" value={data.nafDetailsQuiz} onChange={v=>update("nafDetailsQuiz",v)} />
                <Field label="NAF Link Quiz" placeholder="e.g. 60%, 90%" value={data.nafLinkQuiz} onChange={v=>update("nafLinkQuiz",v)} />
                <Field label="Encompass Quiz" placeholder="e.g. 70%, 80%" value={data.encompassQuiz} onChange={v=>update("encompassQuiz",v)} />
                <Field label="Final Exam" placeholder="e.g. 68%, 78%" value={data.finalExam} onChange={v=>update("finalExam",v)} />
              </div>
            </div>

            {/* Preview */}
            <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", marginBottom:10 }}>
              <span style={{ fontSize:15, fontWeight:700, color:"#1e3a5f" }}>Preview — Copy &amp; Paste into Outlook</span>
              <div style={{ display:"flex", gap:10, alignItems:"center" }}>
                {totalSelected > 1 && (
                  <span style={{ fontSize:12, color:"#888" }}>{currentIdx+1} / {totalSelected}</span>
                )}
                <button onClick={copyTable} style={{
                  background: copied ? "#22c55e" : "#1e3a5f", color:"#fff",
                  border:"none", borderRadius:6, padding:"8px 22px",
                  fontSize:14, fontWeight:600, cursor:"pointer", transition:"background 0.2s",
                }}>
                  {copied ? "✓ Copied!" : "Copy Table"}
                </button>
              </div>
            </div>

            <div ref={tableRef} style={{ display:"inline-block", width:"100%" }}>
              <table border="1" cellPadding="6" cellSpacing="0" style={{
                borderCollapse:"collapse", width:"100%",
                fontFamily:"Calibri, Arial, sans-serif", fontSize:13, border:"1px solid #4472c4",
              }}>
                <tbody>
                  <tr>
                    <td colSpan={2} style={{
                      background:"#1e3a5f", color:"#fff", fontWeight:"bold",
                      textAlign:"center", fontSize:14, padding:"7px 10px", border:"1px solid #4472c4",
                    }}>
                      {data.loName ? `${data.loName} - New Hire Training Review` : "[LO Name] - New Hire Training Review"}
                    </td>
                  </tr>
                  <Row label="Years in the Industry" value={data.yearsInIndustry} />
                  <Row label="Encompass Experience" value={data.encompassExperience} />
                  <Row label="Attendance" value={data.attendance} bold />
                  <Row label="Notable Strengths" value={data.notableStrengths} bold />
                  <Row label="Notable Weaknesses" value={data.notableWeaknesses} bold />
                  <Row label="Unique Fact" value={data.uniqueFact} />
                  <tr>
                    <td colSpan={2} style={{ background:"#dce6f1", border:"1px solid #4472c4", padding:3 }}>&nbsp;</td>
                  </tr>
                  <Row label="NAF Details Quiz" value={data.nafDetailsQuiz} lightHeader />
                  <Row label="NAF Link Quiz" value={data.nafLinkQuiz} lightHeader />
                  <Row label="Encompass Quiz" value={data.encompassQuiz} lightHeader />
                  <Row label="Final Exam" value={data.finalExam} lightHeader />
                </tbody>
              </table>
            </div>
          </>
        )}

        {/* ── MANUAL FALLBACK (no file) ── */}
        {view === "upload" && (
          <div style={{ background:"#fff", borderRadius:8, padding:20, boxShadow:"0 2px 8px rgba(0,0,0,0.08)", marginBottom:28 }}>
            <SectionHeader>Or fill in manually</SectionHeader>
            <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:"12px 20px" }}>
              <Field label="LO Name" placeholder="e.g. Robert Fisher" value={data.loName} onChange={v=>update("loName",v)} />
              <Field label="Years in the Industry" placeholder="e.g. 1.5 yrs" value={data.yearsInIndustry} onChange={v=>update("yearsInIndustry",v)} />
              <div>
                <label style={labelStyle}>Encompass Experience</label>
                <select style={{ ...inputStyle, width:"100%" }} value={data.encompassExperience} onChange={e=>update("encompassExperience",e.target.value)}>
                  <option>No</option>
                  <option>Yes</option>
                </select>
              </div>
              <Field label="Attendance" placeholder="e.g. 100%" value={data.attendance} onChange={v=>update("attendance",v)} />
              <div style={{ gridColumn:"1 / -1" }}>
                <label style={{ ...labelStyle, color:"#b45309" }}>⭐ Notable Strengths</label>
                <textarea style={{ ...inputStyle, width:"100%", minHeight:60, resize:"vertical" }}
                  placeholder="Describe strengths…" value={data.notableStrengths} onChange={e=>update("notableStrengths",e.target.value)} />
              </div>
              <div style={{ gridColumn:"1 / -1" }}>
                <label style={{ ...labelStyle, color:"#b45309" }}>⭐ Notable Weaknesses</label>
                <textarea style={{ ...inputStyle, width:"100%", minHeight:60, resize:"vertical" }}
                  placeholder="Describe weaknesses…" value={data.notableWeaknesses} onChange={e=>update("notableWeaknesses",e.target.value)} />
              </div>
              <div style={{ gridColumn:"1 / -1" }}>
                <Field label="Unique Fact" placeholder="Fun fact shared by the LO" value={data.uniqueFact} onChange={v=>update("uniqueFact",v)} />
              </div>
            </div>
            <SectionHeader style={{ marginTop:16 }}>Quiz &amp; Exam Scores</SectionHeader>
            <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:"12px 20px" }}>
              <Field label="NAF Details Quiz" placeholder="e.g. 100%" value={data.nafDetailsQuiz} onChange={v=>update("nafDetailsQuiz",v)} />
              <Field label="NAF Link Quiz" placeholder="e.g. 60%, 90%" value={data.nafLinkQuiz} onChange={v=>update("nafLinkQuiz",v)} />
              <Field label="Encompass Quiz" placeholder="e.g. 70%, 80%" value={data.encompassQuiz} onChange={v=>update("encompassQuiz",v)} />
              <Field label="Final Exam" placeholder="e.g. 68%, 78%" value={data.finalExam} onChange={v=>update("finalExam",v)} />
            </div>
            <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", marginTop:24, marginBottom:10 }}>
              <span style={{ fontSize:15, fontWeight:700, color:"#1e3a5f" }}>Preview — Copy &amp; Paste into Outlook</span>
              <button onClick={copyTable} style={{
                background: copied ? "#22c55e" : "#1e3a5f", color:"#fff",
                border:"none", borderRadius:6, padding:"8px 22px",
                fontSize:14, fontWeight:600, cursor:"pointer", transition:"background 0.2s",
              }}>
                {copied ? "✓ Copied!" : "Copy Table"}
              </button>
            </div>
            <div ref={tableRef} style={{ display:"inline-block", width:"100%" }}>
              <table border="1" cellPadding="6" cellSpacing="0" style={{
                borderCollapse:"collapse", width:"100%",
                fontFamily:"Calibri, Arial, sans-serif", fontSize:13, border:"1px solid #4472c4",
              }}>
                <tbody>
                  <tr>
                    <td colSpan={2} style={{
                      background:"#1e3a5f", color:"#fff", fontWeight:"bold",
                      textAlign:"center", fontSize:14, padding:"7px 10px", border:"1px solid #4472c4",
                    }}>
                      {data.loName ? `${data.loName} - New Hire Training Review` : "[LO Name] - New Hire Training Review"}
                    </td>
                  </tr>
                  <Row label="Years in the Industry" value={data.yearsInIndustry} />
                  <Row label="Encompass Experience" value={data.encompassExperience} />
                  <Row label="Attendance" value={data.attendance} bold />
                  <Row label="Notable Strengths" value={data.notableStrengths} bold />
                  <Row label="Notable Weaknesses" value={data.notableWeaknesses} bold />
                  <Row label="Unique Fact" value={data.uniqueFact} />
                  <tr>
                    <td colSpan={2} style={{ background:"#dce6f1", border:"1px solid #4472c4", padding:3 }}>&nbsp;</td>
                  </tr>
                  <Row label="NAF Details Quiz" value={data.nafDetailsQuiz} lightHeader />
                  <Row label="NAF Link Quiz" value={data.nafLinkQuiz} lightHeader />
                  <Row label="Encompass Quiz" value={data.encompassQuiz} lightHeader />
                  <Row label="Final Exam" value={data.finalExam} lightHeader />
                </tbody>
              </table>
            </div>
          </div>
        )}

      </div>
    </div>
  );
}

// ── sub-components ────────────────────────────────────────────────────────────
function SectionHeader({ children, style }) {
  return (
    <div style={{ fontSize:13, fontWeight:700, color:"#1e3a5f", marginBottom:10,
      borderBottom:"2px solid #dce6f1", paddingBottom:6, ...style }}>
      {children}
    </div>
  );
}

function Field({ label, placeholder, value, onChange }) {
  return (
    <div>
      <label style={labelStyle}>{label}</label>
      <input style={inputStyle} placeholder={placeholder} value={value} onChange={e=>onChange(e.target.value)} />
    </div>
  );
}

function Row({ label, value, bold, lightHeader }) {
  return (
    <tr>
      <td style={{
        background: lightHeader ? "#dce6f1" : "#4472c4",
        color: lightHeader ? "#1e3a5f" : "#fff",
        fontWeight:"bold", width:"36%", padding:"6px 10px",
        border:"1px solid #4472c4", verticalAlign:"top", fontSize:13,
      }}>{label}</td>
      <td style={{
        background:"#fff", padding:"6px 10px", border:"1px solid #4472c4",
        verticalAlign:"top", fontSize:13, fontWeight: bold ? "bold" : "normal",
      }}>{value}</td>
    </tr>
  );
}

const navBtnStyle = (disabled) => ({
  background: disabled ? "#e5e7eb" : "#4472c4",
  color: disabled ? "#aaa" : "#fff",
  border:"none", borderRadius:6, padding:"7px 18px",
  fontSize:13, fontWeight:600, cursor: disabled ? "not-allowed" : "pointer",
});

const labelStyle = { display:"block", fontSize:12, fontWeight:600, color:"#1e3a5f", marginBottom:3 };
const inputStyle = {
  width:"100%", border:"1px solid #bcd0ea", borderRadius:4,
  padding:"5px 8px", fontSize:13, fontFamily:"Calibri, Arial, sans-serif",
  boxSizing:"border-box", background:"#fff",
};

"use client";

import React, { useMemo, useState } from "react";
import * as XLSX from "xlsx";
import { Bot, ExternalLink, FileSpreadsheet, Trash2, Sparkles, UploadCloud, Download } from "lucide-react";
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer } from "recharts";

const CHATBOT_URL =
  "https://chatgpt.com/g/g-6a1007ef6e308191a9ae7ca31590649c-math-exam-design-assistant";

const TEACHER_WORKBOOK_SHEETS = { students: "BAI_LAM_HOC_SINH" };

const DEFAULT_ANSWER_KEY = {
  mc: ["A", "B", "C", "D", "A", "C", "B", "D", "A", "B", "C", "D"],
  ds: [
    ["D", "S", "D", "S"],
    ["S", "D", "D", "S"],
    ["D", "D", "S", "S"],
    ["S", "S", "D", "D"],
  ],
  nr: [2.5, 4, 0.6, 12, 8, 15],
};

const quickPrompts = [
  "Create a mathematics test blueprint aligned with the Vietnamese 2018 General Education Curriculum.",
  "Generate multiple-choice, true/false, and numerical-response mathematics questions aligned with the Vietnamese 2018 curriculum.",
  "Create answer keys and scoring rubrics for this mathematics test.",
  "Review this mathematics test for curriculum alignment, clarity, scoring rules, and item quality.",
  "Improve the distractors in these multiple-choice mathematics questions.",
];

const researchEvidence = [
  ["Use case", "Post-test item diagnosis"],
  ["Test structure", "12 MC, 4 Đ/S, 6 NR"],
  ["Main output", "Keep / revise / remove questions"],
  ["For teachers", "Question-bank improvement"],
];

function normalizeAnswer(value) {
  return String(value ?? "").trim().toUpperCase();
}

function isBlankAnswer(value) {
  return value === null || value === undefined || String(value).trim() === "";
}

function normalizeDS(value) {
  const v = normalizeAnswer(value);
  if (["D", "Đ", "ĐÚNG", "TRUE", "T", "1"].includes(v)) return "D";
  if (["S", "SAI", "FALSE", "F", "0"].includes(v)) return "S";
  return v;
}

function numericEqual(response, key, tolerance = 0.001) {
  const x = Number(String(response ?? "").replace(",", "."));
  const y = Number(key);
  if (!Number.isFinite(x) || !Number.isFinite(y)) return false;
  return Math.abs(x - y) <= tolerance;
}

function mean(values) {
  if (!values.length) return 0;
  return values.reduce((s, x) => s + x, 0) / values.length;
}

function sd(values) {
  if (values.length < 2) return 0;
  const m = mean(values);
  return Math.sqrt(values.reduce((s, x) => s + Math.pow(x - m, 2), 0) / (values.length - 1));
}

function percent(values, condition) {
  if (!values.length) return 0;
  return (values.filter(condition).length / values.length) * 100;
}

function normalizeHeaderName(value) {
  return String(value ?? "")
    .replace(String.fromCharCode(65279), "")
    .trim()
    .toLowerCase()
    .replaceAll(" ", "_");
}

function normalizeSheetName(name) {
  return String(name || "").trim().toUpperCase();
}

function findSheetName(workbook, target) {
  const targetName = normalizeSheetName(target);
  return workbook.SheetNames.find((name) => normalizeSheetName(name) === targetName);
}

function makeStudentResponseColumns() {
  return [
    "ma_hoc_sinh",
    ...Array.from({ length: 12 }, (_, i) => `D001_C${i + 1}`),
    ...Array.from({ length: 4 }, (_, q) => ["a", "b", "c", "d"].map((part) => `D001_C${q + 13}${part}`)).flat(),
    ...Array.from({ length: 6 }, (_, i) => `D001_C${i + 17}`),
  ];
}

function answerKeyToText(key) {
  const nl = String.fromCharCode(10);
  return {
    mc: key.mc.join(","),
    ds: key.ds.map((row) => row.join(",")).join(nl),
    nr: key.nr.join(","),
  };
}

function parseList(text, fallback, normalizer = normalizeAnswer) {
  const values = String(text || "")
    .replaceAll(";", ",")
    .split(",")
    .map((x) => normalizer(x))
    .filter((x) => x !== "");
  return values.length ? values : fallback;
}

function parseDSText(text, fallback) {
  const rows = String(text || "")
    .trim()
    .split(String.fromCharCode(10))
    .map((line) =>
      line
        .replaceAll(";", ",")
        .split(",")
        .map((x) => normalizeDS(x))
        .filter((x) => x !== "")
    )
    .filter((row) => row.length > 0);
  return rows.length === 4 && rows.every((row) => row.length === 4) ? rows : fallback;
}

function parseNumberList(text, fallback) {
  const values = String(text || "")
    .replaceAll(";", ",")
    .split(",")
    .map((x) => Number(String(x).trim().replace(",", ".")))
    .filter(Number.isFinite);
  return values.length ? values : fallback;
}

function dsQuestionScore(correctParts) {
  if (correctParts === 1) return 0.1;
  if (correctParts === 2) return 0.25;
  if (correctParts === 3) return 0.5;
  if (correctParts === 4) return 1;
  return 0;
}

function scoreStudentRow(row, header, answerKey, index) {
  const get = (name) => {
    const i = header.indexOf(normalizeHeaderName(name));
    return i >= 0 ? row[i] : "";
  };

  const itemScores = {};
  const mcResponses = {};
  const sections = { mc: 0, ds: 0, nr: 0 };

  for (let i = 1; i <= 12; i += 1) {
    const response = get(`D001_C${i}`);
    const normalized = normalizeAnswer(response);
    const correct = !isBlankAnswer(response) && normalized === normalizeAnswer(answerKey.mc[i - 1]);
    itemScores[`item${i}`] = correct ? 1 : 0;
    mcResponses[`item${i}`] = ["A", "B", "C", "D"].includes(normalized) ? normalized : "";
    if (correct) sections.mc += 0.25;
  }

  for (let q = 13; q <= 16; q += 1) {
    let correctParts = 0;
    ["a", "b", "c", "d"].forEach((part, partIndex) => {
      const itemNumber = 12 + (q - 13) * 4 + partIndex + 1;
      const response = get(`D001_C${q}${part}`);
      const correct = !isBlankAnswer(response) && normalizeDS(response) === normalizeDS(answerKey.ds[q - 13][partIndex]);
      itemScores[`item${itemNumber}`] = correct ? 1 : 0;
      correctParts += correct ? 1 : 0;
    });
    sections.ds += dsQuestionScore(correctParts);
  }

  for (let q = 17; q <= 22; q += 1) {
    const itemNumber = 28 + (q - 16);
    const response = get(`D001_C${q}`);
    const correct = !isBlankAnswer(response) && numericEqual(response, answerKey.nr[q - 17]);
    itemScores[`item${itemNumber}`] = correct ? 1 : 0;
    if (correct) sections.nr += 0.5;
  }

  Object.keys(sections).forEach((key) => {
    sections[key] = Number(sections[key].toFixed(2));
  });

  return {
    id: String(get("ma_hoc_sinh") || get("student_id") || `HS${index + 1}`),
    total: Number((sections.mc + sections.ds + sections.nr).toFixed(2)),
    sections,
    itemScores,
    mcResponses,
  };
}

function parseStudentAnswerTable(table, answerKey) {
  if (!table || table.length < 2) return null;
  const header = table[0].map((h) => normalizeHeaderName(h));
  if (!header.includes("d001_c1")) return null;
  const itemColumns = Array.from({ length: 34 }, (_, i) => ({ name: `item${i + 1}` }));
  const rows = table
    .slice(1)
    .map((row, index) => scoreStudentRow(row, header, answerKey, index))
    .filter((row) => Number.isFinite(row.total));
  return { rows, itemColumns, inputType: "student_answers" };
}

function parseExcel(arrayBuffer, answerKey) {
  const workbook = XLSX.read(arrayBuffer, { type: "array" });
  const preferred = findSheetName(workbook, TEACHER_WORKBOOK_SHEETS.students);
  const sheets = preferred ? [preferred, ...workbook.SheetNames.filter((x) => x !== preferred)] : workbook.SheetNames;
  for (const sheetName of sheets) {
    const table = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: "" });
    const parsed = parseStudentAnswerTable(table, answerKey);
    if (parsed && parsed.rows.length) return parsed;
  }
  return { rows: [], itemColumns: [], inputType: "unknown" };
}

function makeSampleRows(answerKey) {
  const perfect = [...answerKey.mc, ...answerKey.ds.flat(), ...answerKey.nr];
  const wrong = perfect.map((x, i) => {
    if (i < 12) return i % 4 === 0 ? "A" : x;
    if (i < 28) return i % 5 === 0 ? (x === "D" ? "S" : "D") : x;
    return i % 3 === 0 ? Number(x) + 1 : x;
  });
  return [makeStudentResponseColumns(), ["HS001", ...perfect], ["HS002", ...wrong], ["HS003", ...perfect.map((x, i) => (i % 6 === 0 ? "" : x))]];
}

function getItemSection(itemName) {
  const n = Number(String(itemName).match(/[0-9]+/)?.[0] || 0);
  if (n >= 1 && n <= 12) return "MC";
  if (n >= 13 && n <= 28) return "Đ/S";
  if (n >= 29 && n <= 34) return "NR";
  return "Other";
}

function displayItemName(itemName) {
  const n = Number(String(itemName).match(/[0-9]+/)?.[0] || 0);
  if (n >= 1 && n <= 12) return `MC${n}`;
  if (n >= 13 && n <= 28) {
    const q = Math.floor((n - 13) / 4) + 1;
    const part = ["a", "b", "c", "d"][(n - 13) % 4];
    return `DS${q}${part}`.toUpperCase();
  }
  if (n >= 29 && n <= 34) return `NR${n - 28}`;
  return String(itemName).toUpperCase();
}

function itemDifficultyAndDiscrimination(rows, itemName, upper, lower) {
  const all = rows.map((r) => Number(r.itemScores?.[itemName])).filter(Number.isFinite);
  if (!all.length) return { difficulty: 0, discrimination: 0 };
  const difficulty = mean(all);
  const upperMean = mean(upper.map((r) => Number(r.itemScores?.[itemName])).filter(Number.isFinite));
  const lowerMean = mean(lower.map((r) => Number(r.itemScores?.[itemName])).filter(Number.isFinite));
  return { difficulty, discrimination: upperMean - lowerMean };
}

function classifyItemQuality(difficulty, discrimination) {
  const labels = [];
  if (difficulty >= 0.85) labels.push("Câu quá dễ");
  if (difficulty <= 0.3) labels.push("Câu quá khó");
  if (discrimination < 0) labels.push("Câu phân hóa âm");
  else if (discrimination < 0.2) labels.push("Câu phân hóa kém");
  return labels.length ? labels.join("; ") : "Câu tạm ổn";
}

function itemQualityAdvice(classification) {
  if (classification.includes("phân hóa âm")) return "Cần kiểm tra lại đáp án, cách hỏi hoặc dữ liệu nhập.";
  if (classification.includes("quá dễ")) return "Hầu hết học sinh làm đúng; ít giá trị phân loại.";
  if (classification.includes("quá khó")) return "Nhiều học sinh sai; cần xem có vượt yêu cầu cần đạt không.";
  if (classification.includes("phân hóa kém")) return "Nhóm học sinh mạnh và yếu làm đúng gần như nhau.";
  return "Có thể giữ lại, nhưng giáo viên vẫn nên rà soát nội dung.";
}

function mcDistractorStatus(rows, itemName, answerKey) {
  const n = Number(String(itemName).match(/[0-9]+/)?.[0] || 0);
  if (n < 1 || n > 12) return "";
  if (!rows.some((row) => row.mcResponses && Object.prototype.hasOwnProperty.call(row.mcResponses, itemName))) return "";
  const correctAnswer = normalizeAnswer(answerKey.mc[n - 1]);
  const options = ["A", "B", "C", "D"];
  const counts = Object.fromEntries(options.map((option) => [option, 0]));
  rows.forEach((row) => {
    const response = normalizeAnswer(row.mcResponses?.[itemName]);
    if (options.includes(response)) counts[response] += 1;
  });
  const zeroDistractors = options.filter((option) => option !== correctAnswer && counts[option] === 0);
  return zeroDistractors.length ? `Distractor không hoạt động: ${zeroDistractors.join(", ")}` : "";
}

function bankDecision(item) {
  if (item.classification.includes("phân hóa âm")) return "Loại bỏ hoặc kiểm tra lại ngay";
  if (item.difficulty >= 0.95 || item.difficulty <= 0.15) return "Loại bỏ khỏi ngân hàng hiện tại";
  if (item.classification.includes("Distractor")) return "Sửa lại phương án nhiễu";
  if (item.classification.includes("quá dễ") || item.classification.includes("quá khó") || item.classification.includes("phân hóa kém")) return "Sửa lại trước khi dùng lại";
  if (item.discrimination >= 0.2 && item.difficulty > 0.3 && item.difficulty < 0.85) return "Giữ lại trong ngân hàng câu hỏi";
  return "Giáo viên rà soát thêm";
}

function calculateItemAnalysis(rows, itemColumns, answerKey) {
  if (!rows.length || !itemColumns.length) return [];
  const sorted = [...rows].sort((a, b) => b.total - a.total);
  const groupSize = Math.max(1, Math.floor(rows.length * 0.27));
  const upper = sorted.slice(0, groupSize);
  const lower = sorted.slice(-groupSize);

  return itemColumns.map((col) => {
    const result = itemDifficultyAndDiscrimination(rows, col.name, upper, lower);
    const baseClass = classifyItemQuality(result.difficulty, result.discrimination);
    const distractor = mcDistractorStatus(rows, col.name, answerKey);
    const classification = distractor ? `${baseClass}; ${distractor}` : baseClass;
    const item = {
      item: displayItemName(col.name),
      rawItem: col.name,
      section: getItemSection(col.name),
      difficulty: result.difficulty,
      discrimination: result.discrimination,
      classification,
      advice: distractor ? `${itemQualityAdvice(baseClass)} ${distractor}.` : itemQualityAdvice(baseClass),
    };
    return { ...item, bankDecision: bankDecision(item) };
  });
}

function calculateSectionStats(rows) {
  if (!rows.length) return [];
  return [
    { key: "mc", label: "MC", maxScore: 3 },
    { key: "ds", label: "Đ/S", maxScore: 4 },
    { key: "nr", label: "NR", maxScore: 3 },
  ].map((section) => {
    const scores = rows.map((row) => row.sections?.[section.key] || 0);
    const difficulty = mean(scores) / section.maxScore;
    let level = "Phù hợp";
    if (difficulty >= 0.85) level = "Quá dễ";
    if (difficulty <= 0.3) level = "Quá khó";
    return { ...section, sectionName: section.label, section: section.label, mean: mean(scores), sd: sd(scores), difficulty, level };
  });
}

function MetricCard({ label, value, note }) {
  return (
    <div className="group rounded-3xl border border-slate-200/80 bg-white/90 p-5 shadow-sm ring-1 ring-slate-950/5 transition hover:-translate-y-0.5 hover:shadow-xl">
      <p className="text-sm font-semibold text-slate-500">{label}</p>
      <p className="mt-2 text-3xl font-black tracking-tight text-slate-950">{value}</p>
      {note && <p className="mt-2 text-sm leading-6 text-slate-500">{note}</p>}
    </div>
  );
}

function SectionHeader({ label, title, description }) {
  return (
    <div className="mb-8">
      <div className="inline-flex items-center gap-2 rounded-full border border-slate-200 bg-white px-3 py-1 text-xs font-bold uppercase tracking-wide text-slate-500 shadow-sm">
        <Sparkles className="h-3.5 w-3.5" /> {label}
      </div>
      <h2 className="mt-4 text-3xl font-black tracking-tight text-slate-950 md:text-4xl">{title}</h2>
      <p className="mt-3 max-w-3xl text-base leading-7 text-slate-600">{description}</p>
    </div>
  );
}

export default function Page() {
  const [keyText, setKeyText] = useState(answerKeyToText(DEFAULT_ANSWER_KEY));
  const [rows, setRows] = useState([]);
  const [itemColumns, setItemColumns] = useState([]);
  const [uploadMessage, setUploadMessage] = useState("No student dataset uploaded yet.");

  const answerKey = useMemo(
    () => ({
      mc: parseList(keyText.mc, DEFAULT_ANSWER_KEY.mc),
      ds: parseDSText(keyText.ds, DEFAULT_ANSWER_KEY.ds),
      nr: parseNumberList(keyText.nr, DEFAULT_ANSWER_KEY.nr),
    }),
    [keyText]
  );

  const stats = useMemo(() => {
    if (!rows.length) return null;
    const scores = rows.map((r) => r.total);
    return {
      n: rows.length,
      mean: mean(scores),
      sd: sd(scores),
      min: Math.min(...scores),
      max: Math.max(...scores),
      pct9: percent(scores, (x) => x >= 9),
      pct10: percent(scores, (x) => x === 10),
    };
  }, [rows]);

  const itemAnalysis = useMemo(() => calculateItemAnalysis(rows, itemColumns, answerKey), [rows, itemColumns, answerKey]);
  const sectionStats = useMemo(() => calculateSectionStats(rows), [rows]);
  const easiest = sectionStats.length ? [...sectionStats].sort((a, b) => b.difficulty - a.difficulty)[0] : null;
  const hardest = sectionStats.length ? [...sectionStats].sort((a, b) => a.difficulty - b.difficulty)[0] : null;
  const keep = itemAnalysis.filter((x) => x.bankDecision.includes("Giữ lại"));
  const revise = itemAnalysis.filter((x) => x.bankDecision.includes("Sửa lại") || x.bankDecision.includes("rà soát"));
  const remove = itemAnalysis.filter((x) => x.bankDecision.includes("Loại bỏ"));

  function updateKeyText(field, value) {
    setKeyText((current) => ({ ...current, [field]: value }));
  }

  function openChatbot() {
    window.location.href = CHATBOT_URL;
  }

  async function handleUpload(event) {
    const file = event.target.files?.[0];
    if (!file) return;
    setUploadMessage(`Reading file: ${file.name} ...`);
    try {
      if (!file.name.toLowerCase().endsWith(".xlsx") && !file.name.toLowerCase().endsWith(".xls")) {
        setUploadMessage("Upload failed. Please upload an Excel file only (.xlsx or .xls).");
        return;
      }
      const parsed = parseExcel(await file.arrayBuffer(), answerKey);
      setRows(parsed.rows);
      setItemColumns(parsed.itemColumns);
      setUploadMessage(parsed.rows.length ? `Upload successful: ${parsed.rows.length} valid student rows detected.` : "Upload failed. Required columns include D001_C1. Recommended sheet name: BAI_LAM_HOC_SINH.");
    } catch (error) {
      setRows([]);
      setItemColumns([]);
      setUploadMessage(`Upload failed. Details: ${error?.message || error}`);
    } finally {
      event.target.value = "";
    }
  }

  function loadSampleData() {
    try {
      setUploadMessage("Loading sample data...");
      const parsed = parseStudentAnswerTable(makeSampleRows(answerKey), answerKey);
      if (!parsed || !parsed.rows.length) {
        setUploadMessage("Sample loading failed. Please check the answer-key format above.");
        return;
      }
      setRows(parsed.rows);
      setItemColumns(parsed.itemColumns);
      setUploadMessage(`Student-answer sample loaded: ${parsed.rows.length} student rows detected.`);
    } catch (error) {
      setUploadMessage(`Sample loading failed. Details: ${error?.message || error}`);
    }
  }

  function saveBlob(blob, fileName) {
    const url = window.URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = fileName;
    link.style.display = "none";
    document.body.appendChild(link);
    link.click();
    window.setTimeout(() => {
      document.body.removeChild(link);
      window.URL.revokeObjectURL(url);
    }, 250);
  }

  function downloadExcelTemplate() {
    try {
      setUploadMessage("Preparing Excel template...");
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(
        workbook,
        XLSX.utils.aoa_to_sheet(makeSampleRows(answerKey)),
        TEACHER_WORKBOOK_SHEETS.students
      );
      const array = XLSX.write(workbook, { bookType: "xlsx", type: "array" });
      saveBlob(
        new Blob([array], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }),
        "student-answer-template-one-form.xlsx"
      );
      setUploadMessage("Excel template downloaded.");
    } catch (error) {
      setUploadMessage(`Excel template download failed. Details: ${error?.message || error}`);
    }
  }

  function clearData() {
    setRows([]);
    setItemColumns([]);
    setUploadMessage("Data cleared. No student dataset uploaded yet.");
  }

  return (
    <main className="min-h-screen bg-[radial-gradient(circle_at_top_left,#e0f2fe,transparent_34%),radial-gradient(circle_at_top_right,#ede9fe,transparent_32%),linear-gradient(to_bottom,#f8fafc,#eef2ff)] text-slate-950">
      <section className="relative overflow-hidden border-b border-white/70 bg-white/75 backdrop-blur">
        <div className="absolute -right-20 -top-24 h-72 w-72 rounded-full bg-indigo-200/60 blur-3xl" />
        <div className="absolute -left-20 top-16 h-72 w-72 rounded-full bg-sky-200/60 blur-3xl" />
        <div className="relative mx-auto max-w-7xl px-6 py-14 md:py-20">
          <div className="grid gap-10 md:grid-cols-[1.3fr_0.7fr] md:items-center">
            <div>
              <div className="inline-flex items-center gap-2 rounded-full border border-slate-200 bg-white px-4 py-2 text-sm font-bold text-slate-600 shadow-sm">
                <Sparkles className="h-4 w-4 text-indigo-600" /> AI Assessment Lab
              </div>
              <h1 className="mt-5 max-w-4xl text-4xl font-black tracking-tight text-slate-950 md:text-6xl">
                Test Quality Diagnosis for Teachers
              </h1>
              <p className="mt-5 max-w-3xl text-lg leading-8 text-slate-600">
                A one-form web tool for entering answer keys, scoring student responses, diagnosing item quality, and making question-bank decisions after each classroom test.
              </p>
              <div className="mt-7 flex flex-wrap gap-3">
                <button type="button" onClick={openChatbot} className="inline-flex items-center gap-2 rounded-2xl bg-slate-950 px-6 py-3 text-sm font-bold text-white shadow-lg shadow-slate-900/20 transition hover:-translate-y-0.5 hover:bg-indigo-700">
                  Start Chatbot <ExternalLink className="h-4 w-4" />
                </button>
                <a href="#student-data" className="inline-flex items-center gap-2 rounded-2xl border border-slate-200 bg-white px-6 py-3 text-sm font-bold text-slate-800 shadow-sm transition hover:-translate-y-0.5 hover:shadow-md">
                  Upload Excel <UploadCloud className="h-4 w-4" />
                </a>
              </div>
            </div>
            <div className="rounded-[2rem] border border-white/80 bg-white/80 p-6 shadow-2xl shadow-slate-900/10 backdrop-blur">
              <div className="grid gap-4">
                {researchEvidence.map(([k, v], index) => (
                  <div key={`hero-evidence-${k}-${index}`} className="rounded-2xl bg-slate-50 p-4 ring-1 ring-slate-200/70">
                    <p className="text-xs font-bold uppercase tracking-wide text-slate-500">{k}</p>
                    <p className="mt-1 text-lg font-black text-slate-950">{v}</p>
                  </div>
                ))}
              </div>
            </div>
          </div>
        </div>
      </section>

      <section className="mx-auto max-w-7xl px-6 py-14">
        <SectionHeader label="1. Chatbot" title="Design assistant" description="Use the chatbot to draft a test blueprint, generate items, improve distractors, and review scoring rubrics." />
        <div className="grid gap-4 md:grid-cols-2">
          {quickPrompts.map((prompt, index) => (
            <div key={`prompt-${index}`} className="rounded-3xl border border-slate-200/80 bg-white/90 p-6 shadow-sm ring-1 ring-slate-950/5 transition hover:-translate-y-1 hover:shadow-xl">
              <div className="flex h-10 w-10 items-center justify-center rounded-2xl bg-indigo-50 text-indigo-700"><Bot className="h-5 w-5" /></div>
              <p className="mt-3 text-sm text-slate-700">{prompt}</p>
            </div>
          ))}
        </div>
      </section>

      <section className="mx-auto max-w-7xl px-6 py-14">
        <SectionHeader label="2. Answer key" title="Enter the answer key for one test" description="MC uses A/B/C/D. Đ/S uses D/S. NR uses numerical answers. Blank student responses are scored as incorrect." />
        <div className="grid gap-5 md:grid-cols-3">
          <label className="rounded-3xl border border-slate-200/80 bg-white/90 p-5 shadow-sm ring-1 ring-slate-950/5">
            <p className="font-semibold">MC key, C1–C12</p>
            <textarea value={keyText.mc} onChange={(e) => updateKeyText("mc", e.target.value)} rows={5} className="mt-3 w-full rounded-2xl border border-slate-200 bg-slate-50 p-3 text-sm outline-none transition focus:border-indigo-400 focus:bg-white focus:ring-4 focus:ring-indigo-100" />
          </label>
          <label className="rounded-3xl border border-slate-200/80 bg-white/90 p-5 shadow-sm ring-1 ring-slate-950/5">
            <p className="font-semibold">Đ/S key, C13–C16</p>
            <textarea value={keyText.ds} onChange={(e) => updateKeyText("ds", e.target.value)} rows={5} className="mt-3 w-full rounded-2xl border border-slate-200 bg-slate-50 p-3 text-sm outline-none transition focus:border-indigo-400 focus:bg-white focus:ring-4 focus:ring-indigo-100" />
          </label>
          <label className="rounded-3xl border border-slate-200/80 bg-white/90 p-5 shadow-sm ring-1 ring-slate-950/5">
            <p className="font-semibold">NR key, C17–C22</p>
            <textarea value={keyText.nr} onChange={(e) => updateKeyText("nr", e.target.value)} rows={5} className="mt-3 w-full rounded-2xl border border-slate-200 bg-slate-50 p-3 text-sm outline-none transition focus:border-indigo-400 focus:bg-white focus:ring-4 focus:ring-indigo-100" />
          </label>
        </div>
      </section>

      <section id="student-data" className="mx-auto max-w-7xl px-6 py-14">
        <SectionHeader label="3. Student data" title="Upload student answers" description="Download the Excel template from /public/templates, then fill student answers. The upload file needs one sheet named BAI_LAM_HOC_SINH. Columns: ma_hoc_sinh, D001_C1...D001_C22." />
        <div className="rounded-[2rem] border border-slate-200/80 bg-white/90 p-6 shadow-xl shadow-slate-900/5 ring-1 ring-slate-950/5">
          <div className="flex flex-wrap gap-3">
            <label className="inline-flex cursor-pointer items-center gap-2 rounded-2xl bg-slate-950 px-5 py-3 text-sm font-bold text-white shadow-lg shadow-slate-900/15 transition hover:-translate-y-0.5 hover:bg-indigo-700">
              <UploadCloud className="h-4 w-4" /> Upload file
              <input type="file" accept=".xlsx,.xls" className="hidden" onClick={(e) => { e.currentTarget.value = ""; }} onChange={handleUpload} />
            </label>
            <button type="button" onClick={() => loadSampleData()} className="rounded-2xl border border-slate-200 bg-white px-5 py-3 text-sm font-bold text-slate-800 shadow-sm transition hover:-translate-y-0.5 hover:shadow-md">Load sample</button>
            <a
              href="/templates/student-answer-template-one-form.xlsx"
              download="student-answer-template-one-form.xlsx"
              className="inline-flex items-center gap-2 rounded-2xl border border-slate-200 bg-white px-5 py-3 text-sm font-bold text-slate-800 shadow-sm transition hover:-translate-y-0.5 hover:shadow-md"
            >
              <Download className="h-4 w-4" /> Download Excel template
            </a>
            <button type="button" onClick={() => clearData()} className="inline-flex items-center gap-2 rounded-2xl border border-slate-200 bg-white px-5 py-3 text-sm font-bold text-slate-800 shadow-sm transition hover:-translate-y-0.5 hover:shadow-md"><Trash2 className="h-4 w-4" /> Clear data</button>
          </div>
          <p className="mt-4 rounded-2xl bg-slate-50 px-4 py-3 text-sm font-medium text-slate-600 ring-1 ring-slate-200">{uploadMessage}</p>
        </div>
      </section>

      <section className="mx-auto max-w-7xl px-6 py-14">
        <SectionHeader label="4. Results" title="Score report and item diagnosis" description="The website identifies easy/difficult sections, weak items, negative discrimination, non-functioning distractors, and question-bank decisions." />

        {stats && (
          <>
            <div className="grid gap-4 md:grid-cols-4">
              <MetricCard label="Detected N" value={stats.n} note="valid student rows" />
              <MetricCard label="Average score" value={stats.mean.toFixed(3)} note={`SD = ${stats.sd.toFixed(3)}`} />
              <MetricCard label="Min–Max" value={`${stats.min.toFixed(2)}–${stats.max.toFixed(2)}`} note="total score" />
              <MetricCard label="% ≥ 9" value={`${stats.pct9.toFixed(1)}%`} note={`% = 10: ${stats.pct10.toFixed(1)}%`} />
            </div>

            <div className="mt-6 overflow-hidden rounded-[2rem] border border-slate-200/80 bg-white/90 shadow-xl shadow-slate-900/5 ring-1 ring-slate-950/5">
              <div className="border-b p-5">
                <h3 className="text-lg font-bold">Section analysis: MC, Đ/S, NR</h3>
                <p className="mt-2 text-sm text-slate-600">This table shows which part of the test is too easy or too difficult.</p>
              </div>
              <table className="w-full text-left text-sm">
                <thead className="bg-slate-100 text-slate-700">
                  <tr><th className="p-4">Section</th><th className="p-4">Max score</th><th className="p-4">Mean</th><th className="p-4">Difficulty</th><th className="p-4">Level</th></tr>
                </thead>
                <tbody>
                  {sectionStats.map((s, index) => (
                    <tr key={`section-${s.sectionName || s.label || s.key}-${index}`} className="border-t"><td className="p-4 font-medium">{s.sectionName || s.label || s.key?.toUpperCase()}</td><td className="p-4">{s.maxScore}</td><td className="p-4">{s.mean.toFixed(3)}</td><td className="p-4">{s.difficulty.toFixed(3)}</td><td className="p-4 font-semibold">{s.level}</td></tr>
                  ))}
                </tbody>
              </table>
            </div>

            <div className="mt-6 grid gap-4 md:grid-cols-3">
              <MetricCard label="Phần dễ nhất" value={easiest?.sectionName || easiest?.label || easiest?.key?.toUpperCase() || "-"} note={easiest ? `Difficulty = ${easiest.difficulty.toFixed(3)}` : ""} />
              <MetricCard label="Phần khó nhất" value={hardest?.sectionName || hardest?.label || hardest?.key?.toUpperCase() || "-"} note={hardest ? `Difficulty = ${hardest.difficulty.toFixed(3)}` : ""} />
              <MetricCard label="Dạng HS gặp khó khăn" value={hardest?.sectionName || hardest?.label || hardest?.key?.toUpperCase() || "-"} note="Dựa trên điểm trung bình theo phần" />
            </div>

            <div className="mt-6 rounded-[2rem] border border-slate-200/80 bg-white/90 p-6 shadow-xl shadow-slate-900/5 ring-1 ring-slate-950/5">
              <h3 className="text-lg font-bold">Question bank recommendations</h3>
              <p className="mt-2 text-sm text-slate-600">After the test, the website recommends which questions to keep, revise, or remove from the question bank.</p>
              <div className="mt-5 grid gap-4 md:grid-cols-3">
                <MetricCard label="Giữ lại" value={keep.length} note="Câu có độ khó và phân hóa tương đối ổn" />
                <MetricCard label="Sửa lại" value={revise.length} note="Câu cần điều chỉnh trước khi dùng lại" />
                <MetricCard label="Loại bỏ" value={remove.length} note="Câu có rủi ro cao về chất lượng" />
              </div>
            </div>

            <div className="mt-6 overflow-hidden rounded-[2rem] border border-slate-200/80 bg-white/90 shadow-xl shadow-slate-900/5 ring-1 ring-slate-950/5">
              <div className="border-b p-5">
                <h3 className="text-lg font-bold">Item/sub-item analysis</h3>
                <p className="mt-2 text-sm text-slate-600">MC items are MC1–MC12. Đ/S items are DS1a–DS4d. NR items are NR1–NR6.</p>
              </div>
              <table className="w-full text-left text-sm">
                <thead className="bg-slate-100 text-slate-700">
                  <tr><th className="p-4">Item</th><th className="p-4">Section</th><th className="p-4">Difficulty</th><th className="p-4">D-index</th><th className="p-4">Classification</th><th className="p-4">Recommendation</th><th className="p-4">Teacher action</th></tr>
                </thead>
                <tbody>
                  {itemAnalysis.map((item, index) => (
                    <tr key={`item-${item.rawItem || item.item}-${index}`} className="border-t align-top"><td className="p-4 font-medium">{item.item}</td><td className="p-4">{item.section || getItemSection(item.rawItem)}</td><td className="p-4">{item.difficulty.toFixed(3)}</td><td className="p-4">{item.discrimination.toFixed(3)}</td><td className="p-4">{item.classification}</td><td className="p-4 font-semibold">{item.bankDecision}</td><td className="p-4">{item.advice}</td></tr>
                  ))}
                </tbody>
              </table>
            </div>

            <div className="mt-6 rounded-[2rem] border border-slate-200/80 bg-white/90 p-6 shadow-xl shadow-slate-900/5 ring-1 ring-slate-950/5">
              <h3 className="text-lg font-bold">Item difficulty chart</h3>
              <div className="mt-4 h-80">
                <ResponsiveContainer width="100%" height="100%">
                  <BarChart data={itemAnalysis.slice(0, 34)}>
                    <CartesianGrid strokeDasharray="3 3" />
                    <XAxis dataKey="item" interval={0} angle={-45} textAnchor="end" height={80} />
                    <YAxis domain={[0, 1]} />
                    <Tooltip />
                    <Bar dataKey="difficulty" />
                  </BarChart>
                </ResponsiveContainer>
              </div>
            </div>
          </>
        )}
      </section>

      <section className="mx-auto max-w-7xl px-6 py-14">
        <SectionHeader label="5. Research-based design" title="Why this web tool is useful" description="The tool transforms student-response data into practical decisions for teachers: keep, revise, or remove questions from the question bank." />
        <div className="grid gap-4 md:grid-cols-4">
          {researchEvidence.map(([k, v], index) => (
            <div key={`evidence-${k}-${index}`} className="rounded-3xl border border-slate-200/80 bg-white/90 p-5 shadow-sm ring-1 ring-slate-950/5"><p className="text-sm text-slate-500">{k}</p><p className="mt-2 font-black text-slate-950">{v}</p></div>
          ))}
        </div>
      </section>
    </main>
  );
}

import React, { useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";
import pptxgen from "pptxgenjs";
import { toPng } from "html-to-image";
import {
  LineChart,
  Line,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  Legend,
  ResponsiveContainer,
} from "recharts";
import "./App.css";

const REQUIRED_COLUMNS = ["Prozess", "Terminart", "Datum", "Status"];
const STATUS_VALUES = ["durchgeführt", "geplant", "offen"];

function getCalendarWeek(dateInput) {
  const date = new Date(dateInput);
  if (Number.isNaN(date.getTime())) return null;

  const tempDate = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  const dayNum = tempDate.getUTCDay() || 7;
  tempDate.setUTCDate(tempDate.getUTCDate() + 4 - dayNum);

  const yearStart = new Date(Date.UTC(tempDate.getUTCFullYear(), 0, 1));
  const weekNo = Math.ceil(((tempDate - yearStart) / 86400000 + 1) / 7);

  return `KW ${weekNo}`;
}

function normalizeStatus(status) {
  if (!status) return "offen";
  const cleaned = String(status).trim().toLowerCase();

  if (cleaned.includes("durch")) return "durchgeführt";
  if (cleaned.includes("geplant")) return "geplant";
  if (cleaned.includes("offen")) return "offen";

  return "offen";
}

function normalizeTerminart(value) {
  if (!value) return "";
  const cleaned = String(value).trim().toLowerCase();

  if (cleaned.includes("briefing")) return "Briefing";
  if (cleaned.includes("workshop")) return "Workshop";

  return value;
}

function parseExcelDate(value) {
  if (!value) return null;

  if (value instanceof Date) return value;

  if (typeof value === "number") {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (!parsed) return null;
    return new Date(parsed.y, parsed.m - 1, parsed.d);
  }

  const parts = String(value).split(".");
  if (parts.length === 3) {
    const [day, month, year] = parts;
    return new Date(Number(year), Number(month) - 1, Number(day));
  }

  return new Date(value);
}

function buildChartData(rows, terminart) {
  const filtered = rows.filter((row) => row.Terminart === terminart && row.KW);

  const weekNumbers = [
    ...new Set(
      filtered.map((row) => Number(String(row.KW).replace("KW ", "")))
    ),
  ].sort((a, b) => a - b);

  const data = [];
  const cumulative = {
    durchgeführt: 0,
    geplant: 0,
    offen: 0,
  };

  weekNumbers.forEach((week) => {
    const weekRows = filtered.filter((row) => row.KW === `KW ${week}`);

    STATUS_VALUES.forEach((status) => {
      cumulative[status] += weekRows.filter((row) => row.Status === status).length;
    });

    data.push({
      kw: `KW ${week}`,
      durchgeführt: cumulative["durchgeführt"],
      geplant: cumulative["geplant"],
      offen: cumulative["offen"],
    });
  });

  return data;
}

function SummaryBox({ title, data }) {
  const latest = data[data.length - 1];

  return (
    <div className="summaryBox">
      <h3>{title}</h3>
      <p>
        <strong>Total:</strong>{" "}
        {latest
          ? latest.durchgeführt + latest.geplant + latest.offen
          : 0}
      </p>
      <p>
        <strong>Latest KW:</strong> {latest ? latest.kw : "-"}
      </p>
      <p>
        <strong>Durchgeführt:</strong> {latest ? latest.durchgeführt : 0}
      </p>
    </div>
  );
}

function ChartCard({ title, data, chartRef }) {
  return (
    <div className="chartCard" ref={chartRef}>
      <div className="chartHeader">
        <div>
          <h2>{title}</h2>
          <p>Kumulative Entwicklung nach Kalenderwoche</p>
        </div>
        <SummaryBox title="Summary" data={data} />
      </div>

      <div className="chartArea">
        {data.length === 0 ? (
          <div className="emptyState">Noch keine Daten vorhanden.</div>
        ) : (
          <ResponsiveContainer width="100%" height={330}>
            <LineChart data={data} margin={{ top: 20, right: 30, left: 10, bottom: 10 }}>
              <CartesianGrid strokeDasharray="3 3" stroke="rgba(255, 255, 255, 0.1)" />
              <XAxis dataKey="kw" tick={{ fill: "#AEB7C8" }} axisLine={{ stroke: "rgba(255, 255, 255, 0.18)" }} tickLine={false} />
              <YAxis allowDecimals={false} tick={{ fill: "#AEB7C8" }} axisLine={{ stroke: "rgba(255, 255, 255, 0.18)" }} tickLine={false} />
              <Tooltip
                contentStyle={{
                  background: "rgba(12, 16, 26, 0.92)",
                  border: "1px solid rgba(255, 255, 255, 0.12)",
                  borderRadius: "8px",
                  color: "#F8FAFC",
                  boxShadow: "0 18px 40px rgba(0, 0, 0, 0.35)",
                }}
                labelStyle={{ color: "#FFFFFF" }}
              />
              <Legend wrapperStyle={{ color: "#D9E2F2" }} />
              <Line type="monotone" dataKey="durchgeführt" stroke="#F8FAFC" strokeWidth={3} dot={{ r: 3 }} />
              <Line type="monotone" dataKey="geplant" stroke="#7DD3FC" strokeWidth={3} dot={{ r: 3 }} />
              <Line type="monotone" dataKey="offen" stroke="#9CA3AF" strokeWidth={3} dot={{ r: 3 }} />
            </LineChart>
          </ResponsiveContainer>
        )}
      </div>
    </div>
  );
}

export default function App() {
  const [rows, setRows] = useState([]);
  const [error, setError] = useState("");
  const briefingRef = useRef(null);
  const workshopRef = useRef(null);

  const briefingData = useMemo(() => buildChartData(rows, "Briefing"), [rows]);
  const workshopData = useMemo(() => buildChartData(rows, "Workshop"), [rows]);

  async function handleFileUpload(event) {
    setError("");
    const file = event.target.files[0];
    if (!file) return;

    try {
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer);
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      const json = XLSX.utils.sheet_to_json(worksheet);

      if (!json.length) {
        setError("Die Excel-Datei enthält keine Daten.");
        return;
      }

      const columns = Object.keys(json[0]);
      const missingColumns = REQUIRED_COLUMNS.filter(
        (column) => !columns.includes(column)
      );

      if (missingColumns.length > 0) {
        setError(`Fehlende Spalten: ${missingColumns.join(", ")}`);
        return;
      }

      const cleanedRows = json.map((row) => {
        const date = parseExcelDate(row.Datum);

        return {
          Prozess: row.Prozess,
          Terminart: normalizeTerminart(row.Terminart),
          Datum: date,
          Status: normalizeStatus(row.Status),
          KW: getCalendarWeek(date),
        };
      });

      setRows(cleanedRows);
    } catch (err) {
      console.error(err);
      setError("Fehler beim Lesen der Excel-Datei.");
    }
  }

  async function exportPowerPoint() {
    if (!briefingRef.current || !workshopRef.current) return;

    const pptx = new pptxgen();
    pptx.layout = "LAYOUT_WIDE";
    pptx.author = "Decura Dashboard Tool";
    pptx.subject = "Excel to PowerPoint Dashboard";

    const slides = [
      {
        title: "Briefing Hochlaufkurve",
        ref: briefingRef,
        data: briefingData,
      },
      {
        title: "Workshop Hochlaufkurve",
        ref: workshopRef,
        data: workshopData,
      },
    ];

    for (const item of slides) {
      const slide = pptx.addSlide();

      slide.background = { color: "FFFFFF" };

      slide.addText(item.title, {
        x: 0.5,
        y: 0.3,
        w: 12,
        h: 0.4,
        fontFace: "Aptos",
        fontSize: 24,
        bold: true,
        color: "111111",
      });

      slide.addText("Automatisch generiert aus Excel-Daten", {
        x: 0.5,
        y: 0.75,
        w: 12,
        h: 0.3,
        fontFace: "Aptos",
        fontSize: 10,
        color: "666666",
      });

      const imageData = await toPng(item.ref.current, {
        cacheBust: true,
        pixelRatio: 2,
        backgroundColor: "#ffffff",
      });

      slide.addImage({
        data: imageData,
        x: 0.5,
        y: 1.25,
        w: 12.3,
        h: 5.6,
      });

      const latest = item.data[item.data.length - 1];
      const total = latest
        ? latest.durchgeführt + latest.geplant + latest.offen
        : 0;

      slide.addShape(pptx.ShapeType.rect, {
        x: 0.5,
        y: 6.95,
        w: 12.3,
        h: 0.5,
        fill: { color: "F3F4F6" },
        line: { color: "E5E7EB" },
      });

      slide.addText(`Total: ${total}   |   Latest KW: ${latest ? latest.kw : "-"}   |   Status: automatisch generiert`, {
        x: 0.75,
        y: 7.1,
        w: 11.8,
        h: 0.25,
        fontFace: "Aptos",
        fontSize: 11,
        color: "111111",
      });
    }

    await pptx.writeFile({ fileName: "Projekt-Dashboard-Hochlaufkurven.pptx" });
  }

  return (
    <main className="app">
      <section className="hero">
        <div>
          <p className="eyebrow">Excel → PowerPoint Tool</p>
          <h1>Projekt-Dashboard Generator</h1>
          <p className="subtitle">
            Lade eine Excel-Datei hoch und erstelle automatisch Hochlaufkurven
            für Briefings und Workshops.
          </p>
        </div>

        <div className="uploadBox">
          <input type="file" accept=".xlsx,.xls" onChange={handleFileUpload} />
          <button onClick={exportPowerPoint} disabled={rows.length === 0}>
            PowerPoint exportieren
          </button>
        </div>
      </section>

      {error && <div className="errorBox">{error}</div>}

      <section className="infoGrid">
        <div>
          <strong>Benötigte Spalten:</strong> Prozess, Terminart, Datum, Status
        </div>
        <div>
          <strong>Terminarten:</strong> Briefing, Workshop
        </div>
        <div>
          <strong>Status:</strong> durchgeführt, geplant, offen
        </div>
      </section>

      <section className="charts">
        <ChartCard
          title="Briefing Hochlaufkurve"
          data={briefingData}
          chartRef={briefingRef}
        />

        <ChartCard
          title="Workshop Hochlaufkurve"
          data={workshopData}
          chartRef={workshopRef}
        />
      </section>
    </main>
  );
}

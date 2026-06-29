import { useState, useCallback } from "react";
import * as XLSX from "xlsx";

const LOGO_B64 = "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";

const PLATFORM_CONFIG = {
  "Shopee": {
    label: "Shopee",
    emoji: "🛒",
    color: "#EE4D2D",
    bg: "#fff5f3",
    border: "#EE4D2D",
    match: (p) => p && p.toLowerCase().includes("shopee"),
  },
  "TikTok MCN": {
    label: "TikTok MCN",
    emoji: "🎬",
    color: "#010101",
    bg: "#f5f5f5",
    border: "#010101",
    match: (p) => p && p.toLowerCase().includes("mcn"),
  },
  "TikTok TAP": {
    label: "TikTok TAP",
    emoji: "🎵",
    color: "#69C9D0",
    bg: "#f0fdfe",
    border: "#69C9D0",
    match: (p) => p && p.toLowerCase().includes("tap"),
  },
  "Mercado Livre": {
    label: "Mercado Livre",
    emoji: "🌐",
    color: "#FFE600",
    bg: "#fffde6",
    border: "#D4A900",
    match: (p) => p && (p.toLowerCase().includes("meli") || p.toLowerCase().includes("mercado")),
  },
};

const STARLIVE_ORANGE = "#F7941D";
const STARLIVE_PURPLE = "#7B2FBE";
const STARLIVE_RED = "#E63946";

function formatCurrency(val) {
  if (!val && val !== 0) return "—";
  const num = parseFloat(val);
  if (isNaN(num)) return val;
  return "R$ " + num.toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function formatDate(val) {
  if (!val) return "—";
  if (typeof val === "string" && val.includes("T")) {
    return val.split("T")[0].split("-").reverse().join("/");
  }
  if (val instanceof Date) {
    return val.toLocaleDateString("pt-BR");
  }
  const str = String(val);
  if (str.includes("-")) return str.split("-").reverse().join("/");
  return str;
}

function StarLiveLogo({ size = 48 }) {
  return (
    <svg width={size} height={size} viewBox="0 0 200 200" fill="none" xmlns="http://www.w3.org/2000/svg">
      <polygon points="100,15 125,80 195,80 140,120 162,185 100,148 38,185 60,120 5,80 75,80" fill={STARLIVE_ORANGE} />
      <polygon points="100,15 125,80 195,80 140,120" fill={STARLIVE_PURPLE} opacity="0.85" />
      <polygon points="100,148 38,185 60,120 5,80" fill={STARLIVE_PURPLE} opacity="0.85" />
      <polygon points="162,185 100,148 140,120" fill={STARLIVE_PURPLE} opacity="0.7" />
      <circle cx="145" cy="155" r="10" fill={STARLIVE_RED} />
    </svg>
  );
}

function MetricCard({ label, value, accent }) {
  return (
    <div style={{
      background: "#fff",
      borderRadius: 12,
      padding: "16px 20px",
      borderLeft: `4px solid ${accent}`,
      boxShadow: "0 2px 8px rgba(0,0,0,0.07)",
      minWidth: 140,
      flex: 1,
    }}>
      <div style={{ fontSize: 11, fontWeight: 700, color: "#888", textTransform: "uppercase", letterSpacing: 1, marginBottom: 6 }}>{label}</div>
      <div style={{ fontSize: 20, fontWeight: 800, color: "#1a1a2e" }}>{value}</div>
    </div>
  );
}

function PlatformReport({ data, config }) {
  if (!data) {
    return (
      <div style={{ textAlign: "center", color: "#aaa", padding: 60, fontSize: 16 }}>
        Nenhum dado encontrado para esta plataforma na planilha carregada.
      </div>
    );
  }

  const gmv = data["Faturamento Bruto (GMV) Atual"];
  const afiliados = data["Nível de Atividade da Base: Número total de afiliados ativos"];
  const criadores = data["Quantos criadores/afiliados postaram ou linkaram produtos nesta semana? Média"];
  const campanhas = data["Quantas campanhas temos ativas atualmente?"];
  const amostras = data["Quantas amostras foram aprovadas nesta semana?"];
  const top3 = data["Top 3 Produtos Mais Vendidos"];
  const gargalo = data["Qual é o principal gargalo operacional que a agência enfrenta hoje para escalar esses canais?"];
  const acao = data["Qual é a principal ação de incentivo ou campanha desenhada para a próxima semana?"];
  const ofensores = data["Principais Ofensores ou Oportunidades"];
  const analista = data["👤 Identificação - Nome do Analista"];
  const dataReport = data["Data do Report"];

  return (
    <div>
      <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 24 }}>
        <span style={{ fontSize: 36 }}>{config.emoji}</span>
        <div>
          <div style={{ fontSize: 22, fontWeight: 900, color: config.color }}>{config.label}</div>
          <div style={{ fontSize: 13, color: "#888" }}>
            Analista: <b style={{ color: "#444" }}>{analista || "—"}</b> &nbsp;|&nbsp; Data: <b style={{ color: "#444" }}>{formatDate(dataReport)}</b>
          </div>
        </div>
      </div>

      <div style={{ display: "flex", gap: 12, flexWrap: "wrap", marginBottom: 24 }}>
        <MetricCard label="GMV (Faturamento Bruto)" value={formatCurrency(gmv)} accent={STARLIVE_ORANGE} />
        <MetricCard label="Afiliados Ativos" value={afiliados ?? "—"} accent={STARLIVE_PURPLE} />
        <MetricCard label="Criadores Postaram" value={typeof criadores === "number" ? criadores : (criadores || "—")} accent={config.color} />
        <MetricCard label="Campanhas Ativas" value={campanhas ?? "—"} accent={STARLIVE_ORANGE} />
        <MetricCard label="Amostras Aprovadas" value={amostras !== undefined && amostras !== null ? amostras : "—"} accent={STARLIVE_PURPLE} />
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16, marginBottom: 16 }}>
        <InfoBlock title="🏆 Top 3 Produtos Mais Vendidos" content={top3} accent={STARLIVE_ORANGE} />
        <InfoBlock title="⚠️ Principais Ofensores / Oportunidades" content={ofensores} accent={STARLIVE_RED} />
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16 }}>
        <InfoBlock title="🚧 Gargalo Operacional" content={gargalo} accent={STARLIVE_PURPLE} />
        <InfoBlock title="🚀 Ação / Campanha Próxima Semana" content={acao} accent={config.color} />
      </div>
    </div>
  );
}

function InfoBlock({ title, content, accent }) {
  return (
    <div style={{
      background: "#fff",
      borderRadius: 12,
      padding: "16px 20px",
      boxShadow: "0 2px 8px rgba(0,0,0,0.06)",
      borderTop: `3px solid ${accent}`,
    }}>
      <div style={{ fontSize: 13, fontWeight: 700, color: "#555", marginBottom: 8 }}>{title}</div>
      <div style={{ fontSize: 14, color: "#1a1a2e", lineHeight: 1.6, whiteSpace: "pre-wrap" }}>
        {content && content !== "nan" ? content : <span style={{ color: "#bbb", fontStyle: "italic" }}>Não informado</span>}
      </div>
    </div>
  );
}

function parseXLSX(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, { type: "array", cellDates: true });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, { defval: null });
        resolve(rows);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

function matchPlatform(row) {
  const plat = row["🚀 Plataforma  🚀 "] || row["Plataforma"] || "";
  for (const [key, cfg] of Object.entries(PLATFORM_CONFIG)) {
    if (cfg.match(plat)) return key;
  }
  return null;
}

export default function App() {
  const [activeTab, setActiveTab] = useState("Shopee");
  const [platformData, setPlatformData] = useState({
    Shopee: null,
    "TikTok MCN": null,
    "TikTok TAP": null,
    "Mercado Livre": null,
  });
  const [loaded, setLoaded] = useState(false);
  const [dragging, setDragging] = useState(false);
  const [fileName, setFileName] = useState(null);

  const processFile = useCallback(async (file) => {
    try {
      const rows = await parseXLSX(file);
      const newData = { Shopee: null, "TikTok MCN": null, "TikTok TAP": null, "Mercado Livre": null };
      for (const row of rows) {
        const key = matchPlatform(row);
        if (key) newData[key] = row;
      }
      setPlatformData(newData);
      setLoaded(true);
      setFileName(file.name);
      // Set first tab with data
      const firstWithData = Object.keys(newData).find(k => newData[k]);
      if (firstWithData) setActiveTab(firstWithData);
    } catch (e) {
      alert("Erro ao ler a planilha. Verifique o formato do arquivo.");
    }
  }, []);

  const handleFile = (e) => {
    const file = e.target.files[0];
    if (file) processFile(file);
  };

  const handleDrop = (e) => {
    e.preventDefault();
    setDragging(false);
    const file = e.dataTransfer.files[0];
    if (file) processFile(file);
  };

  const tabs = Object.keys(PLATFORM_CONFIG);

  return (
    <div style={{ minHeight: "100vh", background: "#f0f2f8", fontFamily: "'Segoe UI', Arial, sans-serif" }}>
      {/* Header */}
      <div style={{
        background: `linear-gradient(135deg, ${STARLIVE_PURPLE} 0%, #3d1773 60%, #1a1a2e 100%)`,
        padding: "24px 32px",
        display: "flex",
        alignItems: "center",
        gap: 20,
        boxShadow: "0 4px 20px rgba(123,47,190,0.4)",
      }}>
        <StarLiveLogo size={56} />
        <div>
          <div style={{ fontSize: 28, fontWeight: 900, color: "#fff", letterSpacing: -0.5 }}>
            Star<span style={{ color: STARLIVE_ORANGE }}>Live</span>
          </div>
          <div style={{ fontSize: 13, color: "rgba(255,255,255,0.7)", fontWeight: 500 }}>
            Relatório Semanal — Performance de Afiliados
          </div>
        </div>
        <div style={{ marginLeft: "auto", textAlign: "right" }}>
          {fileName && (
            <div style={{ fontSize: 12, color: "rgba(255,255,255,0.6)", marginBottom: 4 }}>
              📄 {fileName}
            </div>
          )}
          <label style={{
            background: STARLIVE_ORANGE,
            color: "#fff",
            padding: "8px 18px",
            borderRadius: 8,
            cursor: "pointer",
            fontSize: 13,
            fontWeight: 700,
            boxShadow: "0 2px 8px rgba(247,148,29,0.4)",
            display: "inline-block",
          }}>
            {loaded ? "📂 Trocar Planilha" : "📂 Carregar Planilha"}
            <input type="file" accept=".xlsx,.xls" onChange={handleFile} style={{ display: "none" }} />
          </label>
        </div>
      </div>

      {/* Drop zone or main content */}
      {!loaded ? (
        <div style={{ padding: 40, maxWidth: 700, margin: "40px auto" }}>
          <div
            onDragOver={(e) => { e.preventDefault(); setDragging(true); }}
            onDragLeave={() => setDragging(false)}
            onDrop={handleDrop}
            style={{
              border: `2.5px dashed ${dragging ? STARLIVE_ORANGE : STARLIVE_PURPLE}`,
              borderRadius: 20,
              background: dragging ? "#fff8f0" : "#fff",
              padding: "60px 40px",
              textAlign: "center",
              cursor: "pointer",
              transition: "all 0.2s",
              boxShadow: dragging ? `0 0 0 4px ${STARLIVE_ORANGE}33` : "0 4px 20px rgba(0,0,0,0.06)",
            }}
            onClick={() => document.querySelector('input[type=file]').click()}
          >
            <div style={{ fontSize: 48, marginBottom: 16 }}>📊</div>
            <div style={{ fontSize: 20, fontWeight: 800, color: STARLIVE_PURPLE, marginBottom: 8 }}>
              Arraste sua planilha aqui
            </div>
            <div style={{ fontSize: 14, color: "#888", marginBottom: 20 }}>
              ou clique para selecionar o arquivo .xlsx
            </div>
            <div style={{ fontSize: 12, color: "#bbb" }}>
              Formato: Formulário de Report Semanal — Performance de Afiliados
            </div>
            <input type="file" accept=".xlsx,.xls" onChange={handleFile} style={{ display: "none" }} />
          </div>

          {/* Preview with sample data */}
          <div style={{ marginTop: 32, background: "#fff", borderRadius: 16, padding: 24, boxShadow: "0 2px 12px rgba(0,0,0,0.06)" }}>
            <div style={{ fontSize: 13, fontWeight: 700, color: "#888", marginBottom: 12, textTransform: "uppercase", letterSpacing: 1 }}>
              📋 Colunas esperadas na planilha
            </div>
            {[
              "🚀 Plataforma (Shopee, TTS MCN, TTS TAP, MELI)",
              "Faturamento Bruto (GMV) Atual",
              "Nível de Atividade da Base: afiliados ativos",
              "Criadores que postaram ou linkaram produtos",
              "Top 3 Produtos Mais Vendidos",
              "Campanhas ativas / Amostras aprovadas",
              "Principais Ofensores ou Oportunidades",
              "Gargalo operacional / Ação próxima semana",
            ].map((col, i) => (
              <div key={i} style={{ fontSize: 13, color: "#555", padding: "5px 0", borderBottom: "1px solid #f0f0f0", display: "flex", alignItems: "center", gap: 8 }}>
                <span style={{ color: STARLIVE_ORANGE, fontWeight: 700 }}>•</span> {col}
              </div>
            ))}
          </div>
        </div>
      ) : (
        <div style={{ maxWidth: 960, margin: "0 auto", padding: "24px 16px" }}>
          {/* Platform Tabs */}
          <div style={{ display: "flex", gap: 8, marginBottom: 24, flexWrap: "wrap" }}>
            {tabs.map((tab) => {
              const cfg = PLATFORM_CONFIG[tab];
              const hasData = !!platformData[tab];
              const isActive = activeTab === tab;
              return (
                <button
                  key={tab}
                  onClick={() => setActiveTab(tab)}
                  style={{
                    padding: "10px 20px",
                    borderRadius: 10,
                    border: `2px solid ${isActive ? cfg.color : "#e0e0e0"}`,
                    background: isActive ? cfg.color : "#fff",
                    color: isActive ? (cfg.color === "#FFE600" ? "#1a1a2e" : "#fff") : (hasData ? cfg.color : "#bbb"),
                    fontWeight: 700,
                    fontSize: 14,
                    cursor: "pointer",
                    transition: "all 0.15s",
                    opacity: hasData ? 1 : 0.5,
                    display: "flex",
                    alignItems: "center",
                    gap: 6,
                  }}
                >
                  <span>{cfg.emoji}</span>
                  {cfg.label}
                  {!hasData && <span style={{ fontSize: 11, opacity: 0.7 }}>(sem dados)</span>}
                </button>
              );
            })}
          </div>

          {/* Content */}
          <div style={{
            background: "#fff",
            borderRadius: 20,
            padding: "28px 32px",
            boxShadow: "0 4px 24px rgba(0,0,0,0.08)",
            border: `1px solid ${PLATFORM_CONFIG[activeTab].border}22`,
          }}>
            <PlatformReport
              data={platformData[activeTab]}
              config={PLATFORM_CONFIG[activeTab]}
            />
          </div>

          {/* Footer summary bar */}
          <div style={{
            marginTop: 24,
            background: `linear-gradient(90deg, ${STARLIVE_PURPLE}15, ${STARLIVE_ORANGE}15)`,
            borderRadius: 12,
            padding: "14px 24px",
            display: "flex",
            gap: 24,
            flexWrap: "wrap",
            alignItems: "center",
            border: `1px solid ${STARLIVE_PURPLE}20`,
          }}>
            <div style={{ fontSize: 13, fontWeight: 700, color: STARLIVE_PURPLE }}>📊 Resumo Geral</div>
            {tabs.filter(t => platformData[t]).map(t => {
              const d = platformData[t];
              const gmv = parseFloat(d["Faturamento Bruto (GMV) Atual"]);
              return (
                <div key={t} style={{ fontSize: 12, color: "#555" }}>
                  <span style={{ fontWeight: 700, color: PLATFORM_CONFIG[t].color }}>{PLATFORM_CONFIG[t].emoji} {t}</span>
                  {" "}{isNaN(gmv) ? "" : `— GMV: ${formatCurrency(gmv)}`}
                </div>
              );
            })}
          </div>
        </div>
      )}
    </div>
  );
}

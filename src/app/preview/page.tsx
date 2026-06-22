"use client";
import { useState } from "react";

const themes = [
  {
    id: "lavender",
    name: "라벤더 그라데이션",
    bg: "linear-gradient(135deg, #ede9fe 0%, #dbeafe 100%)",
    card: "white",
    cardBorder: "#c4b5fd",
    header: "rgba(255,255,255,0.7)",
    primary: "#7c3aed",
    primaryBg: "#ede9fe",
    text: "#1e1b4b",
    muted: "#7c3aed",
    tag: "#f5f3ff",
    tagBorder: "#ddd6fe",
    tagText: "#5b21b6",
    result: "#f5f3ff",
    resultBorder: "#c4b5fd",
  },
  {
    id: "warm",
    name: "웜 크림",
    bg: "linear-gradient(135deg, #faf7f2 0%, #fef3e2 100%)",
    card: "white",
    cardBorder: "#fed7aa",
    header: "rgba(255,255,255,0.85)",
    primary: "#ea580c",
    primaryBg: "#fff7ed",
    text: "#1c1917",
    muted: "#78716c",
    tag: "#fff7ed",
    tagBorder: "#fed7aa",
    tagText: "#9a3412",
    result: "#fff7ed",
    resultBorder: "#fdba74",
  },
  {
    id: "dark",
    name: "다크 모드",
    bg: "#0f172a",
    card: "#1e293b",
    cardBorder: "#334155",
    header: "#1e293b",
    primary: "#6366f1",
    primaryBg: "#1e1b4b",
    text: "#f1f5f9",
    muted: "#94a3b8",
    tag: "#334155",
    tagBorder: "#475569",
    tagText: "#cbd5e1",
    result: "#1e293b",
    resultBorder: "#334155",
  },
  {
    id: "indigo",
    name: "딥 인디고",
    bg: "linear-gradient(160deg, #312e81 0%, #1e1b4b 100%)",
    card: "rgba(255,255,255,0.10)",
    cardBorder: "rgba(255,255,255,0.15)",
    header: "rgba(255,255,255,0.08)",
    primary: "#a5b4fc",
    primaryBg: "rgba(165,180,252,0.15)",
    text: "#e0e7ff",
    muted: "#a5b4fc",
    tag: "rgba(255,255,255,0.08)",
    tagBorder: "rgba(165,180,252,0.3)",
    tagText: "#c7d2fe",
    result: "rgba(255,255,255,0.06)",
    resultBorder: "rgba(165,180,252,0.2)",
  },
];

function MockUI({ t }: { t: typeof themes[0] }) {
  const card: React.CSSProperties = {
    background: t.card,
    border: `1px solid ${t.cardBorder}`,
    borderRadius: 14,
    padding: 16,
    marginBottom: 10,
    backdropFilter: "blur(8px)",
  };
  const label: React.CSSProperties = { fontSize: 10, fontWeight: 700, letterSpacing: ".08em", color: t.muted, textTransform: "uppercase", marginBottom: 10, display: "block" };
  const btn: React.CSSProperties = { display: "inline-block", padding: "7px 18px", borderRadius: 8, background: t.primary, color: "white", fontWeight: 700, fontSize: 13, marginTop: 10 };
  const tag: React.CSSProperties = { display: "inline-block", padding: "3px 10px", borderRadius: 20, fontSize: 12, background: t.tag, border: `1px solid ${t.tagBorder}`, color: t.tagText, margin: "2px" };
  const groupBadge: React.CSSProperties = { display: "inline-block", padding: "2px 10px", borderRadius: 20, background: t.primary, color: "white", fontSize: 11, fontWeight: 700, marginBottom: 6 };
  const groupCard: React.CSSProperties = { background: t.result, border: `1px solid ${t.resultBorder}`, borderRadius: 10, padding: "10px 12px", flex: 1 };

  return (
    <div style={{ background: t.bg, borderRadius: 16, padding: 16, minHeight: 440 }}>
      {/* header */}
      <div style={{ background: t.header, borderRadius: 10, padding: "10px 14px", marginBottom: 12, backdropFilter: "blur(8px)" }}>
        <span style={{ fontWeight: 700, fontSize: 15, color: t.text }}>조편성 프로그램</span>
      </div>

      {/* 명단 카드 */}
      <div style={card}>
        <span style={label}>학생 명단</span>
        <div style={{ background: t.id === "dark" ? "#0f172a" : t.id === "indigo" ? "rgba(0,0,0,0.2)" : "#f8fafc", border: `1px solid ${t.cardBorder}`, borderRadius: 8, padding: "8px 10px", fontSize: 12, color: t.muted, marginBottom: 8, height: 36 }}>
          이름 입력...
        </div>
        <div style={{ display: "flex", gap: 6 }}>
          <div style={{ ...btn, padding: "6px 14px", fontSize: 12, marginTop: 0 }}>추가</div>
          <div style={{ padding: "6px 14px", borderRadius: 8, border: `1px solid ${t.cardBorder}`, color: t.muted, fontSize: 12, background: "transparent" }}>엑셀 업로드</div>
        </div>
        <div style={{ marginTop: 10 }}>
          {["김민지", "박서연", "이지우", "최유나", "한서정"].map(n => (
            <span key={n} style={tag}>{n}</span>
          ))}
        </div>
      </div>

      {/* 설정 카드 */}
      <div style={card}>
        <span style={label}>조편성 설정</span>
        <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 10 }}>
          <div style={{ width: 30, height: 30, borderRadius: "50%", border: `1px solid ${t.cardBorder}`, display: "flex", alignItems: "center", justifyContent: "center", color: t.muted, fontSize: 16 }}>−</div>
          <span style={{ fontSize: 22, fontWeight: 700, color: t.text, minWidth: 50, textAlign: "center" }}>4조</span>
          <div style={{ width: 30, height: 30, borderRadius: "50%", border: `1px solid ${t.cardBorder}`, display: "flex", alignItems: "center", justifyContent: "center", color: t.muted, fontSize: 16 }}>+</div>
          <span style={{ fontSize: 12, color: t.muted }}>→ 조당 약 6명</span>
        </div>
        <div style={{ ...btn, display: "block", textAlign: "center", marginTop: 6, fontSize: 13 }}>조편성 시작</div>
      </div>

      {/* 결과 카드 */}
      <div style={card}>
        <span style={label}>결과</span>
        <div style={{ display: "flex", gap: 8 }}>
          {[1, 2, 3].map(n => (
            <div key={n} style={groupCard}>
              <div style={groupBadge}>{n}조</div>
              <div style={{ fontSize: 12, color: t.text, lineHeight: 1.8 }}>김민지<br />박서연</div>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}

export default function Preview() {
  const [selected, setSelected] = useState<string | null>(null);

  return (
    <div style={{ minHeight: "100vh", background: "#f1f5f9", padding: 24, fontFamily: "'Pretendard','Apple SD Gothic Neo',sans-serif" }}>
      <h1 style={{ fontSize: 18, fontWeight: 700, marginBottom: 6, color: "#0f172a" }}>배경 스타일 미리보기</h1>
      <p style={{ fontSize: 13, color: "#64748b", marginBottom: 24 }}>마음에 드는 스타일을 클릭해서 선택하세요.</p>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 20 }}>
        {themes.map(t => (
          <div
            key={t.id}
            onClick={() => setSelected(t.id)}
            style={{ cursor: "pointer", borderRadius: 18, overflow: "hidden", border: selected === t.id ? "3px solid #4f46e5" : "3px solid transparent", transition: "border-color .2s", boxShadow: selected === t.id ? "0 0 0 2px #c7d2fe" : "0 2px 8px rgba(0,0,0,.08)" }}
          >
            <MockUI t={t} />
            <div style={{ background: selected === t.id ? "#eef2ff" : "white", padding: "10px 14px", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
              <span style={{ fontWeight: 600, fontSize: 13, color: selected === t.id ? "#4f46e5" : "#374151" }}>{t.name}</span>
              {selected === t.id && <span style={{ fontSize: 12, color: "#4f46e5", fontWeight: 700 }}>✓ 선택됨</span>}
            </div>
          </div>
        ))}
      </div>

      {selected && (
        <div style={{ marginTop: 24, padding: "14px 18px", background: "#eef2ff", borderRadius: 12, border: "1px solid #c7d2fe", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <span style={{ fontSize: 14, color: "#3730a3", fontWeight: 600 }}>
            "{themes.find(t => t.id === selected)?.name}" 선택됨
          </span>
          <span style={{ fontSize: 13, color: "#6366f1" }}>클로드에게 이 이름을 알려주세요 →</span>
        </div>
      )}
    </div>
  );
}

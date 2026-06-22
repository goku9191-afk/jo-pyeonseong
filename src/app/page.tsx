"use client";

import { useState, useRef } from "react";
import * as XLSX from "xlsx";

type Student = { id: string; name: string };
type Group = { groupNum: number; students: Student[] };
type ManualPin = { studentId: string; groupNum: number };
type ClassGroup = { label: string; names: string[] };

function uid() {
  return Math.random().toString(36).slice(2, 9);
}
function shuffle<T>(arr: T[]): T[] {
  const a = [...arr];
  for (let i = a.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [a[i], a[j]] = [a[j], a[i]];
  }
  return a;
}
function assignGroups(students: Student[], numGroups: number, pins: ManualPin[]): Group[] {
  const groups: Group[] = Array.from({ length: numGroups }, (_, i) => ({ groupNum: i + 1, students: [] }));
  const pinnedIds = new Set(pins.map((p) => p.studentId));
  for (const pin of pins) {
    const g = groups.find((g) => g.groupNum === pin.groupNum);
    const s = students.find((s) => s.id === pin.studentId);
    if (g && s) g.students.push(s);
  }
  for (const s of shuffle(students.filter((s) => !pinnedIds.has(s.id)))) {
    groups.reduce((a, b) => (a.students.length <= b.students.length ? a : b)).students.push(s);
  }
  return groups;
}

const NOT_NAMES = new Set(["성명", "이름", "학생명", "교사명", "담임명", "비고", "합계", "총점", "평균", "학과", "학년", "학급", "번호", "합격", "성별"]);
const isKoreanName = (v: unknown): boolean =>
  typeof v === "string" && /^[가-힣]{2,4}$/.test((v as string).trim()) && !NOT_NAMES.has((v as string).trim());

function parseSheet(rows: unknown[][]): ClassGroup[] {
  if (!rows.length) return [];

  // 1. 이름 열 찾기: 한글 이름이 가장 많은 열
  const maxCols = Math.max(...rows.map((r) => (r as unknown[]).length), 0);
  const colCounts = new Array(maxCols).fill(0);
  for (const row of rows)
    for (let j = 0; j < (row as unknown[]).length; j++)
      if (isKoreanName((row as unknown[])[j])) colCounts[j]++;

  const bestCount = Math.max(...colCounts);
  if (bestCount < 2) return [];
  const nameCol = colCounts.indexOf(bestCount);

  // 2. 반 열 찾기: 이름 열 왼쪽에서 작은 정수가 반복되는 열
  let classCol = -1;
  for (let j = nameCol - 1; j >= 0; j--) {
    const vals = rows.map((r) => (r as unknown[])[j]).filter((v) => v !== null && v !== undefined && v !== "");
    const nums = vals.filter((v) => typeof v === "number" && Number.isInteger(v) && v >= 1 && v <= 30);
    if (nums.length < 2) continue;
    const uniqueNums = new Set(nums).size;
    const avgPerClass = nums.length / uniqueNums;
    // 반 컬럼: 고유값 2~20개이고 학생이 클래스당 평균 3명 이상
    if (uniqueNums >= 2 && uniqueNums <= 20 && avgPerClass >= 3) {
      classCol = j;
      break;
    }
  }

  // 3. 반 컬럼이 있으면 반별로 분리
  if (classCol >= 0) {
    const classMap = new Map<number, string[]>();
    for (const row of rows) {
      const r = row as unknown[];
      const classVal = r[classCol];
      const name = String(r[nameCol] ?? "").trim();
      if (typeof classVal === "number" && isKoreanName(name)) {
        if (!classMap.has(classVal)) classMap.set(classVal, []);
        classMap.get(classVal)!.push(name);
      }
    }
    const sorted = [...classMap.entries()].sort((a, b) => a[0] - b[0]);
    return sorted.map(([num, names]) => ({ label: `${num}반`, names }));
  }

  // 4. 반 컬럼 없으면 시트 전체를 하나로
  const names = rows.map((r) => String((r as unknown[])[nameCol] ?? "").trim()).filter(isKoreanName);
  return [{ label: "전체", names }];
}

export default function Page() {
  const [students, setStudents] = useState<Student[]>([]);
  const [textInput, setTextInput] = useState("");
  const [classGroups, setClassGroups] = useState<{ sheetName: string; classes: ClassGroup[] }[]>([]);
  const [activeSheet, setActiveSheet] = useState<string | null>(null);
  const [activeClass, setActiveClass] = useState<string | null>(null);
  const [numGroups, setNumGroups] = useState(4);
  const [pins, setPins] = useState<ManualPin[]>([]);
  const [pinStudentId, setPinStudentId] = useState("");
  const [pinGroupNum, setPinGroupNum] = useState(1);
  const [showPins, setShowPins] = useState(false);
  const [result, setResult] = useState<Group[] | null>(null);
  const [copyDone, setCopyDone] = useState(false);
  const fileRef = useRef<HTMLInputElement>(null);

  const addNames = (text: string) => {
    const names = text.split(/[\n,，、\t]+/).map((s) => s.trim()).filter(Boolean);
    if (!names.length) return;
    setStudents((prev) => {
      const existing = new Set(prev.map((s) => s.name));
      return [...prev, ...names.filter((n) => !existing.has(n)).map((n) => ({ id: uid(), name: n }))];
    });
    setTextInput("");
    setResult(null);
  };

  const loadClass = (names: string[]) => {
    setStudents(names.map((n) => ({ id: uid(), name: n })));
    setPins([]);
    setResult(null);
  };

  const handleExcel = (file: File) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const data = new Uint8Array(e.target?.result as ArrayBuffer);
      const wb = XLSX.read(data, { type: "array" });
      const parsed = wb.SheetNames.map((sheetName) => {
        const ws = wb.Sheets[sheetName];
        const rows: unknown[][] = XLSX.utils.sheet_to_json(ws, { header: 1 });
        return { sheetName, classes: parseSheet(rows) };
      }).filter((s) => s.classes.length > 0 && s.classes.some((c) => c.names.length > 0));

      setClassGroups(parsed);
      if (parsed.length > 0) {
        setActiveSheet(parsed[0].sheetName);
        const firstClass = parsed[0].classes[0];
        setActiveClass(firstClass.label);
        loadClass(firstClass.names);
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const removeStudent = (id: string) => {
    setStudents((prev) => prev.filter((s) => s.id !== id));
    setPins((prev) => prev.filter((p) => p.studentId !== id));
    setResult(null);
  };

  const addPin = () => {
    if (!pinStudentId) return;
    setPins((prev) => [...prev.filter((p) => p.studentId !== pinStudentId), { studentId: pinStudentId, groupNum: pinGroupNum }]);
    setPinStudentId("");
  };

  const doAssign = () => {
    if (students.length < 2) return;
    const validPins = pins.filter((p) => p.groupNum <= numGroups && students.some((s) => s.id === p.studentId));
    setResult(assignGroups(students, numGroups, validPins));
  };

  const copyResult = async () => {
    if (!result) return;
    await navigator.clipboard.writeText(
      result.map((g) => `[${g.groupNum}조]\n${g.students.map((s) => s.name).join(", ")}`).join("\n\n")
    );
    setCopyDone(true);
    setTimeout(() => setCopyDone(false), 2000);
  };

  const currentSheet = classGroups.find((s) => s.sheetName === activeSheet);

  // ── 스타일 상수
  const S = {
    card: { background: "rgba(255,255,255,0.75)", border: "1px solid rgba(196,181,253,0.4)", borderRadius: 16, padding: 20, boxShadow: "0 2px 12px rgba(124,58,237,.06)", backdropFilter: "blur(12px)" } as React.CSSProperties,
    primary: "#7c3aed",
    primaryBg: "#ede9fe",
    muted: "#6d6a8a",
    label: { fontSize: 11, fontWeight: 700, letterSpacing: ".06em", color: "#a89fc0", textTransform: "uppercase" } as React.CSSProperties,
  };

  return (
    <div style={{ minHeight: "100vh", background: "linear-gradient(145deg, #ede9fe 0%, #dbeafe 50%, #bae6fd 100%)", color: "#1e1b4b", fontFamily: "'Pretendard','Apple SD Gothic Neo','Noto Sans KR',sans-serif" }}>
      {/* Header */}
      <header style={{ borderBottom: "1px solid rgba(196,181,253,0.35)", background: "rgba(255,255,255,0.55)", backdropFilter: "blur(16px)", padding: "0 24px" }}>
        <div style={{ maxWidth: 600, margin: "0 auto", height: 56, display: "flex", alignItems: "center" }}>
          <span style={{ fontWeight: 700, fontSize: 17, color: "#3b0764" }}>조편성 프로그램</span>
        </div>
      </header>

      <main style={{ maxWidth: 600, margin: "0 auto", padding: "24px 16px 80px" }}>

        {/* ── 학생 명단 ── */}
        <section style={S.card}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
            <span style={S.label}>학생 명단</span>
            {students.length > 0 && (
              <button onClick={() => { setStudents([]); setPins([]); setResult(null); }} style={{ fontSize: 12, color: "#ef4444", background: "none", border: "none", cursor: "pointer" }}>
                전체 삭제
              </button>
            )}
          </div>

          {/* 직접 입력 */}
          <textarea
            value={textInput}
            onChange={(e) => setTextInput(e.target.value)}
            onKeyDown={(e) => { if (e.key === "Enter" && !e.shiftKey && textInput.trim()) { e.preventDefault(); addNames(textInput); } }}
            placeholder={"이름 입력 (Enter 또는 쉼표·줄바꿈으로 구분)\n예) 김민지, 박서연, 이지우"}
            rows={4}
            style={{ width: "100%", boxSizing: "border-box", border: "1px solid rgba(196,181,253,0.5)", borderRadius: 10, padding: "12px 14px", fontSize: 14, outline: "none", resize: "none", color: "#1e1b4b", background: "rgba(255,255,255,0.5)", fontFamily: "inherit" }}
          />
          <div style={{ display: "flex", gap: 8, marginTop: 8 }}>
            <button
              onClick={() => addNames(textInput)}
              disabled={!textInput.trim()}
              style={{ flex: 1, padding: "10px 0", borderRadius: 10, border: "none", background: S.primary, color: "white", fontWeight: 600, fontSize: 14, cursor: "pointer", opacity: textInput.trim() ? 1 : 0.4 }}
            >
              추가
            </button>
            <button
              onClick={() => fileRef.current?.click()}
              style={{ flex: 1, padding: "10px 0", borderRadius: 10, border: "1px solid rgba(196,181,253,0.5)", background: "rgba(255,255,255,0.6)", color: S.muted, fontWeight: 600, fontSize: 14, cursor: "pointer" }}
            >
              엑셀 업로드
            </button>
            <input ref={fileRef} type="file" accept=".xlsx,.xls" style={{ display: "none" }} onChange={(e) => { const f = e.target.files?.[0]; if (f) handleExcel(f); e.target.value = ""; }} />
          </div>

          {/* 시트 & 반 선택 */}
          {classGroups.length > 0 && (
            <div style={{ marginTop: 16, paddingTop: 16, borderTop: "1px solid rgba(196,181,253,0.3)" }}>
              {/* 시트 탭 */}
              {classGroups.length > 1 && (
                <div style={{ marginBottom: 10 }}>
                  <div style={S.label}>학년 선택</div>
                  <div style={{ display: "flex", flexWrap: "wrap", gap: 6, marginTop: 6 }}>
                    {classGroups.map((sg) => (
                      <button key={sg.sheetName} onClick={() => setActiveSheet(sg.sheetName)} style={{ padding: "5px 14px", borderRadius: 20, border: "1px solid", fontSize: 13, fontWeight: 600, cursor: "pointer", borderColor: activeSheet === sg.sheetName ? S.primary : "#e2e8f0", background: activeSheet === sg.sheetName ? S.primaryBg : "white", color: activeSheet === sg.sheetName ? S.primary : S.muted }}>
                        {sg.sheetName}
                      </button>
                    ))}
                  </div>
                </div>
              )}

              {/* 반 탭 */}
              {currentSheet && currentSheet.classes.length > 1 && (
                <div>
                  <div style={S.label}>반 선택</div>
                  <div style={{ display: "flex", flexWrap: "wrap", gap: 6, marginTop: 6 }}>
                    {currentSheet.classes.map((cls) => (
                      <button
                        key={cls.label}
                        onClick={() => { setActiveClass(cls.label); loadClass(cls.names); }}
                        style={{ padding: "5px 14px", borderRadius: 20, border: "1px solid", fontSize: 13, fontWeight: 600, cursor: "pointer", borderColor: activeClass === cls.label ? S.primary : "#e2e8f0", background: activeClass === cls.label ? S.primaryBg : "white", color: activeClass === cls.label ? S.primary : S.muted }}
                      >
                        {cls.label} <span style={{ fontWeight: 400, opacity: 0.7 }}>({cls.names.length}명)</span>
                      </button>
                    ))}
                  </div>
                </div>
              )}
            </div>
          )}

          {/* 학생 태그 */}
          {students.length > 0 && (
            <div style={{ marginTop: 14 }}>
              <span style={{ ...S.label, marginBottom: 8, display: "block" }}>현재 {students.length}명</span>
              <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
                {students.map((s) => (
                  <span key={s.id} style={{ display: "flex", alignItems: "center", gap: 4, padding: "4px 10px", borderRadius: 20, fontSize: 13, background: "rgba(237,233,254,0.6)", color: "#3b0764" }}>
                    {s.name}
                    <button onClick={() => removeStudent(s.id)} style={{ background: "none", border: "none", cursor: "pointer", color: "#94a3b8", fontSize: 11, padding: 0, lineHeight: 1 }}>✕</button>
                  </span>
                ))}
              </div>
            </div>
          )}
        </section>

        {/* ── 조편성 설정 ── */}
        <section style={{ ...S.card, marginTop: 12 }}>
          <div style={{ ...S.label, marginBottom: 16 }}>조편성 설정</div>

          {/* 조 수 */}
          <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 4 }}>
            <button onClick={() => setNumGroups((n) => Math.max(2, n - 1))} style={{ width: 36, height: 36, borderRadius: "50%", border: "1px solid rgba(196,181,253,0.5)", background: "rgba(255,255,255,0.6)", fontSize: 20, cursor: "pointer", color: S.muted, display: "flex", alignItems: "center", justifyContent: "center" }}>−</button>
            <span style={{ fontSize: 28, fontWeight: 700, minWidth: 80, textAlign: "center", color: "#1e1b4b" }}>{numGroups}조</span>
            <button onClick={() => setNumGroups((n) => (students.length ? Math.min(students.length, n + 1) : n + 1))} style={{ width: 36, height: 36, borderRadius: "50%", border: "1px solid rgba(196,181,253,0.5)", background: "rgba(255,255,255,0.6)", fontSize: 20, cursor: "pointer", color: S.muted, display: "flex", alignItems: "center", justifyContent: "center" }}>+</button>
            {students.length > 0 && (
              <span style={{ fontSize: 13, color: S.muted }}>→ 조당 약 {Math.ceil(students.length / numGroups)}명</span>
            )}
          </div>

          {/* 수동 배정 토글 */}
          <button onClick={() => setShowPins((v) => !v)} style={{ marginTop: 14, fontSize: 13, color: S.muted, background: "none", border: "none", cursor: "pointer", display: "flex", alignItems: "center", gap: 4, padding: 0 }}>
            <span style={{ fontSize: 10 }}>{showPins ? "▲" : "▼"}</span> 수동 배정 (특정 학생 고정)
          </button>

          {showPins && (
            <div style={{ marginTop: 12, paddingTop: 12, borderTop: "1px solid rgba(196,181,253,0.3)" }}>
              {students.length === 0 ? (
                <span style={{ fontSize: 13, color: "#94a3b8" }}>학생을 먼저 추가해 주세요.</span>
              ) : (
                <>
                  <div style={{ display: "flex", gap: 8 }}>
                    <select value={pinStudentId} onChange={(e) => setPinStudentId(e.target.value)} style={{ flex: 1, padding: "9px 12px", borderRadius: 10, border: "1px solid rgba(196,181,253,0.5)", fontSize: 13, color: "#1e1b4b", outline: "none", background: "rgba(255,255,255,0.6)" }}>
                      <option value="">학생 선택</option>
                      {students.map((s) => {
                        const pin = pins.find((p) => p.studentId === s.id);
                        return <option key={s.id} value={s.id}>{s.name}{pin ? ` (${pin.groupNum}조 고정)` : ""}</option>;
                      })}
                    </select>
                    <select value={pinGroupNum} onChange={(e) => setPinGroupNum(Number(e.target.value))} style={{ width: 80, padding: "9px 12px", borderRadius: 10, border: "1px solid rgba(196,181,253,0.5)", fontSize: 13, color: "#1e1b4b", outline: "none", background: "rgba(255,255,255,0.6)" }}>
                      {Array.from({ length: numGroups }, (_, i) => <option key={i + 1} value={i + 1}>{i + 1}조</option>)}
                    </select>
                    <button onClick={addPin} style={{ padding: "9px 16px", borderRadius: 10, border: "none", background: S.primary, color: "white", fontWeight: 600, fontSize: 13, cursor: "pointer" }}>고정</button>
                  </div>
                  {pins.length > 0 && (
                    <div style={{ display: "flex", flexWrap: "wrap", gap: 6, marginTop: 10 }}>
                      {pins.map((p) => {
                        const s = students.find((s) => s.id === p.studentId);
                        if (!s) return null;
                        return (
                          <span key={p.studentId} style={{ display: "flex", alignItems: "center", gap: 4, padding: "4px 10px", borderRadius: 20, fontSize: 12, background: "#eef2ff", color: S.primary, border: "1px solid #c7d2fe" }}>
                            📌 {s.name} → {p.groupNum}조
                            <button onClick={() => setPins((prev) => prev.filter((x) => x.studentId !== p.studentId))} style={{ background: "none", border: "none", cursor: "pointer", color: S.primary, fontSize: 11, padding: 0, opacity: 0.6 }}>✕</button>
                          </span>
                        );
                      })}
                    </div>
                  )}
                </>
              )}
            </div>
          )}

          {/* 배정 버튼 */}
          <button
            onClick={doAssign}
            disabled={students.length < 2}
            style={{ marginTop: 20, width: "100%", padding: "14px 0", borderRadius: 12, border: "none", background: students.length >= 2 ? S.primary : "#e2e8f0", color: students.length >= 2 ? "white" : "#94a3b8", fontWeight: 700, fontSize: 15, cursor: students.length >= 2 ? "pointer" : "default", letterSpacing: ".02em" }}
          >
            조편성 시작
          </button>
        </section>

        {/* ── 결과 ── */}
        {result && (
          <section style={{ ...S.card, marginTop: 12 }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
              <span style={S.label}>결과</span>
              <div style={{ display: "flex", gap: 8 }}>
                <button onClick={copyResult} style={{ padding: "6px 14px", borderRadius: 8, border: "1px solid #e2e8f0", background: copyDone ? "#dcfce7" : "white", color: copyDone ? "#16a34a" : S.muted, fontWeight: 600, fontSize: 13, cursor: "pointer" }}>
                  {copyDone ? "복사됨 ✓" : "복사"}
                </button>
                <button onClick={() => window.print()} style={{ padding: "6px 14px", borderRadius: 8, border: "1px solid rgba(196,181,253,0.5)", background: "rgba(255,255,255,0.6)", color: S.muted, fontWeight: 600, fontSize: 13, cursor: "pointer" }}>
                  인쇄
                </button>
              </div>
            </div>

            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
              {result.map((g) => (
                <div key={g.groupNum} style={{ background: "rgba(255,255,255,0.45)", border: "1px solid rgba(196,181,253,0.4)", borderRadius: 12, padding: 14 }}>
                  <div style={{ display: "inline-block", padding: "2px 12px", borderRadius: 20, background: S.primary, color: "white", fontSize: 12, fontWeight: 700, marginBottom: 10 }}>
                    {g.groupNum}조
                  </div>
                  <ul style={{ listStyle: "none", padding: 0, margin: 0, display: "flex", flexDirection: "column", gap: 4 }}>
                    {g.students.map((s) => (
                      <li key={s.id} style={{ fontSize: 14, color: "#2e1065", display: "flex", alignItems: "center", gap: 4 }}>
                        {pins.some((p) => p.studentId === s.id) && <span style={{ fontSize: 11 }}>📌</span>}
                        {s.name}
                      </li>
                    ))}
                  </ul>
                </div>
              ))}
            </div>

            <div style={{ display: "flex", gap: 8, marginTop: 12 }}>
              <button onClick={doAssign} style={{ flex: 1, padding: "11px 0", borderRadius: 10, border: "none", background: "#0f172a", color: "white", fontWeight: 600, fontSize: 14, cursor: "pointer" }}>
                다시 뽑기
              </button>
              <button onClick={() => setResult(null)} style={{ flex: 1, padding: "11px 0", borderRadius: 10, border: "1px solid rgba(196,181,253,0.5)", background: "rgba(255,255,255,0.6)", color: S.muted, fontWeight: 600, fontSize: 14, cursor: "pointer" }}>
                설정 변경
              </button>
            </div>
          </section>
        )}
      </main>
    </div>
  );
}

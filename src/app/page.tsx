"use client";

import { useState, useRef, useEffect, useLayoutEffect } from "react";
import * as XLSX from "xlsx";

type Student = { id: string; name: string };
type Group = { groupNum: number; students: Student[] };
type ManualPin = { studentId: string; groupNum: number };
type ClassGroup = { label: string; names: string[] };
type DragState = {
  id: string; chip: HTMLElement; x0: number; y0: number; ox: number; oy: number; w: number;
  started: boolean; ghost: HTMLElement | null; group: number | null; before: string | null;
};

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

// 결과 수정: 학생을 다른 조(또는 특정 학생 앞)로 옮긴다. 변화가 없으면 null
function moveIn(groups: Group[], id: string, toGroup: number, beforeId: string | null): Group[] | null {
  if (beforeId === id) return null;
  const next = groups.map((g) => ({ ...g, students: [...g.students] }));
  const from = next.find((g) => g.students.some((s) => s.id === id));
  const to = next.find((g) => g.groupNum === toGroup);
  if (!from || !to) return null;
  const cur = from.students.findIndex((s) => s.id === id);
  if (from === to && !beforeId && cur === from.students.length - 1) return null;
  if (from === to && beforeId && from.students[cur + 1]?.id === beforeId) return null;
  const [stu] = from.students.splice(cur, 1);
  let idx = beforeId ? to.students.findIndex((s) => s.id === beforeId) : -1;
  if (idx < 0) idx = to.students.length;
  to.students.splice(idx, 0, stu);
  return next;
}
// 결과 수정: 서로 다른 조의 두 학생 자리를 바꾼다
function swapIn(groups: Group[], a: string, b: string): Group[] | null {
  const next = groups.map((g) => ({ ...g, students: [...g.students] }));
  const ga = next.find((g) => g.students.some((s) => s.id === a));
  const gb = next.find((g) => g.students.some((s) => s.id === b));
  if (!ga || !gb || ga === gb) return null;
  const ia = ga.students.findIndex((s) => s.id === a);
  const ib = gb.students.findIndex((s) => s.id === b);
  [ga.students[ia], gb.students[ib]] = [gb.students[ib], ga.students[ia]];
  return next;
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
  const [history, setHistory] = useState<Group[][]>([]);
  const [selected, setSelected] = useState<string | null>(null);
  const [focus, setFocus] = useState(false);
  const [listFolded, setListFolded] = useState(false);
  const [redoArmed, setRedoArmed] = useState(false);
  const fileRef = useRef<HTMLInputElement>(null);
  const groupsRef = useRef<HTMLDivElement>(null);
  const dragRef = useRef<DragState | null>(null);

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
    setHistory([]);
    setSelected(null);
    setListFolded(true);
  };

  // 다시 뽑기: 손으로 옮긴 게 있으면 한 번 더 눌러야 실행
  const redo = () => {
    if (history.length && !redoArmed) {
      setRedoArmed(true);
      setTimeout(() => setRedoArmed(false), 3000);
      return;
    }
    setRedoArmed(false);
    doAssign();
  };

  const commit = (next: Group[] | null) => {
    setSelected(null);
    if (!next || !result) return;
    setHistory((h) => [...h.slice(-49), result]);
    setResult(next);
  };

  const undo = () => {
    if (!history.length) return;
    setResult(history[history.length - 1]);
    setHistory((h) => h.slice(0, -1));
    setSelected(null);
  };

  const copyResult = async () => {
    if (!result) return;
    await navigator.clipboard.writeText(
      result.map((g) => `[${g.groupNum}조]\n${g.students.map((s) => s.name).join(", ")}`).join("\n\n")
    );
    setCopyDone(true);
    setTimeout(() => setCopyDone(false), 2000);
  };

  // ── 결과판 드래그 & 탭 (마우스·터치 공용 pointer 이벤트)
  const onChipPointerDown = (e: React.PointerEvent<HTMLDivElement>, id: string) => {
    if (e.pointerType === "mouse" && e.button !== 0) return;
    e.preventDefault();
    const chip = e.currentTarget;
    const r = chip.getBoundingClientRect();
    dragRef.current = { id, chip, x0: e.clientX, y0: e.clientY, ox: e.clientX - r.left, oy: e.clientY - r.top, w: r.width, started: false, ghost: null, group: null, before: null };
  };

  useEffect(() => {
    const clearMarks = () => {
      groupsRef.current?.querySelectorAll(".over").forEach((el) => el.classList.remove("over"));
      groupsRef.current?.querySelectorAll(".before").forEach((el) => el.classList.remove("before"));
    };
    const onMove = (e: PointerEvent) => {
      const d = dragRef.current;
      if (!d) return;
      if (!d.started) {
        if (Math.hypot(e.clientX - d.x0, e.clientY - d.y0) < 6) return;
        d.started = true;
        const g = d.chip.cloneNode(true) as HTMLElement;
        g.classList.add("ghost");
        g.classList.remove("sel");
        g.style.width = d.w + "px";
        document.body.appendChild(g);
        d.ghost = g;
        d.chip.classList.add("dragging");
      }
      d.ghost!.style.transform = `translate(${e.clientX - d.ox}px, ${e.clientY - d.oy}px)`;
      if (e.clientY < 70) window.scrollBy(0, -14);
      else if (e.clientY > window.innerHeight - 70) window.scrollBy(0, 14);
      const el = document.elementFromPoint(e.clientX, e.clientY);
      const grp = el?.closest<HTMLElement>(".group");
      const over = el?.closest<HTMLElement>(".mchip");
      clearMarks();
      d.group = grp ? Number(grp.dataset.group) : null;
      d.before = over && over !== d.chip && grp ? over.dataset.id ?? null : null;
      grp?.classList.add("over");
      if (d.before) over!.classList.add("before");
    };
    const end = (cancelled: boolean) => {
      const d = dragRef.current;
      if (!d) return;
      dragRef.current = null;
      d.ghost?.remove();
      d.chip.classList.remove("dragging");
      clearMarks();
      if (cancelled || !result) return;
      if (d.started) {
        if (d.group) commit(moveIn(result, d.id, d.group, d.before));
        return;
      }
      // 탭: 선택 → 다른 조 학생을 누르면 자리 바꾸기
      if (!selected || selected === d.id) { setSelected(selected === d.id ? null : d.id); return; }
      const swapped = swapIn(result, selected, d.id);
      if (swapped) commit(swapped);
      else setSelected(d.id);
    };
    const onUp = () => end(false);
    const onCancel = () => end(true);
    const onKey = (e: KeyboardEvent) => { if (e.key === "Escape") setSelected(null); };
    document.addEventListener("pointermove", onMove);
    document.addEventListener("pointerup", onUp);
    document.addEventListener("pointercancel", onCancel);
    document.addEventListener("keydown", onKey);
    return () => {
      document.removeEventListener("pointermove", onMove);
      document.removeEventListener("pointerup", onUp);
      document.removeEventListener("pointercancel", onCancel);
      document.removeEventListener("keydown", onKey);
    };
  });

  // ── 결과만 크게 보기: 스크롤 없이 한 화면에 꽉 차도록 칸 수와 글자 크기를 맞춘다
  useLayoutEffect(() => {
    const on = focus && !!result;
    document.documentElement.classList.toggle("fs", on);
    const g = groupsRef.current;
    if (!g) return;
    if (!on || !result) {
      ["--cols", "--rows", "--fs"].forEach((k) => g.style.removeProperty(k));
      return;
    }
    const fit = () => {
      const n = result.length, gap = 12;
      const W = g.clientWidth, H = g.clientHeight;
      const maxN = Math.max(1, ...result.map((x) => x.students.length));
      const maxLen = Math.max(2, ...result.flatMap((x) => x.students.map((s) => s.name.length)));
      let best = { cols: 1, rows: n, f: -Infinity };
      for (let cols = 1; cols <= n; cols++) {
        const rows = Math.ceil(n / cols);
        const cw = (W - gap * (cols - 1)) / cols, ch = (H - gap * (rows - 1)) / rows;
        const f = Math.min((ch - 70) / (maxN * 1.75), (cw - 32) / (maxLen + 2.6));
        if (f > best.f) best = { cols, rows, f };
      }
      g.style.setProperty("--cols", String(best.cols));
      g.style.setProperty("--rows", String(best.rows));
      let fs = Math.max(10, Math.min(72, Math.floor(best.f)));
      const overflows = () => [...g.querySelectorAll<HTMLElement>(".group")].some(
        (el) => el.scrollHeight > el.clientHeight + 1 || [...el.querySelectorAll<HTMLElement>(".mchip")].some((c) => c.scrollWidth > c.clientWidth + 1)
      );
      g.style.setProperty("--fs", fs + "px");
      while (fs > 10 && overflows()) { fs--; g.style.setProperty("--fs", fs + "px"); }
    };
    fit();
    window.addEventListener("resize", fit);
    return () => window.removeEventListener("resize", fit);
  }, [focus, result]);

  useEffect(() => () => document.documentElement.classList.remove("fs"), []);

  const currentSheet = classGroups.find((s) => s.sheetName === activeSheet);
  const hasResult = !!result;
  const c = hasResult; // 조편성 후에는 명단·설정을 작게
  const folded = hasResult && listFolded;

  // ── 스타일 상수
  const S = {
    card: { background: "rgba(255,255,255,0.75)", border: "1px solid rgba(196,181,253,0.4)", borderRadius: 16, padding: c ? 14 : 20, boxShadow: "0 2px 12px rgba(124,58,237,.06)", backdropFilter: "blur(12px)", minWidth: 0 } as React.CSSProperties,
    primary: "#7c3aed",
    primaryBg: "#ede9fe",
    muted: "#6d6a8a",
    label: { fontSize: 11, fontWeight: 700, letterSpacing: ".06em", color: "#a89fc0", textTransform: "uppercase" } as React.CSSProperties,
  };

  return (
    <div className="page" style={{ minHeight: "100vh", background: "linear-gradient(145deg, #ede9fe 0%, #dbeafe 50%, #bae6fd 100%)", color: "#1e1b4b", fontFamily: "'Pretendard','Apple SD Gothic Neo','Noto Sans KR',sans-serif" }}>
      {/* Header */}
      <header className="top" style={{ borderBottom: "1px solid rgba(196,181,253,0.35)", background: "rgba(255,255,255,0.55)", backdropFilter: "blur(16px)", padding: "0 24px" }}>
        <div style={{ maxWidth: hasResult ? 1400 : 600, margin: "0 auto", height: 56, display: "flex", alignItems: "center", justifyContent: "space-between" }}>
          <span style={{ fontWeight: 700, fontSize: 17, color: "#3b0764" }}>조편성 프로그램</span>
          {hasResult && (
            <button onClick={() => setFocus((v) => !v)} className="no-print" style={{ padding: "7px 14px", borderRadius: 10, border: "1px solid rgba(196,181,253,0.5)", background: "rgba(255,255,255,0.7)", color: S.muted, fontWeight: 600, fontSize: 13, cursor: "pointer" }}>
              {focus ? "명단·설정 보기" : "결과만 크게 보기"}
            </button>
          )}
        </div>
      </header>

      <main className={`app${hasResult ? " has-result" : ""}${hasResult && focus ? " focus" : ""}`}>
        <div className="side">
          {/* ── 학생 명단 ── */}
          <section style={S.card}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: folded ? 0 : 14 }}>
              <span style={S.label}>{folded ? `학생 명단 · ${students.length}명` : "학생 명단"}</span>
              <span style={{ display: "flex", gap: 10 }}>
                {students.length > 0 && !folded && (
                  <button onClick={() => { setStudents([]); setPins([]); setResult(null); }} style={{ fontSize: 12, color: "#ef4444", background: "none", border: "none", cursor: "pointer" }}>
                    전체 삭제
                  </button>
                )}
                {hasResult && (
                  <button onClick={() => setListFolded((v) => !v)} style={{ fontSize: 12, color: S.muted, background: "none", border: "none", cursor: "pointer" }}>
                    {folded ? "펼치기" : "접기"}
                  </button>
                )}
              </span>
            </div>

            {!folded && (
              <>
                {/* 직접 입력 */}
                <textarea
                  value={textInput}
                  onChange={(e) => setTextInput(e.target.value)}
                  onKeyDown={(e) => { if (e.key === "Enter" && !e.shiftKey && !e.nativeEvent.isComposing && textInput.trim()) { e.preventDefault(); addNames(textInput); } }}
                  placeholder={"이름 입력 (Enter 또는 쉼표·줄바꿈으로 구분)\n예) 김민지, 박서연, 이지우"}
                  rows={c ? 2 : 4}
                  style={{ width: "100%", boxSizing: "border-box", border: "1px solid rgba(196,181,253,0.5)", borderRadius: 10, padding: c ? "8px 10px" : "12px 14px", fontSize: c ? 13 : 14, outline: "none", resize: "none", color: "#1e1b4b", background: "rgba(255,255,255,0.5)", fontFamily: "inherit" }}
                />
                <div style={{ display: "flex", gap: 8, marginTop: 8 }}>
                  <button
                    onClick={() => addNames(textInput)}
                    disabled={!textInput.trim()}
                    style={{ flex: 1, padding: c ? "7px 0" : "10px 0", borderRadius: 10, border: "none", background: S.primary, color: "white", fontWeight: 600, fontSize: c ? 13 : 14, cursor: "pointer", opacity: textInput.trim() ? 1 : 0.4 }}
                  >
                    추가
                  </button>
                  <button
                    onClick={() => fileRef.current?.click()}
                    style={{ flex: 1, padding: c ? "7px 0" : "10px 0", borderRadius: 10, border: "1px solid rgba(196,181,253,0.5)", background: "rgba(255,255,255,0.6)", color: S.muted, fontWeight: 600, fontSize: c ? 13 : 14, cursor: "pointer" }}
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
                    <div style={{ display: "flex", flexWrap: "wrap", gap: 6, maxHeight: c ? 180 : undefined, overflowY: c ? "auto" : undefined }}>
                      {students.map((s) => (
                        <span key={s.id} style={{ display: "flex", alignItems: "center", gap: 4, padding: c ? "2px 8px" : "4px 10px", borderRadius: 20, fontSize: c ? 12 : 13, background: "rgba(237,233,254,0.6)", color: "#3b0764" }}>
                          {s.name}
                          <button onClick={() => removeStudent(s.id)} style={{ background: "none", border: "none", cursor: "pointer", color: "#94a3b8", fontSize: 11, padding: 0, lineHeight: 1 }}>✕</button>
                        </span>
                      ))}
                    </div>
                  </div>
                )}
              </>
            )}
          </section>

          {/* ── 조편성 설정 ── */}
          <section style={S.card}>
            <div style={{ ...S.label, marginBottom: c ? 12 : 16 }}>조편성 설정</div>

            {/* 조 수 */}
            <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 4, flexWrap: "wrap" }}>
              <button onClick={() => setNumGroups((n) => Math.max(2, n - 1))} style={{ width: c ? 30 : 36, height: c ? 30 : 36, borderRadius: "50%", border: "1px solid rgba(196,181,253,0.5)", background: "rgba(255,255,255,0.6)", fontSize: c ? 17 : 20, cursor: "pointer", color: S.muted, display: "flex", alignItems: "center", justifyContent: "center" }}>−</button>
              <span style={{ fontSize: c ? 22 : 28, fontWeight: 700, minWidth: c ? 60 : 80, textAlign: "center", color: "#1e1b4b" }}>{numGroups}조</span>
              <button onClick={() => setNumGroups((n) => (students.length ? Math.min(students.length, n + 1) : n + 1))} style={{ width: c ? 30 : 36, height: c ? 30 : 36, borderRadius: "50%", border: "1px solid rgba(196,181,253,0.5)", background: "rgba(255,255,255,0.6)", fontSize: c ? 17 : 20, cursor: "pointer", color: S.muted, display: "flex", alignItems: "center", justifyContent: "center" }}>+</button>
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
                    <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                      <select value={pinStudentId} onChange={(e) => setPinStudentId(e.target.value)} style={{ flex: 1, minWidth: 110, padding: "9px 12px", borderRadius: 10, border: "1px solid rgba(196,181,253,0.5)", fontSize: 13, color: "#1e1b4b", outline: "none", background: "rgba(255,255,255,0.6)" }}>
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
              style={{ marginTop: c ? 16 : 20, width: "100%", padding: c ? "10px 0" : "14px 0", borderRadius: 12, border: "none", background: students.length >= 2 ? S.primary : "#e2e8f0", color: students.length >= 2 ? "white" : "#94a3b8", fontWeight: 700, fontSize: c ? 14 : 15, cursor: students.length >= 2 ? "pointer" : "default", letterSpacing: ".02em" }}
            >
              {hasResult ? "새로 조편성" : "조편성 시작"}
            </button>
          </section>
        </div>

        {/* ── 결과 ── */}
        {result && (
          <section className="result">
            <div className="result-head">
              <span className="result-title">조편성 결과</span>
              <div className="no-print" style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                <button onClick={undo} disabled={!history.length} className="rbtn">
                  {history.length ? `되돌리기 (${history.length})` : "되돌리기"}
                </button>
                <button onClick={copyResult} className="rbtn" style={{ background: copyDone ? "#dcfce7" : undefined, color: copyDone ? "#16a34a" : undefined }}>
                  {copyDone ? "복사됨 ✓" : "복사"}
                </button>
                <button onClick={() => window.print()} className="rbtn">
                  인쇄
                </button>
              </div>
            </div>
            <p className="hint no-print">이름을 끌어서 다른 조에 놓으면 옮겨집니다. 이름을 눌러 선택한 뒤 다른 조를 누르면 이동, 다른 조 학생을 누르면 서로 자리를 바꿉니다.</p>

            <div className="groups" ref={groupsRef}>
              {result.map((g) => (
                <div
                  key={g.groupNum}
                  data-group={g.groupNum}
                  className={`group${selected && !g.students.some((s) => s.id === selected) ? " target" : ""}`}
                  onClick={(e) => {
                    if (!selected || (e.target as HTMLElement).closest(".mchip")) return;
                    commit(moveIn(result, selected, g.groupNum, null));
                  }}
                >
                  <div className="ghead">
                    <span className="gbadge">{g.groupNum}조</span>
                    <span className="gcount">{g.students.length}명</span>
                  </div>
                  <div className="glist">
                    {g.students.length ? g.students.map((s) => (
                      <div
                        key={s.id}
                        data-id={s.id}
                        className={`mchip${selected === s.id ? " sel" : ""}`}
                        onPointerDown={(e) => onChipPointerDown(e, s.id)}
                      >
                        {pins.some((p) => p.studentId === s.id) && <span className="pin">📌</span>}
                        {s.name}
                        <span className="grip" aria-hidden="true">⋮⋮</span>
                      </div>
                    )) : <div className="empty">여기로 끌어다 놓기</div>}
                  </div>
                </div>
              ))}
            </div>

            <div className="actions no-print">
              <button onClick={redo} style={{ flex: 1, padding: "11px 0", borderRadius: 10, border: "none", background: "#0f172a", color: "white", fontWeight: 600, fontSize: 14, cursor: "pointer" }}>
                {redoArmed ? "옮긴 내용이 사라져요. 한 번 더 누르기" : "다시 뽑기"}
              </button>
              <button onClick={() => { setResult(null); setFocus(false); setHistory([]); setSelected(null); }} style={{ flex: 1, padding: "11px 0", borderRadius: 10, border: "1px solid rgba(196,181,253,0.5)", background: "rgba(255,255,255,0.6)", color: S.muted, fontWeight: 600, fontSize: 14, cursor: "pointer" }}>
                설정 변경
              </button>
            </div>
          </section>
        )}
      </main>
    </div>
  );
}

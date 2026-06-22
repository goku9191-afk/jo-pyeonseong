"use client";

import { useState, useRef, useCallback } from "react";
import * as XLSX from "xlsx";

type Student = { id: string; name: string };
type ClassRoster = { id: string; className: string; students: Student[] };
type Group = { groupNum: number; students: Student[] };
type ManualPin = { studentId: string; groupNum: number };

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
  const groups: Group[] = Array.from({ length: numGroups }, (_, i) => ({
    groupNum: i + 1,
    students: [],
  }));

  const pinnedIds = new Set(pins.map((p) => p.studentId));
  for (const pin of pins) {
    const g = groups.find((g) => g.groupNum === pin.groupNum);
    const student = students.find((s) => s.id === pin.studentId);
    if (g && student) g.students.push(student);
  }

  const free = shuffle(students.filter((s) => !pinnedIds.has(s.id)));
  for (const student of free) {
    const smallest = groups.reduce((a, b) => (a.students.length <= b.students.length ? a : b));
    smallest.students.push(student);
  }

  return groups;
}

const TAB_LABELS = ["📋 명렬 관리", "🎲 조편성 설정", "✨ 결과 보기"];
type Tab = 0 | 1 | 2;

export default function Page() {
  const [tab, setTab] = useState<Tab>(0);
  const [rosters, setRosters] = useState<ClassRoster[]>([]);
  const [selectedRosterId, setSelectedRosterId] = useState<string | null>(null);
  const [newClassName, setNewClassName] = useState("");
  const [directInput, setDirectInput] = useState("");
  const [numGroups, setNumGroups] = useState(4);
  const [pins, setPins] = useState<ManualPin[]>([]);
  const [result, setResult] = useState<Group[] | null>(null);
  const [pinStudentId, setPinStudentId] = useState("");
  const [pinGroupNum, setPinGroupNum] = useState(1);
  const [copyDone, setCopyDone] = useState(false);
  const fileRef = useRef<HTMLInputElement>(null);

  const selectedRoster = rosters.find((r) => r.id === selectedRosterId) ?? null;

  /* ─── 명렬 관리 ─── */
  const addRoster = () => {
    const name = newClassName.trim();
    if (!name) return;
    const id = uid();
    setRosters((prev) => [...prev, { id, className: name, students: [] }]);
    setSelectedRosterId(id);
    setNewClassName("");
  };

  const deleteRoster = (id: string) => {
    setRosters((prev) => prev.filter((r) => r.id !== id));
    if (selectedRosterId === id) {
      setSelectedRosterId(rosters.find((r) => r.id !== id)?.id ?? null);
    }
  };

  const addStudentsFromText = useCallback(
    (text: string, rosterId: string) => {
      const names = text
        .split(/[\n,，、\t]+/)
        .map((s) => s.trim())
        .filter(Boolean);
      if (!names.length) return;
      setRosters((prev) =>
        prev.map((r) =>
          r.id === rosterId
            ? {
                ...r,
                students: [
                  ...r.students,
                  ...names
                    .filter((n) => !r.students.some((s) => s.name === n))
                    .map((n) => ({ id: uid(), name: n })),
                ],
              }
            : r
        )
      );
      setDirectInput("");
    },
    []
  );

  const removeStudent = (rosterId: string, studentId: string) => {
    setRosters((prev) =>
      prev.map((r) =>
        r.id === rosterId ? { ...r, students: r.students.filter((s) => s.id !== studentId) } : r
      )
    );
    setPins((prev) => prev.filter((p) => p.studentId !== studentId));
  };

  const handleExcel = useCallback(
    (file: File, rosterId: string) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const wb = XLSX.read(data, { type: "array" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows: string[][] = XLSX.utils.sheet_to_json(ws, { header: 1 });

        const header = rows[0]?.map((c) => String(c ?? "").trim()) ?? [];
        const nameColIdx = (() => {
          const idx = header.findIndex((h) => ["이름", "성명", "학생명"].includes(h));
          return idx >= 0 ? idx : 1;
        })();

        const names = rows
          .slice(1)
          .map((r) => String(r[nameColIdx] ?? "").trim())
          .filter(Boolean);

        addStudentsFromText(names.join("\n"), rosterId);
      };
      reader.readAsArrayBuffer(file);
    },
    [addStudentsFromText]
  );

  /* ─── 조편성 ─── */
  const addPin = () => {
    if (!pinStudentId) return;
    setPins((prev) => {
      const next = prev.filter((p) => p.studentId !== pinStudentId);
      return [...next, { studentId: pinStudentId, groupNum: pinGroupNum }];
    });
    setPinStudentId("");
  };

  const removePin = (studentId: string) => {
    setPins((prev) => prev.filter((p) => p.studentId !== studentId));
  };

  const doAssign = () => {
    if (!selectedRoster || selectedRoster.students.length === 0) return;
    const validPins = pins.filter(
      (p) => p.groupNum <= numGroups && selectedRoster.students.some((s) => s.id === p.studentId)
    );
    setResult(assignGroups(selectedRoster.students, numGroups, validPins));
    setTab(2);
  };

  /* ─── 결과 복사 ─── */
  const copyResult = async () => {
    if (!result) return;
    const text = result
      .map((g) => `[${g.groupNum}조]\n${g.students.map((s) => s.name).join(", ")}`)
      .join("\n\n");
    await navigator.clipboard.writeText(text);
    setCopyDone(true);
    setTimeout(() => setCopyDone(false), 2000);
  };

  /* ─── UI ─── */
  return (
    <div
      className="min-h-screen"
      style={{ background: "linear-gradient(135deg, #fdf4ff 0%, #ffe8f0 50%, #e8f4ff 100%)" }}
    >
      {/* Header */}
      <header className="text-center pt-8 pb-2">
        <h1 className="text-3xl font-bold" style={{ color: "#c084fc" }}>
          🌸 조편성 도우미
        </h1>
        <p className="text-sm mt-1" style={{ color: "#a78bca" }}>
          쉽고 빠른 학급 조편성
        </p>
      </header>

      {/* Tabs */}
      <div className="flex justify-center gap-2 mt-6 px-4">
        {TAB_LABELS.map((label, i) => (
          <button
            key={i}
            onClick={() => setTab(i as Tab)}
            className="px-4 py-2 rounded-full text-sm font-semibold transition-all"
            style={
              tab === i
                ? { background: "#c084fc", color: "white", boxShadow: "0 4px 12px #c084fc66" }
                : { background: "white", color: "#a78bca", border: "2px solid #e9d5ff" }
            }
          >
            {label}
          </button>
        ))}
      </div>

      <main className="max-w-2xl mx-auto px-4 py-6">
        {/* ═══ TAB 0: 명렬 관리 ═══ */}
        {tab === 0 && (
          <div className="space-y-4">
            {/* 반 추가 */}
            <div
              className="bg-white rounded-2xl p-5 shadow-sm"
              style={{ border: "2px solid #e9d5ff" }}
            >
              <h2 className="font-bold mb-3" style={{ color: "#9333ea" }}>
                반 추가
              </h2>
              <div className="flex gap-2">
                <input
                  value={newClassName}
                  onChange={(e) => setNewClassName(e.target.value)}
                  onKeyDown={(e) => e.key === "Enter" && addRoster()}
                  placeholder="예: 2-6반, 1반, 과학탐구부..."
                  className="flex-1 rounded-xl px-4 py-2 text-sm outline-none"
                  style={{ border: "2px solid #e9d5ff", color: "#4a3f5c" }}
                />
                <button
                  onClick={addRoster}
                  className="px-4 py-2 rounded-xl text-sm font-bold text-white"
                  style={{ background: "#c084fc" }}
                >
                  추가
                </button>
              </div>
            </div>

            {/* 반 목록 */}
            {rosters.length > 0 && (
              <div className="flex flex-wrap gap-2">
                {rosters.map((r) => (
                  <button
                    key={r.id}
                    onClick={() => setSelectedRosterId(r.id)}
                    className="flex items-center gap-2 px-3 py-1.5 rounded-full text-sm font-semibold transition-all"
                    style={
                      selectedRosterId === r.id
                        ? { background: "#f0abfc", color: "white" }
                        : { background: "white", color: "#9333ea", border: "2px solid #f0abfc" }
                    }
                  >
                    {r.className}
                    <span className="text-xs opacity-70">({r.students.length}명)</span>
                    <span
                      onClick={(e) => {
                        e.stopPropagation();
                        deleteRoster(r.id);
                      }}
                      className="ml-1 opacity-60 hover:opacity-100 cursor-pointer"
                    >
                      ✕
                    </span>
                  </button>
                ))}
              </div>
            )}

            {/* 학생 입력 */}
            {selectedRoster && (
              <div
                className="bg-white rounded-2xl p-5 shadow-sm"
                style={{ border: "2px solid #e9d5ff" }}
              >
                <h2 className="font-bold mb-1" style={{ color: "#9333ea" }}>
                  {selectedRoster.className} 학생 추가
                </h2>
                <p className="text-xs mb-3" style={{ color: "#a78bca" }}>
                  이름을 줄바꿈·쉼표로 구분해 입력하거나, 엑셀 파일을 업로드하세요.
                </p>

                <textarea
                  value={directInput}
                  onChange={(e) => setDirectInput(e.target.value)}
                  placeholder={"김민지\n박서연\n이지우\n..."}
                  rows={5}
                  className="w-full rounded-xl px-4 py-3 text-sm outline-none resize-none"
                  style={{ border: "2px solid #e9d5ff", color: "#4a3f5c" }}
                />
                <div className="flex gap-2 mt-2">
                  <button
                    onClick={() => addStudentsFromText(directInput, selectedRoster.id)}
                    className="flex-1 py-2 rounded-xl text-sm font-bold text-white"
                    style={{ background: "#c084fc" }}
                  >
                    입력 추가
                  </button>
                  <button
                    onClick={() => fileRef.current?.click()}
                    className="flex-1 py-2 rounded-xl text-sm font-bold"
                    style={{
                      background: "white",
                      color: "#9333ea",
                      border: "2px solid #e9d5ff",
                    }}
                  >
                    엑셀 업로드 📂
                  </button>
                  <input
                    ref={fileRef}
                    type="file"
                    accept=".xlsx,.xls"
                    className="hidden"
                    onChange={(e) => {
                      const f = e.target.files?.[0];
                      if (f && selectedRoster) handleExcel(f, selectedRoster.id);
                      e.target.value = "";
                    }}
                  />
                </div>
              </div>
            )}

            {/* 학생 목록 */}
            {selectedRoster && selectedRoster.students.length > 0 && (
              <div
                className="bg-white rounded-2xl p-5 shadow-sm"
                style={{ border: "2px solid #e9d5ff" }}
              >
                <div className="flex justify-between items-center mb-3">
                  <h2 className="font-bold" style={{ color: "#9333ea" }}>
                    학생 목록 ({selectedRoster.students.length}명)
                  </h2>
                  <button
                    onClick={() =>
                      setRosters((prev) =>
                        prev.map((r) =>
                          r.id === selectedRoster.id ? { ...r, students: [] } : r
                        )
                      )
                    }
                    className="text-xs px-3 py-1 rounded-full"
                    style={{
                      color: "#f87171",
                      border: "1.5px solid #fecaca",
                      background: "white",
                    }}
                  >
                    전체 삭제
                  </button>
                </div>
                <div className="flex flex-wrap gap-2">
                  {selectedRoster.students.map((s) => (
                    <span
                      key={s.id}
                      className="flex items-center gap-1 px-3 py-1 rounded-full text-sm pop-in"
                      style={{
                        background: "#fdf4ff",
                        color: "#7e22ce",
                        border: "1.5px solid #e9d5ff",
                      }}
                    >
                      {s.name}
                      <button
                        onClick={() => removeStudent(selectedRoster.id, s.id)}
                        className="opacity-40 hover:opacity-100 text-xs"
                      >
                        ✕
                      </button>
                    </span>
                  ))}
                </div>
              </div>
            )}

            {rosters.length === 0 && (
              <div className="text-center py-12 opacity-60" style={{ color: "#a78bca" }}>
                <p className="text-4xl mb-2">📚</p>
                <p className="text-sm">반을 먼저 추가해 주세요!</p>
              </div>
            )}
          </div>
        )}

        {/* ═══ TAB 1: 조편성 설정 ═══ */}
        {tab === 1 && (
          <div className="space-y-4">
            {/* 반 선택 */}
            <div
              className="bg-white rounded-2xl p-5 shadow-sm"
              style={{ border: "2px solid #fce7f3" }}
            >
              <h2 className="font-bold mb-3" style={{ color: "#db2777" }}>
                반 선택
              </h2>
              {rosters.length === 0 ? (
                <p className="text-sm" style={{ color: "#f9a8d4" }}>
                  먼저 명렬 관리에서 반을 추가해 주세요.
                </p>
              ) : (
                <div className="flex flex-wrap gap-2">
                  {rosters.map((r) => (
                    <button
                      key={r.id}
                      onClick={() => {
                        setSelectedRosterId(r.id);
                        setPins([]);
                      }}
                      className="px-4 py-2 rounded-full text-sm font-semibold"
                      style={
                        selectedRosterId === r.id
                          ? { background: "#fb7185", color: "white" }
                          : {
                              background: "white",
                              color: "#db2777",
                              border: "2px solid #fce7f3",
                            }
                      }
                    >
                      {r.className} ({r.students.length}명)
                    </button>
                  ))}
                </div>
              )}
            </div>

            {selectedRoster && selectedRoster.students.length > 0 && (
              <>
                {/* 조 수 설정 */}
                <div
                  className="bg-white rounded-2xl p-5 shadow-sm"
                  style={{ border: "2px solid #fce7f3" }}
                >
                  <h2 className="font-bold mb-4" style={{ color: "#db2777" }}>
                    조 수 설정
                  </h2>
                  <div className="flex items-center gap-4">
                    <button
                      onClick={() => setNumGroups((n) => Math.max(2, n - 1))}
                      className="w-10 h-10 rounded-full text-xl font-bold flex items-center justify-center"
                      style={{ background: "#fce7f3", color: "#db2777" }}
                    >
                      −
                    </button>
                    <span
                      className="text-3xl font-bold"
                      style={{ color: "#db2777", minWidth: 70, textAlign: "center" }}
                    >
                      {numGroups}조
                    </span>
                    <button
                      onClick={() =>
                        setNumGroups((n) => Math.min(selectedRoster.students.length, n + 1))
                      }
                      className="w-10 h-10 rounded-full text-xl font-bold flex items-center justify-center"
                      style={{ background: "#fce7f3", color: "#db2777" }}
                    >
                      +
                    </button>
                  </div>
                  <p className="text-xs mt-3" style={{ color: "#f9a8d4" }}>
                    {selectedRoster.students.length}명 → 조당 약{" "}
                    {Math.ceil(selectedRoster.students.length / numGroups)}명
                  </p>
                </div>

                {/* 수동 배정 */}
                <div
                  className="bg-white rounded-2xl p-5 shadow-sm"
                  style={{ border: "2px solid #fce7f3" }}
                >
                  <h2 className="font-bold mb-1" style={{ color: "#db2777" }}>
                    수동 배정 <span className="font-normal text-sm opacity-70">(선택)</span>
                  </h2>
                  <p className="text-xs mb-3" style={{ color: "#f9a8d4" }}>
                    특정 학생을 고정 배치하고 나머지는 자동 배정됩니다.
                  </p>
                  <div className="flex gap-2 mb-3">
                    <select
                      value={pinStudentId}
                      onChange={(e) => setPinStudentId(e.target.value)}
                      className="flex-1 rounded-xl px-3 py-2 text-sm outline-none"
                      style={{ border: "2px solid #fce7f3", color: "#4a3f5c" }}
                    >
                      <option value="">학생 선택</option>
                      {selectedRoster.students.map((s) => {
                        const pin = pins.find((p) => p.studentId === s.id);
                        return (
                          <option key={s.id} value={s.id}>
                            {s.name}
                            {pin ? ` → ${pin.groupNum}조 고정` : ""}
                          </option>
                        );
                      })}
                    </select>
                    <select
                      value={pinGroupNum}
                      onChange={(e) => setPinGroupNum(Number(e.target.value))}
                      className="w-24 rounded-xl px-3 py-2 text-sm outline-none"
                      style={{ border: "2px solid #fce7f3", color: "#4a3f5c" }}
                    >
                      {Array.from({ length: numGroups }, (_, i) => (
                        <option key={i + 1} value={i + 1}>
                          {i + 1}조
                        </option>
                      ))}
                    </select>
                    <button
                      onClick={addPin}
                      className="px-4 py-2 rounded-xl text-sm font-bold text-white"
                      style={{ background: "#fb7185" }}
                    >
                      고정
                    </button>
                  </div>

                  {pins.length > 0 && (
                    <div className="flex flex-wrap gap-2">
                      {pins.map((p) => {
                        const student = selectedRoster.students.find((s) => s.id === p.studentId);
                        if (!student) return null;
                        return (
                          <span
                            key={p.studentId}
                            className="flex items-center gap-1 px-3 py-1 rounded-full text-sm"
                            style={{
                              background: "#fff1f2",
                              color: "#be123c",
                              border: "1.5px solid #fecdd3",
                            }}
                          >
                            📌 {student.name} → {p.groupNum}조
                            <button
                              onClick={() => removePin(p.studentId)}
                              className="opacity-40 hover:opacity-100 text-xs ml-1"
                            >
                              ✕
                            </button>
                          </span>
                        );
                      })}
                    </div>
                  )}
                </div>

                {/* 배정 버튼 */}
                <button
                  onClick={doAssign}
                  className="w-full py-4 rounded-2xl text-lg font-bold text-white transition-transform active:scale-95"
                  style={{
                    background: "linear-gradient(135deg, #f472b6, #c084fc)",
                    boxShadow: "0 4px 20px #c084fc55",
                  }}
                >
                  🎲 랜덤 조편성 시작!
                </button>
              </>
            )}

            {(!selectedRoster || selectedRoster.students.length === 0) &&
              rosters.length > 0 && (
                <div className="text-center py-8 opacity-60" style={{ color: "#f9a8d4" }}>
                  <p className="text-4xl mb-2">👆</p>
                  <p className="text-sm">반을 선택하고 학생을 추가해 주세요.</p>
                </div>
              )}
          </div>
        )}

        {/* ═══ TAB 2: 결과 ═══ */}
        {tab === 2 && (
          <div className="space-y-4">
            {result ? (
              <>
                <div className="flex justify-between items-center">
                  <h2 className="font-bold text-lg" style={{ color: "#0369a1" }}>
                    {selectedRoster?.className} 결과
                  </h2>
                  <div className="flex gap-2">
                    <button
                      onClick={copyResult}
                      className="px-4 py-2 rounded-full text-sm font-bold transition-all"
                      style={
                        copyDone
                          ? { background: "#bbf7d0", color: "#15803d" }
                          : {
                              background: "white",
                              color: "#0369a1",
                              border: "2px solid #bae6fd",
                            }
                      }
                    >
                      {copyDone ? "복사됨 ✓" : "복사 📋"}
                    </button>
                    <button
                      onClick={() => window.print()}
                      className="px-4 py-2 rounded-full text-sm font-bold"
                      style={{
                        background: "white",
                        color: "#0369a1",
                        border: "2px solid #bae6fd",
                      }}
                    >
                      인쇄 🖨️
                    </button>
                  </div>
                </div>

                <div className="grid grid-cols-2 gap-3">
                  {result.map((g) => (
                    <div
                      key={g.groupNum}
                      className="bg-white rounded-2xl p-4 shadow-sm pop-in"
                      style={{ border: "2px solid #bae6fd" }}
                    >
                      <div
                        className="text-sm font-bold mb-2 px-3 py-1 rounded-full inline-block"
                        style={{ background: "#e0f2fe", color: "#0369a1" }}
                      >
                        {g.groupNum}조
                      </div>
                      <ul className="space-y-1 mt-1">
                        {g.students.map((s) => {
                          const isPinned = pins.some((p) => p.studentId === s.id);
                          return (
                            <li
                              key={s.id}
                              className="text-sm flex items-center gap-1"
                              style={{ color: "#334155" }}
                            >
                              {isPinned && <span className="text-xs">📌</span>}
                              {s.name}
                            </li>
                          );
                        })}
                      </ul>
                    </div>
                  ))}
                </div>

                <button
                  onClick={doAssign}
                  className="w-full py-3 rounded-2xl text-sm font-bold text-white"
                  style={{ background: "linear-gradient(135deg, #60a5fa, #a78bfa)" }}
                >
                  🔄 다시 뽑기
                </button>
                <button
                  onClick={() => setTab(1)}
                  className="w-full py-3 rounded-2xl text-sm font-bold"
                  style={{
                    background: "white",
                    color: "#0369a1",
                    border: "2px solid #bae6fd",
                  }}
                >
                  설정 변경하기
                </button>
              </>
            ) : (
              <div className="text-center py-12 opacity-60" style={{ color: "#7dd3fc" }}>
                <p className="text-4xl mb-2">🎲</p>
                <p className="text-sm">조편성 설정에서 배정을 먼저 실행해 주세요!</p>
              </div>
            )}
          </div>
        )}
      </main>
    </div>
  );
}

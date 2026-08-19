"use client";

import React, { useState, useEffect, useCallback, useRef } from "react";
import Image from "next/image";
import * as XLSX from "xlsx";

// ─── Constants ──────────────────────────────────────────────────────────────────

const CORRECT_PASSWORD = process.env.NEXT_PUBLIC_ADMIN_PASSWORD || "internalaffairs@abhiraj";

const DAY_ORDERS = ["DO1", "DO2", "DO3", "DO4", "DO5"] as const;
type DayOrder = (typeof DAY_ORDERS)[number];

/**
 * Canonical hourly slots matching the university timetable.
 * Format: "Hour N (start-end)" — these are used both as slot labels in output
 * and as matching keys against the spreadsheet values.
 */
const HOUR_SLOTS = [
  { hour: 1, label: "Hour 1 (8:00-8:50)",   time: "8:00-8:50" },
  { hour: 2, label: "Hour 2 (8:50-9:40)",   time: "8:50-9:40" },
  { hour: 3, label: "Hour 3 (9:45-10:35)",  time: "9:45-10:35" },
  { hour: 4, label: "Hour 4 (10:40-11:30)", time: "10:40-11:30" },
  { hour: 5, label: "Hour 5 (11:35-12:25)", time: "11:35-12:25" },
  { hour: 6, label: "Hour 6 (12:30-1:20)",  time: "12:30-1:20" },
  { hour: 7, label: "Hour 7 (1:25-2:15)",   time: "1:25-2:15" },
  { hour: 8, label: "Hour 8 (2:15-3:10)",   time: "2:15-3:10" },
  { hour: 9, label: "Hour 9 (3:15-4:00)",   time: "3:15-4:00" },
  { hour: 10, label: "Hour 10 (4:00-4:50)", time: "4:00-4:50" },
] as const;

const DOMAINS_LOWER = [
  "da",
  "operations",
  "hospitality",
  "brm",
  "decor",
  "pr",
  "design",
];

// Short labels for output
const DOMAIN_SHORT: Record<string, string> = {
  da: "da",
  operations: "ops",
  hospitality: "hospitality",
  brm: "brm",
  decor: "decor",
  pr: "pr",
  design: "design",
};

const LS_HEADS_DATA = "mun_heads_data";
const LS_HEADS_NAME = "mun_heads_name";
const LS_MEMBERS_DATA = "mun_members_data";
const LS_MEMBERS_NAME = "mun_members_name";

// ─── Types ──────────────────────────────────────────────────────────────────────

interface HeadRow {
  Name: string;
  [dayColumn: string]: string;
}

interface MemberRow {
  Name: string;
  Domain: string;
  [dayColumn: string]: string;
}

// ─── Helpers ────────────────────────────────────────────────────────────────────

/** Fisher-Yates shuffle (in-place). */
function shuffleArray<T>(arr: T[]): T[] {
  for (let i = arr.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [arr[i], arr[j]] = [arr[j], arr[i]];
  }
  return arr;
}

/** Random integer between min and max (inclusive). */
function randBetween(min: number, max: number): number {
  return Math.floor(Math.random() * (max - min + 1)) + min;
}

/**
 * Check whether a person's availability string for a given day contains
 * the target hourly slot.
 *
 * The spreadsheet values look like:
 *   "Hour 8 (2:20-3:10), Hour 9 (3:10-4:00), Hour 10 (4:00-4:50)"
 *
 * We extract the hour numbers from both the target and the availability
 * string, then check for a match.
 */
function isAvailableForSlot(
  availabilityStr: string | undefined,
  hourNumber: number
): boolean {
  if (!availabilityStr || availabilityStr.trim() === "") return false;

  // 1. Extract all "Hour N" numbers from the availability string
  const hourRegex = /Hour\s+(\d+)/gi;
  let match;
  let foundHourMatch = false;
  while ((match = hourRegex.exec(availabilityStr)) !== null) {
    foundHourMatch = true;
    if (parseInt(match[1], 10) === hourNumber) return true;
  }
  if (foundHourMatch) return false;

  // 2. Otherwise, check for time strings.
  // The heads form uses variations like "11:35 am to 12:25 pm" vs "11:30-12:20"
  // We can just use the start hour:minute to identify the slot.
  const hourSignatures: Record<number, string[]> = {
    1: ["8:00"],
    2: ["8:50"],
    3: ["9:45"],
    4: ["10:40"],
    5: ["11:30", "11:35"],
    6: ["12:25", "12:30"],
    7: ["1:25"],
    8: ["2:15", "2:20"],
    9: ["3:10", "3:15"],
    10: ["4:00"]
  };
  
  const signatures = hourSignatures[hourNumber] || [];
  const parts = availabilityStr.split(",").map((s) => s.trim());
  for (const part of parts) {
    for (const sig of signatures) {
      if (part.startsWith(sig)) return true;
    }
  }

  return false;
}

/**
 * Resolve which column in the parsed data corresponds to the chosen Day Order.
 *
 * Handles real column names like:
 *   "Day order 1 free slots", "Day order 2 free slots", ...
 * as well as simpler variants:
 *   "Day 1", "DO1", etc.
 */
function resolveDayColumn(keys: string[], dayOrder: DayOrder): string | null {
  const dayNum = dayOrder.replace("DO", "");

  // 1. Match "Day order N free slots" or "Free hours on DON" (the Google Form formats)
  for (const k of keys) {
    const lower = k.toLowerCase().trim();
    if (
      lower === `day order ${dayNum} free slots` ||
      lower.startsWith(`day order ${dayNum}`) ||
      lower.includes(`free hours on do${dayNum}`) ||
      lower.includes(`free hours on ${dayOrder.toLowerCase()}`)
    ) {
      return k;
    }
  }

  // 2. Match "Day N" or "DO1" style
  for (const k of keys) {
    const lower = k.toLowerCase().trim();
    if (
      lower === `day ${dayNum}` ||
      lower === `day${dayNum}` ||
      lower === dayOrder.toLowerCase()
    ) {
      return k;
    }
  }

  // 3. Fuzzy: column containing "day" AND the day number
  for (const k of keys) {
    const lower = k.toLowerCase();
    if (lower.includes("day") && lower.includes(dayNum)) {
      return k;
    }
  }

  // 4. Fuzzy fallback: column containing dayOrder directly
  for (const k of keys) {
    if (k.toLowerCase().includes(dayOrder.toLowerCase())) {
      return k;
    }
  }

  return null;
}

/**
 * Sort by duty tally ascending, shuffle ties.
 */
function fairnessSort<T extends { name: string }>(
  pool: T[],
  tally: Record<string, number>
): T[] {
  // Group by count
  const grouped: Record<number, T[]> = {};
  for (const p of pool) {
    const count = tally[p.name] ?? 0;
    if (!grouped[count]) grouped[count] = [];
    grouped[count].push(p);
  }

  const sorted: T[] = [];
  const counts = Object.keys(grouped)
    .map(Number)
    .sort((a, b) => a - b);
  for (const c of counts) {
    sorted.push(...shuffleArray(grouped[c]));
  }
  return sorted;
}

// ─── Allocation Engine ──────────────────────────────────────────────────────────

function generateRoster(
  headsData: HeadRow[],
  membersData: MemberRow[],
  dayOrder: DayOrder
): { headsOutput: string; membersOutput: string } {
  const slots = HOUR_SLOTS;

  // Resolve column names
  const headKeys = headsData.length > 0 ? Object.keys(headsData[0]) : [];
  const memberKeys = membersData.length > 0 ? Object.keys(membersData[0]) : [];
  const headDayCol = resolveDayColumn(headKeys, dayOrder);
  const memberDayCol = resolveDayColumn(memberKeys, dayOrder);

  // Duty tallies
  const headTally: Record<string, number> = {};
  const memberTally: Record<string, number> = {};

  // Initialize tallies
  headsData.forEach((h) => {
    if (h.Name || h["Full name"]) {
      const name = (h.Name || h["Full name"]).trim();
      headTally[name] = 0;
    }
  });
  membersData.forEach((m) => {
    if (m.Name) memberTally[m.Name.trim()] = 0;
  });

  // Calculate max duties for heads based on free hours (free/2 rounded)
  const headMaxDuty: Record<string, number> = {};
  if (headDayCol) {
    for (const h of headsData) {
      const name = (h.Name || h["Full name"] || "").trim();
      if (!name) continue;
      let freeCount = 0;
      for (let i = 0; i < slots.length; i++) {
        if (isAvailableForSlot(h[headDayCol], slots[i].hour)) {
          freeCount++;
        }
      }
      headMaxDuty[name] = Math.max(1, Math.round(freeCount / 2));
    }
  }

  // ── HEADS ALLOCATION ──────────────────────────────────────────────────────
  const headsLines: string[] = [`UB DESK DUTY (HEADS) - ${dayOrder}`, ""];

  for (let i = 0; i < slots.length; i++) {
    const { hour, label } = slots[i];
    // Target at least 4 heads per hour
    const target = randBetween(4, 5);

    // Find available heads
    let available: { name: string }[] = [];
    let fallback: { name: string }[] = [];
    if (headDayCol) {
      for (const h of headsData) {
        const name = (h.Name || h["Full name"] || "").trim();
        if (!name) continue;
        if (isAvailableForSlot(h[headDayCol], hour)) {
          // Check if they are under their maximum duty cap
          if ((headTally[name] || 0) < (headMaxDuty[name] || 99)) {
            available.push({ name });
          } else {
            fallback.push({ name });
          }
        }
      }
    }

    // Fairness sort & pick
    available = fairnessSort(available, headTally);
    let picked = available.slice(0, target);
    
    // Fill from fallback if needed to reach target
    if (picked.length < target && fallback.length > 0) {
      fallback = fairnessSort(fallback, headTally);
      picked.push(...fallback.slice(0, target - picked.length));
    }

    headsLines.push(`${label}`);
    if (picked.length === 0) {
      headsLines.push("UNSTAFFED");
    } else {
      for (const p of picked) {
        headsLines.push(p.name);
        headTally[p.name] = (headTally[p.name] ?? 0) + 1;
      }
    }
    headsLines.push("");
  }

  // ── MEMBERS ALLOCATION ────────────────────────────────────────────────────
  const membersLines: string[] = [`UB DESK DUTY - ${dayOrder}`, ""];

  for (let i = 0; i < slots.length; i++) {
    const { hour, label } = slots[i];
    const totalTarget = randBetween(8, 9);
    const DA_QUOTA = 3;
    const OPS_QUOTA = 3;

    // Build availability pools by domain
    const domainPools: Record<string, { name: string; domain: string }[]> = {};
    for (const d of DOMAINS_LOWER) domainPools[d] = [];

    if (memberDayCol) {
      for (const m of membersData) {
        const name = (m.Name ?? "").trim();
        const domain = (m.Domain ?? "").trim().toLowerCase();
        if (!name) continue;
        if (isAvailableForSlot(m[memberDayCol], hour)) {
          const resolved = DOMAINS_LOWER.includes(domain) ? domain : "";
          if (resolved) {
            domainPools[resolved].push({ name, domain: resolved });
          }
        }
      }
    }

    // Sort each pool by fairness
    for (const d of DOMAINS_LOWER) {
      domainPools[d] = fairnessSort(domainPools[d], memberTally);
    }

    const picked: { name: string; domain: string }[] = [];
    const pickedNames = new Set<string>();

    // Helper: pick N from a domain pool
    const pickFrom = (
      domain: string,
      count: number
    ): { name: string; domain: string }[] => {
      const result: { name: string; domain: string }[] = [];
      for (const p of domainPools[domain]) {
        if (result.length >= count) break;
        if (!pickedNames.has(p.name)) {
          result.push(p);
          pickedNames.add(p.name);
        }
      }
      return result;
    };

    // 1. Mandatory DA
    const daMembers = pickFrom("da", DA_QUOTA);
    picked.push(...daMembers);

    // 2. Mandatory Operations
    const opsMembers = pickFrom("operations", OPS_QUOTA);
    picked.push(...opsMembers);

    // 3. Calculate remaining spots — fill up to totalTarget
    // Build wildcard pool: everyone NOT already picked, from any domain
    let wildcardPool: { name: string; domain: string }[] = [];
    for (const d of DOMAINS_LOWER) {
      for (const p of domainPools[d]) {
        if (!pickedNames.has(p.name)) {
          wildcardPool.push(p);
        }
      }
    }
    wildcardPool = fairnessSort(wildcardPool, memberTally);

    for (const p of wildcardPool) {
      if (picked.length >= totalTarget) break;
      if (!pickedNames.has(p.name)) {
        picked.push(p);
        pickedNames.add(p.name);
      }
    }

    membersLines.push(`${label}`);
    if (picked.length === 0) {
      membersLines.push("UNSTAFFED");
    } else {
      for (const p of picked) {
        const short = DOMAIN_SHORT[p.domain] ?? p.domain;
        membersLines.push(`${p.name} (${short})`);
        memberTally[p.name] = (memberTally[p.name] ?? 0) + 1;
      }
    }
    membersLines.push("");
  }

  return {
    headsOutput: headsLines.join("\n").trim(),
    membersOutput: membersLines.join("\n").trim(),
  };
}

// ─── Parse File ─────────────────────────────────────────────────────────────────

function parseFile(file: File): Promise<Record<string, string>[]> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = e.target?.result;
        const workbook = XLSX.read(data, { type: "array" });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        const json = XLSX.utils.sheet_to_json<Record<string, string>>(sheet, {
          defval: "",
        });
        resolve(json);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = () => reject(reader.error);
    reader.readAsArrayBuffer(file);
  });
}

// ─── SVG Icons ──────────────────────────────────────────────────────────────────

const UploadIcon = () => (
  <svg
    width="40"
    height="40"
    viewBox="0 0 24 24"
    fill="none"
    stroke="currentColor"
    strokeWidth="1.5"
    strokeLinecap="round"
    strokeLinejoin="round"
    className="mx-auto mb-3 text-mun-gold/25"
  >
    <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
    <polyline points="17 8 12 3 7 8" />
    <line x1="12" y1="3" x2="12" y2="15" />
  </svg>
);

const CheckIcon = () => (
  <svg
    width="20"
    height="20"
    viewBox="0 0 24 24"
    fill="none"
    stroke="currentColor"
    strokeWidth="2.5"
    strokeLinecap="round"
    strokeLinejoin="round"
    className="text-green-400 inline mr-2"
  >
    <polyline points="20 6 9 17 4 12" />
  </svg>
);

// ─── Login Screen ───────────────────────────────────────────────────────────────

function LoginScreen({ onLogin }: { onLogin: () => void }) {
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [shake, setShake] = useState(false);
  const [showPassword, setShowPassword] = useState(false);

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (password === CORRECT_PASSWORD) {
      onLogin();
    } else {
      setError("Invalid credentials. Access denied.");
      setShake(true);
      setTimeout(() => setShake(false), 500);
    }
  };

  return (
    <div className="min-h-screen flex items-center justify-center px-4">
      <div
        className={`w-full max-w-sm animate-fade-in-up ${shake ? "animate-[shake_0.5s_ease-in-out]" : ""}`}
      >
        <div className="glass-card shimmer-border p-8 text-center">
          {/* Logo */}
          <div className="mb-6 flex justify-center">
            <Image
              src="/logo-white.png"
              alt="SRM MUN Society"
              width={200}
              height={80}
              className="opacity-90"
              priority
            />
          </div>

          <div className="gold-divider mb-6" />

          <p className="text-xs text-mun-gold/50 mb-6 font-semibold tracking-[0.15em] uppercase">
            DeskDuty Pro
          </p>

          <form onSubmit={handleSubmit} className="space-y-4">
            <div className="relative">
              <input
                id="login-password"
                type={showPassword ? "text" : "password"}
                value={password}
                onChange={(e) => {
                  setPassword(e.target.value);
                  setError("");
                }}
                placeholder="Enter access key"
                className="input-styled text-center pr-12"
                autoFocus
              />
              <button
                type="button"
                onClick={() => setShowPassword(!showPassword)}
                className="absolute right-3 top-1/2 -translate-y-1/2 p-1.5 rounded-lg text-white/30 hover:text-white/60 hover:bg-white/5 transition-all"
                aria-label={showPassword ? "Hide password" : "Show password"}
              >
                {showPassword ? (
                  /* Eye-off icon */
                  <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round">
                    <path d="M17.94 17.94A10.07 10.07 0 0 1 12 20c-7 0-11-8-11-8a18.45 18.45 0 0 1 5.06-5.94" />
                    <path d="M9.9 4.24A9.12 9.12 0 0 1 12 4c7 0 11 8 11 8a18.5 18.5 0 0 1-2.16 3.19" />
                    <path d="M14.12 14.12a3 3 0 1 1-4.24-4.24" />
                    <line x1="1" y1="1" x2="23" y2="23" />
                  </svg>
                ) : (
                  /* Eye icon */
                  <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round">
                    <path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z" />
                    <circle cx="12" cy="12" r="3" />
                  </svg>
                )}
              </button>
            </div>
            {error && (
              <p className="text-red-400 text-sm font-medium animate-fade-in-up">
                {error}
              </p>
            )}
            <button type="submit" className="btn-primary w-full text-sm">
              Authenticate →
            </button>
          </form>
        </div>
        <p className="text-center text-white/10 text-xs mt-6 font-medium">
          Authorised personnel only
        </p>
      </div>
    </div>
  );
}

// ─── File Upload Zone ───────────────────────────────────────────────────────────

function FileUploadZone({
  label,
  storedName,
  onFileLoaded,
  onRemove,
}: {
  label: string;
  storedName: string | null;
  onFileLoaded: (data: Record<string, string>[], fileName: string) => void;
  onRemove: () => void;
}) {
  const [dragOver, setDragOver] = useState(false);
  const [parsing, setParsing] = useState(false);
  const inputRef = useRef<HTMLInputElement>(null);

  const handleFile = useCallback(
    async (file: File) => {
      setParsing(true);
      try {
        const data = await parseFile(file);
        onFileLoaded(data, file.name);
      } catch {
        alert(
          `Failed to parse "${file.name}". Please upload a valid .xlsx or .csv file.`
        );
      } finally {
        setParsing(false);
      }
    },
    [onFileLoaded]
  );

  const handleDrop = useCallback(
    (e: React.DragEvent) => {
      e.preventDefault();
      setDragOver(false);
      const file = e.dataTransfer.files?.[0];
      if (file) handleFile(file);
    },
    [handleFile]
  );

  if (storedName) {
    return (
      <div className="loaded-state animate-fade-in-up">
        <div className="flex items-center justify-between gap-4">
          <div className="flex items-center min-w-0">
            <CheckIcon />
            <div className="min-w-0">
              <p className="text-sm font-semibold text-green-400">
                ✅ Data Loaded
              </p>
              <p className="text-xs text-white/40 truncate" title={storedName}>
                {storedName}
              </p>
            </div>
          </div>
          <button onClick={onRemove} className="btn-danger shrink-0">
            Remove
          </button>
        </div>
      </div>
    );
  }

  return (
    <div
      className={`upload-zone ${dragOver ? "drag-over" : ""}`}
      onDragOver={(e) => {
        e.preventDefault();
        setDragOver(true);
      }}
      onDragLeave={() => setDragOver(false)}
      onDrop={handleDrop}
      onClick={() => inputRef.current?.click()}
    >
      <input
        ref={inputRef}
        type="file"
        accept=".xlsx,.csv"
        className="hidden"
        onChange={(e) => {
          const file = e.target.files?.[0];
          if (file) handleFile(file);
        }}
      />
      {parsing ? (
        <div className="flex flex-col items-center gap-2">
          <div className="w-6 h-6 border-2 border-mun-gold border-t-transparent rounded-full animate-spin" />
          <p className="text-sm text-white/40">Parsing…</p>
        </div>
      ) : (
        <>
          <UploadIcon />
          <p className="text-sm font-semibold text-white/55 mb-1">{label}</p>
          <p className="text-xs text-white/20">
            Drag & drop or click · .xlsx / .csv
          </p>
        </>
      )}
    </div>
  );
}

// ─── Output Panel ───────────────────────────────────────────────────────────────

function OutputPanel({ label, value }: { label: string; value: string }) {
  const [copied, setCopied] = useState(false);

  const handleCopy = async () => {
    if (!value) return;
    try {
      await navigator.clipboard.writeText(value);
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
    } catch {
      // Fallback
      const ta = document.createElement("textarea");
      ta.value = value;
      document.body.appendChild(ta);
      ta.select();
      document.execCommand("copy");
      document.body.removeChild(ta);
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
    }
  };

  return (
    <div className="flex flex-col h-full">
      <div className="flex items-center justify-between mb-3">
        <span className="label-text">{label}</span>
        <button
          onClick={handleCopy}
          disabled={!value}
          className={`btn-copy ${copied ? "copied" : ""}`}
        >
          {copied ? "✓ Copied" : "Copy Text"}
        </button>
      </div>
      <textarea
        readOnly
        value={value}
        placeholder="Output will appear here after generating the roster…"
        className="output-textarea flex-1 min-h-[400px]"
      />
    </div>
  );
}

// ─── Main Dashboard ─────────────────────────────────────────────────────────────

function Dashboard() {
  const [headsData, setHeadsData] = useState<HeadRow[] | null>(null);
  const [headsFileName, setHeadsFileName] = useState<string | null>(null);
  const [membersData, setMembersData] = useState<MemberRow[] | null>(null);
  const [membersFileName, setMembersFileName] = useState<string | null>(null);
  const [dayOrder, setDayOrder] = useState<DayOrder>("DO1");
  const [headsOutput, setHeadsOutput] = useState("");
  const [membersOutput, setMembersOutput] = useState("");
  const [generating, setGenerating] = useState(false);

  // Load from localStorage on mount
  useEffect(() => {
    try {
      const hd = localStorage.getItem(LS_HEADS_DATA);
      const hn = localStorage.getItem(LS_HEADS_NAME);
      if (hd && hn) {
        setHeadsData(JSON.parse(hd));
        setHeadsFileName(hn);
      }
      const md = localStorage.getItem(LS_MEMBERS_DATA);
      const mn = localStorage.getItem(LS_MEMBERS_NAME);
      if (md && mn) {
        setMembersData(JSON.parse(md));
        setMembersFileName(mn);
      }
    } catch {
      // Ignore corrupted data
    }
  }, []);

  const handleHeadsLoaded = useCallback(
    (data: Record<string, string>[], fileName: string) => {
      const typed = data as unknown as HeadRow[];
      setHeadsData(typed);
      setHeadsFileName(fileName);
      localStorage.setItem(LS_HEADS_DATA, JSON.stringify(typed));
      localStorage.setItem(LS_HEADS_NAME, fileName);
    },
    []
  );

  const handleMembersLoaded = useCallback(
    (data: Record<string, string>[], fileName: string) => {
      const typed = data as unknown as MemberRow[];
      setMembersData(typed);
      setMembersFileName(fileName);
      localStorage.setItem(LS_MEMBERS_DATA, JSON.stringify(typed));
      localStorage.setItem(LS_MEMBERS_NAME, fileName);
    },
    []
  );

  const removeHeads = useCallback(() => {
    setHeadsData(null);
    setHeadsFileName(null);
    localStorage.removeItem(LS_HEADS_DATA);
    localStorage.removeItem(LS_HEADS_NAME);
  }, []);

  const removeMembers = useCallback(() => {
    setMembersData(null);
    setMembersFileName(null);
    localStorage.removeItem(LS_MEMBERS_DATA);
    localStorage.removeItem(LS_MEMBERS_NAME);
  }, []);

  const handleGenerate = () => {
    if (!headsData && !membersData) return;
    setGenerating(true);
    // Slight delay to show loading state
    setTimeout(() => {
      const result = generateRoster(
        headsData ?? [],
        membersData ?? [],
        dayOrder
      );
      if (headsData) setHeadsOutput(result.headsOutput);
      else setHeadsOutput("");
      if (membersData) setMembersOutput(result.membersOutput);
      else setMembersOutput("");
      setGenerating(false);
    }, 300);
  };

  const canGenerate = headsData || membersData;

  return (
    <div className="min-h-screen flex flex-col">
      {/* Header */}
      <header className="border-b border-white/5 backdrop-blur-md bg-black/30 sticky top-0 z-50">
        <div className="max-w-7xl mx-auto px-6 py-3 flex items-center justify-between">
          <div className="flex items-center gap-4">
            <Image
              src="/logo-white.png"
              alt="SRM MUN Society"
              width={140}
              height={48}
              className="opacity-85"
              priority
            />
            <div className="hidden sm:block w-px h-6 bg-white/10" />
            <p className="hidden sm:block text-[10px] text-mun-gold/50 font-semibold tracking-[0.12em] uppercase">
              DeskDuty Pro
            </p>
          </div>
          <div className="flex items-center gap-2">
            <div className="w-2 h-2 rounded-full bg-emerald-400 animate-pulse" />
            <span className="text-xs text-white/25 font-medium">Active</span>
          </div>
        </div>
      </header>

      {/* Main Content */}
      <main className="flex-1 max-w-7xl w-full mx-auto px-6 py-8 space-y-8">
        {/* Uploads + Controls Row */}
        <div className="grid grid-cols-1 lg:grid-cols-3 gap-6 animate-fade-in-up animate-delay-100">
          {/* Heads Upload */}
          <div className="glass-card p-6">
            <h2 className="label-text mb-4">📋 Heads Slots</h2>
            <FileUploadZone
              label="Upload Heads Data"
              storedName={headsFileName}
              onFileLoaded={handleHeadsLoaded}
              onRemove={removeHeads}
            />
          </div>

          {/* Members Upload */}
          <div className="glass-card p-6">
            <h2 className="label-text mb-4">👥 Members Slots</h2>
            <FileUploadZone
              label="Upload Members Data"
              storedName={membersFileName}
              onFileLoaded={handleMembersLoaded}
              onRemove={removeMembers}
            />
          </div>

          {/* Controls */}
          <div className="glass-card p-6 flex flex-col justify-between">
            <div>
              <h2 className="label-text mb-4">⚙️ Controls</h2>
              <label className="block mb-2 text-xs text-white/35 font-medium">
                Day Order
              </label>
              <select
                id="day-order-select"
                value={dayOrder}
                onChange={(e) => setDayOrder(e.target.value as DayOrder)}
                className="select-styled w-full mb-6"
              >
                {DAY_ORDERS.map((d) => (
                  <option key={d} value={d}>
                    {d}
                  </option>
                ))}
              </select>
            </div>
            <button
              id="generate-roster-btn"
              onClick={handleGenerate}
              disabled={!canGenerate || generating}
              className="btn-primary w-full flex items-center justify-center gap-2"
            >
              {generating ? (
                <>
                  <div className="w-4 h-4 border-2 border-black/40 border-t-transparent rounded-full animate-spin" />
                  Generating…
                </>
              ) : (
                <>🚀 Generate Roster</>
              )}
            </button>
            {!canGenerate && (
              <p className="text-[11px] text-white/20 mt-3 text-center">
                Upload at least one file to enable generation
              </p>
            )}
          </div>
        </div>

        {/* Output Row — show only panels for uploaded data */}
        {(headsData || membersData) && (
          <div
            className={`grid grid-cols-1 ${
              headsData && membersData ? "lg:grid-cols-2" : ""
            } gap-6 animate-fade-in-up animate-delay-300`}
          >
            {headsData && (
              <div className="glass-card p-6">
                <OutputPanel label="Heads Output" value={headsOutput} />
              </div>
            )}
            {membersData && (
              <div className="glass-card p-6">
                <OutputPanel label="Members Output" value={membersOutput} />
              </div>
            )}
          </div>
        )}
      </main>

      {/* Footer */}
      <footer className="border-t border-white/5 py-4">
        <p className="text-center text-[11px] text-white/12 font-medium">
          DeskDuty Pro · SRM MUN Society · All data stays in your browser
        </p>
      </footer>
    </div>
  );
}

// ─── Root App ───────────────────────────────────────────────────────────────────

export default function Home() {
  const [authenticated, setAuthenticated] = useState(false);

  return authenticated ? (
    <Dashboard />
  ) : (
    <LoginScreen onLogin={() => setAuthenticated(true)} />
  );
}

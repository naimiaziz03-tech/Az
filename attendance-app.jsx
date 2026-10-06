import { useState, useEffect, useMemo, useRef } from "react";
import * as XLSX from "xlsx";

const EMPLOYEE_NAME = "عزيز";

const STATUS = {
  TC: { label: "TC", color: "#3C7A5D", bg: "#E7F1EC" },
  RS: { label: "RS", color: "#C1791F", bg: "#FBF0E1" },
  RC: { label: "RC", color: "#4A6FA5", bg: "#EAF0F8" },
  DATA: { label: "✎", color: "#8A5FB0", bg: "#F1EAF8" },
};

const todayStr = () => new Date().toISOString().slice(0, 10);

const TIME_OPTIONS = Array.from({ length: 48 }, (_, i) => {
  const h = String(Math.floor(i / 2)).padStart(2, "0");
  const m = i % 2 === 0 ? "00" : "30";
  return `${h}:${m}`;
});

const diffMinutes = (start, end) => {
  if (!start || !end) return null;
  const [sh, sm] = start.split(":").map(Number);
  const [eh, em] = end.split(":").map(Number);
  let diffMin = eh * 60 + em - (sh * 60 + sm);
  if (diffMin < 0) diffMin += 24 * 60; // crosses midnight
  return diffMin;
};

const formatMinutes = (mins) => {
  if (mins === null || mins === undefined) return "—";
  const sign = mins < 0 ? "-" : "";
  const abs = Math.abs(mins);
  const h = Math.floor(abs / 60);
  const m = Math.round(abs % 60);
  return `${sign}${h}س ${m > 0 ? m + "د" : ""}`.trim();
};

const formatDiff = (start, end) => formatMinutes(diffMinutes(start, end));

const decimalHours = (start, end) => {
  if (!start || !end) return null;
  const [sh, sm] = start.split(":").map(Number);
  const [eh, em] = end.split(":").map(Number);
  let diffMin = eh * 60 + em - (sh * 60 + sm);
  if (diffMin < 0) diffMin += 24 * 60;
  return diffMin / 60;
};

const RATES = { R175: 1.75, R150: 1.5 };

const formatArabicDate = (isoDate) => {
  const d = new Date(isoDate + "T00:00:00");
  return d.toLocaleDateString("ar-EG-u-nu-latn", { weekday: "long", year: "numeric", month: "long", day: "numeric" });
};

export default function AttendanceApp() {
  const [records, setRecords] = useState({}); // { date: status }
  const [loading, setLoading] = useState(true);
  const [selectedDate, setSelectedDate] = useState(todayStr());
  const knownDayRef = useRef(todayStr());
  const dataInputRef = useRef(null);
  const [tab, setTab] = useState("today");
  const [savedFlash, setSavedFlash] = useState(false);
  const [expandedMonth, setExpandedMonth] = useState(null);
  const [expandedDate, setExpandedDate] = useState(null);
  const [expandedHoursMonth, setExpandedHoursMonth] = useState(null);
  const [showDataModal, setShowDataModal] = useState(false);
  const [checkpoints, setCheckpoints] = useState({}); // { date: [bool, bool, bool, bool] }
  const [notes, setNotes] = useState({}); // { date: string }
  const [workHours, setWorkHours] = useState({}); // { date: { start, end } }
  const [overtimeRate, setOvertimeRateState] = useState({}); // { date: "R175" | "R150" }
  const [confirmClear, setConfirmClear] = useState(false);
  const [showBackupBox, setShowBackupBox] = useState(false);
  const [showRestoreBox, setShowRestoreBox] = useState(false);
  const [restoreText, setRestoreText] = useState("");
  const [restoreMessage, setRestoreMessage] = useState("");
  const [viewMonth, setViewMonth] = useState(todayStr().slice(0, 7)); // YYYY-MM

  useEffect(() => {
    (async () => {
      try {
        const res = await window.storage.get("aziz-attendance-records");
        setRecords(res ? JSON.parse(res.value) : {});
      } catch {
        setRecords({});
      }
      try {
        const cp = await window.storage.get("aziz-checkpoints");
        setCheckpoints(cp ? JSON.parse(cp.value) : {});
      } catch {
        setCheckpoints({});
      }
      try {
        const nt = await window.storage.get("aziz-notes");
        setNotes(nt ? JSON.parse(nt.value) : {});
      } catch {
        setNotes({});
      }
      try {
        const wh = await window.storage.get("aziz-work-hours");
        setWorkHours(wh ? JSON.parse(wh.value) : {});
      } catch {
        setWorkHours({});
      }
      try {
        const or = await window.storage.get("aziz-overtime-rate");
        setOvertimeRateState(or ? JSON.parse(or.value) : {});
      } catch {
        setOvertimeRateState({});
      }
      setLoading(false);
    })();
  }, []);

  useEffect(() => {
    const interval = setInterval(() => {
      const nowDay = todayStr();
      if (nowDay !== knownDayRef.current) {
        setSelectedDate((prev) => (prev === knownDayRef.current ? nowDay : prev));
        knownDayRef.current = nowDay;
      }
    }, 30000); // check every 30 seconds
    return () => clearInterval(interval);
  }, []);

  const saveRecords = async (data) => {
    setRecords(data);
    try {
      await window.storage.set("aziz-attendance-records", JSON.stringify(data));
    } catch (e) {
      console.error("فشل حفظ سجل الحضور", e);
    }
  };

  const setStatus = (date, status) => {
    const next = { ...records };
    if (next[date] === status) {
      delete next[date]; // toggle off
    } else {
      next[date] = status;
    }
    saveRecords(next);
    saveNote(date, "");
    setSavedFlash(true);
    setTimeout(() => setSavedFlash(false), 900);
  };

  const saveCheckpoints = async (data) => {
    setCheckpoints(data);
    try {
      await window.storage.set("aziz-checkpoints", JSON.stringify(data));
    } catch (e) {
      console.error("فشل حفظ نقاط التحقق", e);
    }
  };

  const toggleCheckpoint = (date, index) => {
    const current = checkpoints[date] || [false, false, false, false];
    const updated = current.map((v, i) => (i === index ? !v : v));
    saveCheckpoints({ ...checkpoints, [date]: updated });
  };

  const saveNote = async (date, text) => {
    const next = { ...notes, [date]: text };
    setNotes(next);
    try {
      await window.storage.set("aziz-notes", JSON.stringify(next));
    } catch (e) {
      console.error("فشل حفظ الملاحظة", e);
    }
  };

  const setWorkHour = async (date, field, value) => {
    const current = workHours[date] || { start: "", end: "" };
    const next = { ...workHours, [date]: { ...current, [field]: value } };
    setWorkHours(next);
    try {
      await window.storage.set("aziz-work-hours", JSON.stringify(next));
    } catch (e) {
      console.error("فشل حفظ ساعات العمل", e);
    }
  };

  const toggleOvertimeRate = async (date, rateKey) => {
    const next = { ...overtimeRate };
    if (next[date] === rateKey) delete next[date]; // toggle off
    else next[date] = rateKey; // mutually exclusive - only one active
    setOvertimeRateState(next);
    try {
      await window.storage.set("aziz-overtime-rate", JSON.stringify(next));
    } catch (e) {
      console.error("فشل حفظ نسبة الإضافي", e);
    }
  };

  const handleDataInput = (date, text) => {
    saveNote(date, text);
    const next = { ...records };
    if (text.trim()) {
      next[date] = "DATA";
    } else if (next[date] === "DATA") {
      delete next[date];
    }
    saveRecords(next);
  };

  const clearAllData = async () => {
    if (!confirmClear) {
      setConfirmClear(true);
      return;
    }
    setRecords({});
    setCheckpoints({});
    setNotes({});
    setWorkHours({});
    setOvertimeRateState({});
    try {
      await Promise.all([
        window.storage.set("aziz-attendance-records", "{}"),
        window.storage.set("aziz-checkpoints", "{}"),
        window.storage.set("aziz-notes", "{}"),
        window.storage.set("aziz-work-hours", "{}"),
        window.storage.set("aziz-overtime-rate", "{}"),
      ]);
    } catch (e) {
      console.error("فشل مسح البيانات", e);
    }
    setConfirmClear(false);
    setShowDataModal(false);
  };

  const allMonthsPresent = () => {
    const keys = new Set();
    Object.keys(records).forEach((d) => keys.add(d.slice(0, 7)));
    Object.keys(checkpoints).forEach((d) => keys.add(d.slice(0, 7)));
    Object.keys(workHours).forEach((d) => keys.add(d.slice(0, 7)));
    return Array.from(keys).sort();
  };

  const exportExcel = () => {
    const months = allMonthsPresent();
    if (months.length === 0) {
      alert("لا توجد بيانات لتصديرها بعد.");
      return;
    }
    const wb = XLSX.utils.book_new();
    months.forEach((month) => {
      const allDates = new Set();
      Object.keys(records).filter((d) => d.startsWith(month)).forEach((d) => allDates.add(d));
      Object.keys(checkpoints)
        .filter((d) => d.startsWith(month) && (checkpoints[d] || []).some(Boolean))
        .forEach((d) => allDates.add(d));
      Object.keys(workHours)
        .filter((d) => d.startsWith(month) && workHours[d] && workHours[d].start && workHours[d].end)
        .forEach((d) => allDates.add(d));
      const sortedMonthDates = Array.from(allDates).sort();

      const rows = sortedMonthDates.map((date) => {
        const status = records[date] || "";
        const statusLabel = status === "DATA" ? "بيانات" : status;
        const dataText = status === "DATA" ? notes[date] || "" : "";
        const dc = checkpoints[date] || [];
        const cpTotal = computeDayTotal(dc);
        const cpDetail = dc.map((a, i) => (a ? i + 1 : null)).filter((v) => v !== null).join(",");
        const wh = workHours[date];
        const hasHours = wh && wh.start && wh.end;
        const rateKey = overtimeRate[date];
        const rateVal = rateKey ? RATES[rateKey] : "";
        const result = hasHours || status === "RC" ? dayResult(date).toFixed(2) : "";
        return {
          التاريخ: date,
          اليوم: formatArabicDate(date),
          الحالة: statusLabel,
          "بيانات مُدخلة": dataText,
          "نقاط التحقق": cpDetail,
          "مجموع النقاط": cpTotal || "",
          من: hasHours ? wh.start : "",
          إلى: hasHours ? wh.end : "",
          "مدة العمل": hasHours ? formatDiff(wh.start, wh.end) : "",
          النسبة: rateVal,
          الناتج: result,
        };
      });

      const ws = XLSX.utils.json_to_sheet(rows);
      ws["!cols"] = [
        { wch: 12 }, { wch: 22 }, { wch: 10 }, { wch: 24 }, { wch: 12 },
        { wch: 12 }, { wch: 8 }, { wch: 8 }, { wch: 10 }, { wch: 8 }, { wch: 10 },
      ];
      XLSX.utils.book_append_sheet(wb, ws, month);
    });
    XLSX.writeFile(wb, "سجل-حضور-عزيز.xlsx");
  };

  const buildBackupJSON = () =>
    JSON.stringify(
      {
        version: 1,
        exportedAt: new Date().toISOString(),
        records,
        checkpoints,
        notes,
        workHours,
        overtimeRate,
      },
      null,
      2
    );

  const restoreFromBackup = async () => {
    let parsed;
    try {
      parsed = JSON.parse(restoreText);
    } catch (e) {
      setRestoreMessage("❌ النص المُلصق ليس نسخة صالحة (تأكد من نسخ كل النص كاملاً).");
      return;
    }
    if (!parsed || typeof parsed !== "object") {
      setRestoreMessage("❌ النص المُلصق ليس نسخة صالحة.");
      return;
    }
    const newRecords = parsed.records && typeof parsed.records === "object" ? parsed.records : {};
    const newCheckpoints = parsed.checkpoints && typeof parsed.checkpoints === "object" ? parsed.checkpoints : {};
    const newNotes = parsed.notes && typeof parsed.notes === "object" ? parsed.notes : {};
    const newWorkHours = parsed.workHours && typeof parsed.workHours === "object" ? parsed.workHours : {};
    const newOvertimeRate = parsed.overtimeRate && typeof parsed.overtimeRate === "object" ? parsed.overtimeRate : {};
    setRecords(newRecords);
    setCheckpoints(newCheckpoints);
    setNotes(newNotes);
    setWorkHours(newWorkHours);
    setOvertimeRateState(newOvertimeRate);
    try {
      await Promise.all([
        window.storage.set("aziz-attendance-records", JSON.stringify(newRecords)),
        window.storage.set("aziz-checkpoints", JSON.stringify(newCheckpoints)),
        window.storage.set("aziz-notes", JSON.stringify(newNotes)),
        window.storage.set("aziz-work-hours", JSON.stringify(newWorkHours)),
        window.storage.set("aziz-overtime-rate", JSON.stringify(newOvertimeRate)),
      ]);
    } catch (e) {
      console.error("فشل حفظ البيانات المسترجعة", e);
    }
    setRestoreMessage("✓ تم استرجاع البيانات بنجاح.");
    setRestoreText("");
  };

  const currentStatus = records[selectedDate];

  const CHECKPOINT_WEIGHTS = [40, 80, 80, 130];

  const dailyTotal = useMemo(() => {
    const dc = checkpoints[selectedDate] || [false, false, false, false];
    return dc.reduce((sum, active, i) => sum + (active ? CHECKPOINT_WEIGHTS[i] : 0), 0);
  }, [checkpoints, selectedDate]);

  const computeDayTotal = (dayArr) => (dayArr || []).reduce((sum, active, i) => sum + (active ? CHECKPOINT_WEIGHTS[i] : 0), 0);

  const currentMonthKey = viewMonth; // YYYY-MM

  const monthlyCheckpointDates = useMemo(
    () =>
      Object.keys(checkpoints)
        .filter((d) => d.startsWith(currentMonthKey) && (checkpoints[d] || []).some(Boolean))
        .sort((a, b) => (a < b ? 1 : -1)),
    [checkpoints, currentMonthKey]
  );

  const monthlyTotal = useMemo(
    () => monthlyCheckpointDates.reduce((sum, d) => sum + computeDayTotal(checkpoints[d]), 0),
    [monthlyCheckpointDates, checkpoints]
  );

  const workHourDates = useMemo(
    () =>
      Object.keys(workHours)
        .filter((d) => workHours[d] && workHours[d].start && workHours[d].end)
        .sort((a, b) => (a < b ? 1 : -1)),
    [workHours]
  );

  const dayResult = (d) => {
    const wh = workHours[d];
    const hours = wh ? (decimalHours(wh.start, wh.end) || 0) : 0;
    const rateKey = overtimeRate[d];
    const rateVal = rateKey ? RATES[rateKey] : 1;
    const base = hours * rateVal;
    const deduction = records[d] === "RC" ? 8.5 : 0;
    return base - deduction;
  };

  const combinedHoursDates = useMemo(() => {
    const rcDates = Object.keys(records).filter((d) => records[d] === "RC");
    const set = new Set([...workHourDates, ...rcDates]);
    return Array.from(set).sort((a, b) => (a < b ? 1 : -1));
  }, [workHourDates, records]);

  const groupedHoursByMonth = useMemo(() => {
    const groups = {};
    combinedHoursDates.forEach((d) => {
      const month = d.slice(0, 7);
      if (!groups[month]) groups[month] = [];
      groups[month].push(d);
    });
    return groups;
  }, [combinedHoursDates]);

  const monthlyHoursDates = useMemo(
    () => combinedHoursDates.filter((d) => d.startsWith(viewMonth)),
    [combinedHoursDates, viewMonth]
  );

  const monthlyHoursTotal = useMemo(
    () => monthlyHoursDates.reduce((sum, d) => sum + dayResult(d), 0),
    [monthlyHoursDates, workHours, overtimeRate, records]
  );

  const sortedDates = useMemo(() => Object.keys(records).sort((a, b) => (a < b ? 1 : -1)), [records]);

  const isCountedDate = (d) => {
    if (records[d] === "RC" || records[d] === "RS") return false;
    if (records[d] === "DATA") return !!(notes[d] && notes[d].trim());
    return true;
  };

  const groupedByMonth = useMemo(() => {
    const groups = {};
    sortedDates.forEach((d) => {
      const month = d.slice(0, 7);
      if (!groups[month]) groups[month] = [];
      groups[month].push(d);
    });
    return groups;
  }, [sortedDates]);

  const countedDays = useMemo(
    () => sortedDates.filter((d) => d.startsWith(todayStr().slice(0, 7)) && isCountedDate(d)).length,
    [sortedDates, records, notes]
  );

  const missingDays = useMemo(() => {
    const now = new Date();
    const year = now.getFullYear();
    const month = now.getMonth(); // 0-indexed
    const todayDay = now.getDate();
    let missing = 0;
    for (let day = 1; day <= todayDay; day++) {
      const dateStr = `${year}-${String(month + 1).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
      if (!records[dateStr]) missing++;
    }
    return missing;
  }, [records]);

  const selectedMonthKey = selectedDate.slice(0, 7);
  const selectedMonthCountedDays = useMemo(
    () => sortedDates.filter((d) => d.startsWith(selectedMonthKey) && isCountedDate(d)).length,
    [sortedDates, selectedMonthKey, records, notes]
  );
  const selectedMonthHours = useMemo(
    () => combinedHoursDates.filter((d) => d.startsWith(selectedMonthKey)).reduce((sum, d) => sum + dayResult(d), 0),
    [combinedHoursDates, selectedMonthKey, workHours, overtimeRate, records]
  );

  if (loading) {
    return (
      <div dir="rtl" style={{ fontFamily: "'Tajawal', sans-serif" }} className="min-h-screen flex items-center justify-center bg-[#F5F6F2]">
        <div className="text-[#5B6470] text-sm">جارِ التحميل...</div>
      </div>
    );
  }

  return (
    <div dir="rtl" style={{ fontFamily: "'Tajawal', sans-serif" }} className="min-h-screen bg-[#F5F6F2] text-[#1B2430] pb-16">
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Tajawal:wght@400;500;700;900&family=Cairo:wght@700;800&display=swap');
        .font-display { font-family: 'Cairo', 'Tajawal', sans-serif; }
        .stamp-btn { transition: transform 0.12s ease, box-shadow 0.12s ease; }
        .stamp-btn:active { transform: scale(0.93); }
        .ledger-row { border-bottom: 1px dashed #D7DAD3; }
        .ledger-row:last-child { border-bottom: none; }
      `}</style>

      {/* Header with centered name */}
      <header className="bg-[#0B4F4C] text-white text-center pt-10 pb-8 px-5 relative">
        <button
          onClick={() => setShowDataModal(true)}
          title="عرض ملف تخزين البيانات"
          className="absolute top-4 left-4 z-10 w-10 h-10 rounded-full bg-[#126762] border border-[#2E8A83] flex items-center justify-center hover:bg-[#158079] text-lg"
        >
          🗄️
        </button>
        <p className="text-[#9FC7C0] text-xs tracking-widest mb-2">سجلّ الحضور</p>
        <div className="w-16 h-16 rounded-full bg-[#126762] border border-[#2E8A83] flex items-center justify-center mx-auto font-display text-2xl font-extrabold">
          {EMPLOYEE_NAME.charAt(0)}
        </div>

        <div className="flex justify-center gap-1 bg-[#0A4441] rounded-xl p-1 mt-6 max-w-sm mx-auto">
          {[
            { id: "today", label: "تسجيل" },
            { id: "history", label: "السجل" },
            { id: "monthly", label: "الشهري" },
          ].map((t) => (
            <button
              key={t.id}
              onClick={() => setTab(t.id)}
              className={`flex-1 py-2 rounded-lg text-sm font-medium transition-colors ${
                tab === t.id ? "bg-[#F5F6F2] text-[#0B4F4C]" : "text-[#B9D6D1] hover:text-white"
              }`}
            >
              {t.label}
            </button>
          ))}
        </div>
      </header>

      <main className="max-w-md mx-auto px-5 mt-6">
        {tab === "today" && (
          <div className="bg-white rounded-xl shadow-sm border border-[#E4E6DF] p-6 text-center">
            <label className="text-xs text-[#8A93A0] block mb-1">التاريخ</label>
            <input
              type="date"
              value={selectedDate}
              onChange={(e) => setSelectedDate(e.target.value)}
              className="border border-[#D7DAD3] rounded-lg px-3 py-1.5 text-sm outline-none focus:border-[#0B4F4C] focus:ring-2 focus:ring-[#0B4F4C]/20 mb-1"
            />
            <p className="text-xs text-[#8A93A0] mb-4">{formatArabicDate(selectedDate)}</p>

            <div className="flex gap-2 mb-6">
              <div className="flex-1 bg-[#F5F6F2] rounded-lg px-3 py-2 text-center">
                <p className="text-[10px] text-[#8A93A0] mb-0.5">أيام هذا الشهر</p>
                <p className="font-display font-extrabold text-sm text-[#3C7A5D]">{selectedMonthCountedDays}</p>
              </div>
              <div className="flex-1 bg-[#F5F6F2] rounded-lg px-3 py-2 text-center">
                <p className="text-[10px] text-[#8A93A0] mb-0.5">ساعات هذا الشهر</p>
                <p className={`font-display font-extrabold text-sm ${selectedMonthHours < 0 ? "text-[#B5473F]" : "text-[#0B4F4C]"}`}>{formatMinutes(selectedMonthHours * 60)}</p>
              </div>
            </div>

            <div className="flex justify-center gap-2 flex-wrap items-stretch">
              {["TC", "RS", "RC"].map((key) => {
                const s = STATUS[key];
                const isActive = currentStatus === key;
                return (
                  <button
                    key={key}
                    onClick={() => setStatus(selectedDate, key)}
                    className="stamp-btn w-16 h-16 rounded-2xl border-2 flex flex-col items-center justify-center font-display text-base font-extrabold px-1"
                    style={{
                      background: isActive ? s.bg : "#fff",
                      color: isActive ? s.color : "#9AA1AB",
                      borderColor: isActive ? s.color : "#E4E6DF",
                    }}
                  >
                    {s.label}
                  </button>
                );
              })}
              <input
                type="text"
                ref={dataInputRef}
                value={notes[selectedDate] || ""}
                onChange={(e) => {
                  handleDataInput(selectedDate, e.target.value);
                  const input = e.target;
                  requestAnimationFrame(() => {
                    input.setSelectionRange(input.value.length, input.value.length);
                    input.scrollLeft = input.scrollWidth;
                  });
                }}
                placeholder="إدخال معلومات..."
                dir="ltr"
                className="w-28 h-16 rounded-2xl border-2 text-left text-xs font-medium px-2 outline-none"
                style={{
                  background: currentStatus === "DATA" ? STATUS.DATA.bg : "#fff",
                  color: currentStatus === "DATA" ? STATUS.DATA.color : "#5B6470",
                  borderColor: currentStatus === "DATA" ? STATUS.DATA.color : "#E4E6DF",
                }}
              />
            </div>

            <div className={`text-xs mt-5 transition-opacity ${savedFlash ? "opacity-100" : "opacity-0"} text-[#3C7A5D]`}>
              تم الحفظ ✓
            </div>

            <div className="mt-6 pt-5 border-t border-dashed border-[#D7DAD3]">
              <p className="text-xs text-[#8A93A0] mb-3">نقاط التحقق</p>
              <div dir="ltr" className="flex justify-center gap-2.5">
                {(checkpoints[selectedDate] || [false, false, false, false]).map((active, i) => (
                  <div key={i} className="flex flex-col items-center gap-1">
                    <span className="text-[11px] font-bold text-[#8A93A0]">{i + 1}</span>
                    <button
                      onClick={() => toggleCheckpoint(selectedDate, i)}
                      className="stamp-btn w-12 h-12 rounded-lg border-2 flex items-center justify-center text-lg font-bold"
                      style={{
                        background: active ? "#E7F1EC" : "#fff",
                        borderColor: active ? "#3C7A5D" : "#E4E6DF",
                        color: active ? "#3C7A5D" : "#C7CCD3",
                      }}
                    >
                      {active ? "✓" : ""}
                    </button>
                  </div>
                ))}
              </div>

              <div className="mt-4 bg-[#F5F6F2] rounded-lg px-4 py-3 flex items-center justify-between">
                <span className="text-sm font-medium text-[#5B6470]">إجمالي التنشيط اليومي</span>
                <span className="font-display font-extrabold text-lg text-[#0B4F4C]">{dailyTotal}</span>
              </div>
            </div>

            <div className="mt-6 pt-5 border-t border-dashed border-[#D7DAD3]">
              <p className="text-xs text-[#8A93A0] mb-3">ساعات العمل</p>
              <div className="flex items-center gap-2">
                <div className="flex-1">
                  <label className="text-[11px] text-[#8A93A0] block mb-1">من</label>
                  <select
                    value={(workHours[selectedDate] || {}).start || ""}
                    onChange={(e) => setWorkHour(selectedDate, "start", e.target.value)}
                    className="w-full border border-[#D7DAD3] rounded-lg px-2 py-2 text-sm outline-none focus:border-[#0B4F4C] focus:ring-2 focus:ring-[#0B4F4C]/20"
                  >
                    <option value="">—</option>
                    {TIME_OPTIONS.map((t) => (
                      <option key={t} value={t}>{t}</option>
                    ))}
                  </select>
                </div>
                <div className="flex-1">
                  <label className="text-[11px] text-[#8A93A0] block mb-1">إلى</label>
                  <select
                    value={(workHours[selectedDate] || {}).end || ""}
                    onChange={(e) => setWorkHour(selectedDate, "end", e.target.value)}
                    className="w-full border border-[#D7DAD3] rounded-lg px-2 py-2 text-sm outline-none focus:border-[#0B4F4C] focus:ring-2 focus:ring-[#0B4F4C]/20"
                  >
                    <option value="">—</option>
                    {TIME_OPTIONS.map((t) => (
                      <option key={t} value={t}>{t}</option>
                    ))}
                  </select>
                </div>
                <div className="flex-1">
                  <label className="text-[11px] text-[#8A93A0] block mb-1">الفارق</label>
                  <div className="w-full border border-[#D7DAD3] rounded-lg px-2 py-2 text-sm bg-[#F5F6F2] text-center font-bold text-[#0B4F4C]">
                    {formatDiff((workHours[selectedDate] || {}).start, (workHours[selectedDate] || {}).end)}
                  </div>
                </div>
              </div>

              <div className="flex items-center gap-2 mt-3">
                {Object.entries(RATES).map(([key, val]) => {
                  const active = overtimeRate[selectedDate] === key;
                  return (
                    <button
                      key={key}
                      onClick={() => toggleOvertimeRate(selectedDate, key)}
                      className="stamp-btn flex-1 py-2 rounded-lg text-sm font-bold border-2"
                      style={{
                        background: active ? "#FBF0E1" : "#fff",
                        color: active ? "#C1791F" : "#9AA1AB",
                        borderColor: active ? "#C1791F" : "#E4E6DF",
                      }}
                    >
                      × {val}
                    </button>
                  );
                })}
              </div>
            </div>
          </div>
        )}

        {tab === "history" && (
          <div className="bg-white rounded-xl shadow-sm border border-[#E4E6DF] p-5">
            {sortedDates.length > 0 && (
              <div className="flex items-center justify-between bg-[#E7F1EC] rounded-lg px-4 py-3 mb-3">
                <span className="text-sm font-medium text-[#3C7A5D]">الأيام المحتسبة هذا الشهر</span>
                <span className="font-display font-extrabold text-lg text-[#3C7A5D]">{countedDays}</span>
              </div>
            )}
            <div className="flex items-center justify-between bg-[#FBEAE8] rounded-lg px-4 py-3 mb-4">
              <span className="text-sm font-medium text-[#B5473F]">أيام بدون بيانات (هذا الشهر)</span>
              <span className="font-display font-extrabold text-lg text-[#B5473F]">{missingDays}</span>
            </div>
            {sortedDates.length === 0 ? (
              <div className="text-center py-10 text-[#8A93A0]">
                <div className="text-3xl mb-2">📋</div>
                <p className="text-sm font-medium mb-1">لا يوجد سجل حضور بعد</p>
                <p className="text-xs">ابدأ بالتسجيل من تبويب «تسجيل».</p>
              </div>
            ) : (
              Object.keys(groupedByMonth)
                .sort((a, b) => (a < b ? 1 : -1))
                .map((month) => {
                  const isCurrentMonth = month === todayStr().slice(0, 7);
                  const isOpen =
                    Object.keys(groupedByMonth).length === 1 ||
                    expandedMonth === month ||
                    (expandedMonth === null && isCurrentMonth);
                  const monthDates = groupedByMonth[month];
                  const monthCounted = monthDates.filter(isCountedDate).length;
                  const monthLabel = new Date(month + "-01T00:00:00").toLocaleDateString("ar-EG-u-nu-latn", { year: "numeric", month: "long" });
                  return (
                    <div key={month} className="mb-2">
                      <button
                        onClick={() =>
                          setExpandedMonth(
                            isOpen ? (isCurrentMonth && expandedMonth === null ? "NONE" : null) : month
                          )
                        }
                        className="w-full flex items-center justify-between px-3 py-2 bg-[#FAFAF8] rounded-lg text-sm font-bold"
                      >
                        <span className="flex items-center gap-2">
                          {monthLabel}
                          <span className="text-xs font-bold px-2 py-0.5 rounded-full bg-[#E7F1EC] text-[#3C7A5D]">{monthCounted}</span>
                        </span>
                        <span className="text-sm">{isOpen ? "▲" : "▼"}</span>
                      </button>
                      {isOpen && (
                        <div className="px-1">
                          {monthDates.map((date) => {
                            const s = STATUS[records[date]];
                            if (!s) return null;
                            const isDateOpen = expandedDate === date;
                            return (
                              <div
                                key={date}
                                className="ledger-row py-2.5 cursor-pointer"
                                onClick={() => setExpandedDate(isDateOpen ? null : date)}
                              >
                                <div className={`flex items-center justify-between ${isDateOpen ? "flex-wrap gap-1.5" : ""}`}>
                                  <span className="text-sm">{formatArabicDate(date)}</span>
                                  <span
                                    className={`text-xs font-bold px-2.5 py-1 rounded-full ${isDateOpen ? "whitespace-normal break-words w-full text-right" : "max-w-[140px] truncate"}`}
                                    style={{ background: s.bg, color: s.color }}
                                  >
                                    {records[date] === "DATA" ? (notes[date] || "بيانات") : s.label}
                                  </span>
                                </div>
                              </div>
                            );
                          })}
                        </div>
                      )}
                    </div>
                  );
                })
            )}
          </div>
        )}

        {tab === "monthly" && (
          <div className="bg-white rounded-xl shadow-sm border border-[#E4E6DF] p-5">
            <div className="flex items-center justify-between mb-4">
              <button
                onClick={() => {
                  const [y, m] = viewMonth.split("-").map(Number);
                  const d = new Date(y, m, 1);
                  setViewMonth(`${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`);
                }}
                className="w-8 h-8 rounded-full border border-[#D7DAD3] flex items-center justify-center text-[#5B6470]"
              >
                ›
              </button>
              <span className="text-sm font-bold text-[#1B2430]">
                {new Date(viewMonth + "-01T00:00:00").toLocaleDateString("ar-EG-u-nu-latn", { year: "numeric", month: "long" })}
              </span>
              <button
                onClick={() => {
                  const [y, m] = viewMonth.split("-").map(Number);
                  const d = new Date(y, m - 2, 1);
                  setViewMonth(`${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`);
                }}
                className="w-8 h-8 rounded-full border border-[#D7DAD3] flex items-center justify-center text-[#5B6470]"
              >
                ‹
              </button>
            </div>

            <div className="flex gap-2 mb-4">
              <div className="flex-1 bg-[#0B4F4C] rounded-lg px-3 py-3 text-center">
                <p className="text-[11px] text-[#B9D6D1] mb-1">إجمالي التنشيط</p>
                <p className="font-display font-extrabold text-xl text-white">{monthlyTotal}</p>
              </div>
              <div className="flex-1 bg-[#0B4F4C] rounded-lg px-3 py-3 text-center">
                <p className="text-[11px] text-[#B9D6D1] mb-1">إجمالي الساعات</p>
                <p className={`font-display font-extrabold text-xl ${monthlyHoursTotal < 0 ? "text-[#FF9B8C]" : "text-white"}`}>{formatMinutes(monthlyHoursTotal * 60)}</p>
              </div>
            </div>

            <p className="text-xs text-[#8A93A0] mb-3">تفاصيل التنشيط اليومي</p>

            {monthlyCheckpointDates.length === 0 ? (
              <div className="text-center py-8 text-[#8A93A0]">
                <div className="text-3xl mb-2">🗓️</div>
                <p className="text-sm font-medium mb-1">لا يوجد تنشيط مسجّل لهذا الشهر</p>
                <p className="text-xs">فعّل نقاط التحقق من تبويب «تسجيل».</p>
              </div>
            ) : (
              <div className="mb-6">
                {monthlyCheckpointDates.map((date) => {
                  const dc = checkpoints[date] || [];
                  const total = computeDayTotal(dc);
                  const wh = workHours[date];
                  const hasTime = wh && wh.start && wh.end;
                  return (
                    <div key={date} className="ledger-row py-2.5">
                      <div className="flex items-center justify-between mb-1.5">
                        <span className="text-sm">{formatArabicDate(date)}</span>
                        <span className="text-xs font-bold px-2.5 py-1 rounded-full bg-[#E7F1EC] text-[#3C7A5D]">
                          {total}
                        </span>
                      </div>
                      <div dir="ltr" className="flex gap-1.5 mb-1.5">
                        {dc.map((active, i) => (
                          <span
                            key={i}
                            className="w-6 h-6 rounded text-[10px] font-bold flex items-center justify-center"
                            style={{
                              background: active ? "#E7F1EC" : "#F5F6F2",
                              color: active ? "#3C7A5D" : "#C7CCD3",
                            }}
                          >
                            {i + 1}
                          </span>
                        ))}
                      </div>
                      {hasTime && (
                        <p className="text-xs text-[#5B6470]">⏱ من {wh.start} إلى {wh.end}</p>
                      )}
                    </div>
                  );
                })}
              </div>
            )}

            <p className="text-xs text-[#8A93A0] mb-3">تفاصيل ساعات العمل</p>
            {monthlyHoursDates.length === 0 ? (
              <div className="text-center py-8 text-[#8A93A0]">
                <p className="text-sm">لا توجد ساعات عمل مسجّلة لهذا الشهر.</p>
              </div>
            ) : (
              <table className="w-full text-sm">
                <thead>
                  <tr className="text-xs text-[#8A93A0] border-b border-[#E4E6DF]">
                    <th className="text-right font-medium py-1.5">التاريخ</th>
                    <th className="text-right font-medium py-1.5">الوقت</th>
                    <th className="text-right font-medium py-1.5">النسبة</th>
                    <th className="text-right font-medium py-1.5">الناتج</th>
                  </tr>
                </thead>
                <tbody>
                  {monthlyHoursDates.map((date) => {
                    const rateKey = overtimeRate[date];
                    const rateVal = rateKey ? RATES[rateKey] : 1;
                    const result = dayResult(date).toFixed(2);
                    const hasHours = workHours[date] && workHours[date].start && workHours[date].end;
                    const isRC = records[date] === "RC";
                    return (
                      <tr key={date} className="ledger-row">
                        <td className="py-2 text-sm">
                          {formatArabicDate(date)}
                          {isRC && <span className="text-[10px] text-[#B5473F] mr-1">(RC)</span>}
                        </td>
                        <td className="py-2 text-sm">
                          {hasHours ? (
                            <div>
                              <div className="font-bold text-[#0B4F4C]">{workHours[date].start} → {workHours[date].end}</div>
                              <div className="text-[10px] text-[#8A93A0]">{formatDiff(workHours[date].start, workHours[date].end)}</div>
                            </div>
                          ) : (
                            "—"
                          )}
                        </td>
                        <td className="py-2 text-sm">{rateKey ? `× ${rateVal}` : "—"}</td>
                        <td className={`py-2 text-sm font-bold ${dayResult(date) < 0 ? "text-[#B5473F]" : "text-[#C1791F]"}`}>{result}</td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            )}
          </div>
        )}

      </main>

      {/* Storage data viewer modal */}
      {showDataModal && (
        <div
          className="fixed inset-0 bg-black/40 flex items-center justify-center z-40 px-5"
          onClick={() => { setShowDataModal(false); setConfirmClear(false); setShowBackupBox(false); setShowRestoreBox(false); setRestoreMessage(""); }}
        >
          <div
            dir="rtl"
            onClick={(e) => e.stopPropagation()}
            className="bg-white rounded-2xl shadow-xl w-full max-w-md max-h-[80vh] flex flex-col border border-[#E4E6DF]"
          >
            <div className="flex items-center justify-between px-5 pt-5 pb-3 border-b border-[#E4E6DF]">
              <div className="flex items-center gap-2">
                <span className="text-lg">🗄️</span>
                <h3 className="font-display font-bold text-base">ملف تخزين البيانات</h3>
              </div>
              <button onClick={() => { setShowDataModal(false); setConfirmClear(false); setShowBackupBox(false); setShowRestoreBox(false); setRestoreMessage(""); }} className="text-[#1B2430] hover:opacity-70 text-xl font-bold w-7 h-7 flex items-center justify-center">
                ✕
              </button>
            </div>

            <div className="overflow-y-auto px-5 py-3 flex-1">
              <p className="text-xs text-[#8A93A0] mb-3">
                عدد الأيام المحتسبة هذا الشهر: <b className="text-[#3C7A5D]">{countedDays}</b>
              </p>
              <p className="text-xs text-[#8A93A0] mb-3">
                أيام بدون بيانات (هذا الشهر حتى اليوم): <b className="text-[#B5473F]">{missingDays}</b>
              </p>
              {sortedDates.length === 0 ? (
                <p className="text-sm text-[#8A93A0] text-center py-6">لا توجد بيانات محفوظة بعد.</p>
              ) : (
                <table className="w-full text-sm">
                  <thead>
                    <tr className="text-xs text-[#8A93A0] border-b border-[#E4E6DF]">
                      <th className="text-right font-medium py-1.5">التاريخ</th>
                      <th className="text-right font-medium py-1.5">الحالة</th>
                    </tr>
                  </thead>
                  <tbody>
                    {sortedDates.map((date) => {
                      const s = STATUS[records[date]];
                      if (!s) return null;
                      return (
                        <tr key={date} className="ledger-row">
                          <td className="py-2">{date}</td>
                          <td className="py-2">
                            <span
                              className="text-xs font-bold px-2.5 py-1 rounded-full max-w-[160px] inline-block truncate align-bottom"
                              style={{ background: s.bg, color: s.color }}
                            >
                              {records[date] === "DATA" ? (notes[date] || "بيانات") : s.label}
                            </span>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              )}
            </div>

            <div className="px-5 py-3 border-t border-[#E4E6DF] bg-[#FAFAF8] rounded-b-2xl">
              <p className="text-[11px] text-[#8A93A0] leading-relaxed break-all mb-3">
                مفتاح التخزين: <code>aziz-attendance-records</code>
              </p>

              <button
                onClick={exportExcel}
                className="w-full py-2 rounded-lg text-xs font-bold bg-[#0B4F4C] text-white hover:bg-[#0A4441] mb-2"
              >
                📊 تصدير ملف Excel (كل الأشهر)
              </button>

              <button
                onClick={() => {
                  setShowBackupBox(!showBackupBox);
                  if (!showBackupBox) setShowRestoreBox(false);
                }}
                className="w-full py-2 rounded-lg text-xs font-bold bg-[#E7F1EC] text-[#3C7A5D] hover:bg-[#DAEBE3] mb-2"
              >
                {showBackupBox ? "إخفاء نسخة الاحتياط" : "📤 تصدير نسخة احتياطية (نص)"}
              </button>
              {showBackupBox && (
                <div className="mb-2">
                  <p className="text-[11px] text-[#5B6470] mb-1.5">
                    انسخ كل النص أدناه (اضغط مطولًا ← تحديد الكل ← نسخ) واحفظه في تطبيق الملاحظات قبل أي تحديث:
                  </p>
                  <textarea
                    readOnly
                    onClick={(e) => e.target.select()}
                    value={buildBackupJSON()}
                    dir="ltr"
                    className="w-full h-28 text-[10px] font-mono border border-[#D7DAD3] rounded-lg p-2"
                  />
                </div>
              )}

              <button
                onClick={() => {
                  setShowRestoreBox(!showRestoreBox);
                  if (!showRestoreBox) setShowBackupBox(false);
                  setRestoreMessage("");
                }}
                className="w-full py-2 rounded-lg text-xs font-bold bg-[#EAF0F8] text-[#4A6FA5] hover:bg-[#DDE9F5] mb-2"
              >
                {showRestoreBox ? "إخفاء الاسترجاع" : "📥 استرجاع نسخة احتياطية"}
              </button>
              {showRestoreBox && (
                <div className="mb-2">
                  <p className="text-[11px] text-[#5B6470] mb-1.5">الصق النص المحفوظ هنا:</p>
                  <textarea
                    value={restoreText}
                    onChange={(e) => setRestoreText(e.target.value)}
                    placeholder="الصق نص النسخة الاحتياطية هنا..."
                    dir="ltr"
                    className="w-full h-24 text-[10px] font-mono border border-[#D7DAD3] rounded-lg p-2 mb-1.5"
                  />
                  <button
                    onClick={restoreFromBackup}
                    className="w-full py-2 rounded-lg text-xs font-bold bg-[#4A6FA5] text-white hover:bg-[#3F5F8F]"
                  >
                    استرجاع الآن
                  </button>
                  {restoreMessage && (
                    <p className={`text-[11px] mt-1.5 ${restoreMessage.startsWith("✓") ? "text-[#3C7A5D]" : "text-[#B5473F]"}`}>
                      {restoreMessage}
                    </p>
                  )}
                </div>
              )}

              <button
                onClick={clearAllData}
                className="w-full py-2 rounded-lg text-xs font-bold bg-[#FBEAE8] text-[#B5473F] hover:bg-[#F5DAD6]"
              >
                {confirmClear ? "اضغط للتأكيد — لا يمكن التراجع" : "مسح جميع البيانات المخزّنة"}
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

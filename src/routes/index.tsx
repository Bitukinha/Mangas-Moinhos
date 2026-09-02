import { createFileRoute } from "@tanstack/react-router";
import { useServerFn } from "@tanstack/react-start";
import { useEffect, useMemo, useState } from "react";
import logo from "@/assets/nutrimilho-logo.png";
import {
  getRecords,
  addRecord as addRecordServerFn,
  deleteRecord as deleteRecordServerFn,
} from "@/server/records";
import {
  getMills,
  addMill as addMillServerFn,
  deleteMill as deleteMillServerFn,
} from "@/server/mills";
import {
  Bar as RBar,
  BarChart,
  CartesianGrid,
  Cell,
  Legend,
  Line,
  LineChart,
  Pie,
  PieChart,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis,
} from "recharts";
import * as XLSX from "xlsx";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";

export const Route = createFileRoute("/")({
  loader: async () => {
    const [records, mills] = await Promise.all([getRecords(), getMills()]);
    return { records, mills };
  },
  component: Index,
});

type Shift = "A" | "B" | "C";

type Mill = { id: string; name: string; area: string; mangas: number };

const SHIFTS: Shift[] = ["A", "B", "C"];

type Record = {
  id: string;
  millId: string;
  date: string; // YYYY-MM-DD
  shift: Shift;
  hour: string; // HH:MM
  responsavelLimpeza: string;
  responsavelMonitoramento: string;
  mangas: ("C" | "NC")[];
  createdAt: string;
};

function todayISO() {
  return new Date().toISOString().slice(0, 10);
}

function daysBetween(a: string, b: string) {
  const d1 = new Date(a + "T00:00:00").getTime();
  const d2 = new Date(b + "T00:00:00").getTime();
  return Math.max(1, Math.round((d2 - d1) / 86400000) + 1);
}

function Index() {
  const { records: initialRecords, mills: initialMills } = Route.useLoaderData();
  const [records, setRecords] = useState<Record[]>(initialRecords);
  const [mills, setMills] = useState<Mill[]>(initialMills);

  const addRecordFn = useServerFn(addRecordServerFn);
  const deleteRecordFn = useServerFn(deleteRecordServerFn);
  const addMillFn = useServerFn(addMillServerFn);
  const deleteMillFn = useServerFn(deleteMillServerFn);

  const [from, setFrom] = useState(() => {
    const d = new Date();
    d.setDate(d.getDate() - 6);
    return d.toISOString().slice(0, 10);
  });
  const [to, setTo] = useState(todayISO);
  const [tab, setTab] = useState<
    "dashboard" | "novo" | "historico" | "pendencias" | "moinhos"
  >("dashboard");

  const [millFilter, setMillFilter] = useState<string[]>(() =>
    initialMills.map((m) => m.id),
  );
  const [shiftFilter, setShiftFilter] = useState<Shift[]>(() => [...SHIFTS]);
  const [statusFilter, setStatusFilter] = useState<"all" | "C" | "NC">("all");

  const [drill, setDrill] = useState<{ title: string; recs: Record[] } | null>(null);
  const openDrill = (title: string, recs: Record[]) => setDrill({ title, recs });

  const toggleMill = (id: string) =>
    setMillFilter((prev) =>
      prev.includes(id) ? prev.filter((x) => x !== id) : [...prev, id],
    );
  const toggleShift = (s: Shift) =>
    setShiftFilter((prev) =>
      prev.includes(s) ? prev.filter((x) => x !== s) : [...prev, s],
    );
  const resetFilters = () => {
    setMillFilter(mills.map((m) => m.id));
    setShiftFilter([...SHIFTS]);
    setStatusFilter("all");
  };

  const filtered = useMemo(
    () =>
      records.filter((r) => {
        if (r.date < from || r.date > to) return false;
        if (!millFilter.includes(r.millId)) return false;
        if (!shiftFilter.includes(r.shift)) return false;
        if (statusFilter !== "all") {
          const isC = r.mangas.every((x) => x === "C");
          if (statusFilter === "C" && !isC) return false;
          if (statusFilter === "NC" && isC) return false;
        }
        return true;
      }),
    [records, from, to, millFilter, shiftFilter, statusFilter],
  );

  const activeMills = mills.filter((m) => millFilter.includes(m.id));
  const activeShifts = SHIFTS.filter((s) => shiftFilter.includes(s));
  const totalDias = daysBetween(from, to);
  const esperadoPorMoinho = totalDias * activeShifts.length;
  const esperadoTotal = esperadoPorMoinho * activeMills.length;

  const perMill = activeMills.map((m) => {
    const recs = filtered.filter((r) => r.millId === m.id);
    const feitos = recs.length;
    const conformes = recs.filter((r) => r.mangas.every((x) => x === "C")).length;
    const naoConformes = feitos - conformes;
    const aderencia = esperadoPorMoinho ? (feitos / esperadoPorMoinho) * 100 : 0;
    const conformidade = feitos ? (conformes / feitos) * 100 : 0;
    return { mill: m, feitos, conformes, naoConformes, aderencia, conformidade };
  });

  const totalFeitos = filtered.length;
  const totalConformes = filtered.filter((r) => r.mangas.every((x) => x === "C")).length;
  const aderenciaGeral = esperadoTotal ? (totalFeitos / esperadoTotal) * 100 : 0;
  const conformidadeGeral = totalFeitos ? (totalConformes / totalFeitos) * 100 : 0;

  // Série por data
  const perDate = useMemo(() => {
    const map = new Map<string, { date: string; feitos: number; conformes: number }>();
    for (let i = 0; i < totalDias; i++) {
      const d = new Date(from + "T00:00:00");
      d.setDate(d.getDate() + i);
      const iso = d.toISOString().slice(0, 10);
      map.set(iso, { date: iso, feitos: 0, conformes: 0 });
    }
    for (const r of filtered) {
      const row = map.get(r.date);
      if (!row) continue;
      row.feitos += 1;
      if (r.mangas.every((x) => x === "C")) row.conformes += 1;
    }
    const esperadoDia = activeMills.length * activeShifts.length;
    return Array.from(map.values()).map((row) => ({
      date: row.date.slice(5),
      feitos: row.feitos,
      conformes: row.conformes,
      naoConformes: row.feitos - row.conformes,
      aderencia: esperadoDia ? +((row.feitos / esperadoDia) * 100).toFixed(1) : 0,
      conformidade: row.feitos
        ? +((row.conformes / row.feitos) * 100).toFixed(1)
        : 0,
    }));
  }, [filtered, from, totalDias, activeMills.length, activeShifts.length]);

  // Por turno
  const perShift = activeShifts.map((s) => {
    const recs = filtered.filter((r) => r.shift === s);
    const feitos = recs.length;
    const conformes = recs.filter((r) => r.mangas.every((x) => x === "C")).length;
    const esperado = totalDias * activeMills.length;
    return {
      turno: `Turno ${s}`,
      feitos,
      conformes,
      naoConformes: feitos - conformes,
      aderencia: esperado ? +((feitos / esperado) * 100).toFixed(1) : 0,
      conformidade: feitos ? +((conformes / feitos) * 100).toFixed(1) : 0,
    };
  });

  // Pendências: combinações dia × turno × moinho sem registro no período
  const pendencias = useMemo(() => {
    const done = new Set(
      records
        .filter((r) => r.date >= from && r.date <= to)
        .filter((r) => millFilter.includes(r.millId))
        .filter((r) => shiftFilter.includes(r.shift))
        .map((r) => `${r.date}|${r.shift}|${r.millId}`),
    );
    const list: { date: string; shift: Shift; mill: Mill }[] = [];
    for (let i = 0; i < totalDias; i++) {
      const d = new Date(from + "T00:00:00");
      d.setDate(d.getDate() + i);
      const iso = d.toISOString().slice(0, 10);
      for (const s of activeShifts) {
        for (const m of activeMills) {
          if (!done.has(`${iso}|${s}|${m.id}`)) {
            list.push({ date: iso, shift: s, mill: m });
          }
        }
      }
    }
    return list.sort((a, b) => (b.date + b.shift).localeCompare(a.date + a.shift));
  }, [records, from, to, millFilter, shiftFilter, totalDias, activeShifts, activeMills]);

  const pendenciasPorTurno = activeShifts.map((s) => ({
    turno: `Turno ${s}`,
    pendencias: pendencias.filter((p) => p.shift === s).length,
  }));

  const pendenciasPorMoinho = activeMills.map((m) => ({
    moinho: m.name,
    pendencias: pendencias.filter((p) => p.mill.id === m.id).length,
  }));

  const pendenciasPorDia = useMemo(() => {
    const map = new Map<string, number>();
    for (const p of pendencias) map.set(p.date, (map.get(p.date) ?? 0) + 1);
    return Array.from(map.entries())
      .map(([date, count]) => ({ date, count }))
      .sort((a, b) => b.date.localeCompare(a.date));
  }, [pendencias]);

  const perMillChart = perMill.map((r) => ({
    moinho: r.mill.name,
    aderencia: +r.aderencia.toFixed(1),
    conformidade: +r.conformidade.toFixed(1),
    naoConformes: r.naoConformes,
  }));

  const pieData = [
    { name: "Conformes", value: totalConformes },
    { name: "Não conformes", value: totalFeitos - totalConformes },
    { name: "Não realizados", value: Math.max(0, esperadoTotal - totalFeitos) },
  ];

  const pieDrill = (name: string) => {
    if (name === "Conformes")
      return filtered.filter((r) => r.mangas.every((x) => x === "C"));
    if (name === "Não conformes")
      return filtered.filter((r) => r.mangas.some((x) => x === "NC"));
    return [];
  };

  const addRecord = (r: Record) => {
    setRecords((prev) => [r, ...prev]);
    addRecordFn({ data: r }).catch((err) => {
      console.error(err);
      setRecords((prev) => prev.filter((x) => x.id !== r.id));
    });
  };
  const removeRecord = (id: string) => {
    setRecords((prev) => prev.filter((r) => r.id !== id));
    deleteRecordFn({ data: { id } }).catch((err) => console.error(err));
  };

  const addMill = async (m: { name: string; area: string; mangas: number }) => {
    const created = await addMillFn({ data: m });
    setMills((prev) => [...prev, created]);
  };
  const removeMill = async (id: string) => {
    await deleteMillFn({ data: { id } });
    setMills((prev) => prev.filter((m) => m.id !== id));
    setMillFilter((prev) => prev.filter((x) => x !== id));
  };

  const exportExcel = () => {
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(
      wb,
      XLSX.utils.json_to_sheet(
        filtered.map((r) => {
          const m = mills.find((x) => x.id === r.millId);
          const nc = r.mangas.filter((x) => x === "NC").length;
          return {
            Data: r.date,
            Turno: r.shift,
            Hora: r.hour,
            Area: m?.area,
            Moinho: m?.name,
            "Mangás totais": r.mangas.length,
            "Mangás C": r.mangas.length - nc,
            "Mangás NC": nc,
            Status: nc === 0 ? "C" : "NC",
            "Resp. Limpeza": r.responsavelLimpeza,
            "Resp. Monitoramento": r.responsavelMonitoramento,
          };
        }),
      ),
      "Registros",
    );
    XLSX.utils.book_append_sheet(
      wb,
      XLSX.utils.json_to_sheet(
        perMill.map((r) => ({
          Moinho: r.mill.name,
          Area: r.mill.area,
          Mangas: r.mill.mangas,
          Feitos: r.feitos,
          Esperado: esperadoPorMoinho,
          Conformes: r.conformes,
          "Não Conformes": r.naoConformes,
          "Aderência %": +r.aderencia.toFixed(1),
          "Conformidade %": +r.conformidade.toFixed(1),
        })),
      ),
      "Por Moinho",
    );
    XLSX.utils.book_append_sheet(
      wb,
      XLSX.utils.json_to_sheet(perShift),
      "Por Turno",
    );
    XLSX.utils.book_append_sheet(
      wb,
      XLSX.utils.json_to_sheet(perDate),
      "Por Data",
    );
    XLSX.writeFile(wb, `aderencia_limpeza_${from}_a_${to}.xlsx`);
  };

  const exportPDF = () => {
    const doc = new jsPDF({ orientation: "landscape", unit: "pt", format: "a4" });
    // Sem cabeçalho ou rodapé — apenas as tabelas.
    let y = 30;
    autoTable(doc, {
      startY: y,
      head: [["Indicador", "Valor"]],
      body: [
        ["Período", `${from} a ${to}`],
        ["Registros feitos", `${totalFeitos} de ${esperadoTotal}`],
        ["Aderência à frequência", `${aderenciaGeral.toFixed(1)}%`],
        ["Conformidade", `${conformidadeGeral.toFixed(1)}%`],
        ["Não conformes", `${totalFeitos - totalConformes}`],
      ],
      styles: { fontSize: 9 },
      headStyles: { fillColor: [30, 90, 50] },
    });
    y = (doc as unknown as { lastAutoTable: { finalY: number } }).lastAutoTable.finalY + 20;
    autoTable(doc, {
      startY: y,
      head: [["Moinho", "Área", "Mangás", "Feitos", "Esperado", "Conformes", "NC", "Aderência %", "Conformidade %"]],
      body: perMill.map((r) => [
        r.mill.name,
        r.mill.area,
        r.mill.mangas,
        r.feitos,
        esperadoPorMoinho,
        r.conformes,
        r.naoConformes,
        r.aderencia.toFixed(1),
        r.conformidade.toFixed(1),
      ]),
      styles: { fontSize: 9 },
      headStyles: { fillColor: [30, 90, 50] },
    });
    y = (doc as unknown as { lastAutoTable: { finalY: number } }).lastAutoTable.finalY + 20;
    autoTable(doc, {
      startY: y,
      head: [["Turno", "Feitos", "Conformes", "NC", "Aderência %", "Conformidade %"]],
      body: perShift.map((r) => [
        r.turno,
        r.feitos,
        r.conformes,
        r.naoConformes,
        r.aderencia,
        r.conformidade,
      ]),
      styles: { fontSize: 9 },
      headStyles: { fillColor: [30, 90, 50] },
    });
    y = (doc as unknown as { lastAutoTable: { finalY: number } }).lastAutoTable.finalY + 20;
    autoTable(doc, {
      startY: y,
      head: [["Data", "Feitos", "Conformes", "NC", "Aderência %", "Conformidade %"]],
      body: perDate.map((r) => [
        r.date,
        r.feitos,
        r.conformes,
        r.naoConformes,
        r.aderencia,
        r.conformidade,
      ]),
      styles: { fontSize: 9 },
      headStyles: { fillColor: [30, 90, 50] },
    });
    const registrosOrdenados = [...filtered].sort((a, b) =>
      (b.date + b.hour).localeCompare(a.date + a.hour),
    );
    y = (doc as unknown as { lastAutoTable: { finalY: number } }).lastAutoTable.finalY + 20;
    autoTable(doc, {
      startY: y,
      head: [
        [
          "Data",
          "Turno",
          "Hora",
          "Moinho",
          "Área",
          "Mangás C",
          "Mangás NC",
          "Status",
          "Resp. Limpeza",
          "Resp. Monitoramento",
        ],
      ],
      body: registrosOrdenados.map((r) => {
        const m = mills.find((x) => x.id === r.millId);
        const nc = r.mangas.filter((x) => x === "NC").length;
        return [
          r.date,
          r.shift,
          r.hour,
          m?.name ?? "",
          m?.area ?? "",
          r.mangas.length - nc,
          nc,
          nc === 0 ? "C" : "NC",
          r.responsavelLimpeza,
          r.responsavelMonitoramento,
        ];
      }),
      styles: { fontSize: 8 },
      headStyles: { fillColor: [30, 90, 50] },
      didParseCell: (data) => {
        if (data.section === "body" && data.column.index === 7) {
          data.cell.styles.textColor = data.cell.raw === "NC" ? [180, 40, 40] : [30, 120, 60];
          data.cell.styles.fontStyle = "bold";
        }
      },
    });
    doc.save(`aderencia_limpeza_${from}_a_${to}.pdf`);
  };

  return (
    <div className="min-h-screen bg-background text-foreground">
      <header className="border-b border-border bg-primary text-primary-foreground">
        <div className="mx-auto flex max-w-6xl flex-col gap-3 px-4 py-3 sm:flex-row sm:items-center sm:justify-between sm:gap-4 sm:px-6 sm:py-4">
          <div className="flex items-center gap-3">
            <div className="shrink-0 rounded-md bg-white/95 px-3 py-2">
              <img src={logo} alt="Nutrimilho" className="h-8 w-auto shrink-0" />
            </div>
            <div className="min-w-0">
              <h1 className="text-base font-semibold leading-tight sm:text-lg">
                Aderência — Limpeza de Mangas
              </h1>
              <p className="text-xs opacity-80">
                FOR-RQ-001.07 · Moinhos Nutrimilho
              </p>
            </div>
          </div>
          <nav className="flex gap-1 rounded-lg bg-white/10 p-1 text-sm">
            {(
              [
                ["dashboard", "Indicador"],
                ["novo", "Registrar"],
                ["historico", "Histórico"],
                ["pendencias", "Pendências"],
                ["moinhos", "Moinhos"],
              ] as const
            ).map(([k, l]) => (
              <button
                key={k}
                onClick={() => setTab(k)}
                className={`flex-1 rounded-md px-3 py-1.5 transition sm:flex-initial ${
                  tab === k
                    ? "bg-white text-primary shadow"
                    : "text-primary-foreground/90 hover:bg-white/10"
                }`}
              >
                {l}
              </button>
            ))}
          </nav>
        </div>
      </header>

      <main className="mx-auto max-w-6xl px-6 py-8">
        {tab === "dashboard" && (
          <section className="space-y-6">
            <div className="flex flex-wrap items-end gap-3">
              <div>
                <label className="block text-xs font-medium text-muted-foreground">De</label>
                <input
                  type="date"
                  value={from}
                  onChange={(e) => setFrom(e.target.value)}
                  className="mt-1 rounded-md border border-input bg-card px-3 py-2 text-sm"
                />
              </div>
              <div>
                <label className="block text-xs font-medium text-muted-foreground">Até</label>
                <input
                  type="date"
                  value={to}
                  onChange={(e) => setTo(e.target.value)}
                  className="mt-1 rounded-md border border-input bg-card px-3 py-2 text-sm"
                />
              </div>
              <p className="ml-auto text-xs text-muted-foreground">
                Esperado: {activeShifts.length} turno(s)/dia × {activeMills.length} moinho(s) ={" "}
                <strong>{esperadoTotal}</strong> registros
              </p>
              <div className="flex gap-2">
                <button
                  onClick={exportExcel}
                  className="rounded-md border border-input bg-card px-3 py-2 text-sm font-medium hover:bg-muted"
                >
                  ↓ Excel
                </button>
                <button
                  onClick={exportPDF}
                  className="rounded-md bg-primary px-3 py-2 text-sm font-medium text-primary-foreground hover:opacity-90"
                >
                  ↓ PDF
                </button>
              </div>
            </div>

            <div className="rounded-xl border border-border bg-card p-4">
              <div className="mb-3 flex items-center justify-between">
                <h3 className="text-sm font-semibold">Filtros avançados</h3>
                <button
                  onClick={resetFilters}
                  className="text-xs text-muted-foreground underline hover:text-foreground"
                >
                  Limpar filtros
                </button>
              </div>
              <div className="grid gap-4 md:grid-cols-3">
                <div>
                  <p className="mb-2 text-xs font-medium text-muted-foreground">
                    Moinhos ({millFilter.length}/{mills.length})
                  </p>
                  <div className="flex flex-wrap gap-1.5">
                    {mills.map((m) => {
                      const on = millFilter.includes(m.id);
                      return (
                        <button
                          key={m.id}
                          type="button"
                          onClick={() => toggleMill(m.id)}
                          className={`rounded-full border px-2.5 py-1 text-xs font-medium transition ${
                            on
                              ? "border-primary bg-primary text-primary-foreground"
                              : "border-input bg-background text-muted-foreground hover:bg-muted"
                          }`}
                          title={m.area}
                        >
                          {m.name}
                        </button>
                      );
                    })}
                  </div>
                </div>
                <div>
                  <p className="mb-2 text-xs font-medium text-muted-foreground">Turnos</p>
                  <div className="flex gap-1.5">
                    {SHIFTS.map((s) => {
                      const on = shiftFilter.includes(s);
                      return (
                        <button
                          key={s}
                          type="button"
                          onClick={() => toggleShift(s)}
                          className={`flex-1 rounded-md border px-3 py-1.5 text-xs font-semibold transition ${
                            on
                              ? "border-primary bg-primary text-primary-foreground"
                              : "border-input bg-background text-muted-foreground hover:bg-muted"
                          }`}
                        >
                          Turno {s}
                        </button>
                      );
                    })}
                  </div>
                </div>
                <div>
                  <p className="mb-2 text-xs font-medium text-muted-foreground">Status</p>
                  <div className="flex gap-1.5">
                    {(
                      [
                        ["all", "Todos"],
                        ["C", "Conforme"],
                        ["NC", "Não Conforme"],
                      ] as const
                    ).map(([k, l]) => (
                      <button
                        key={k}
                        type="button"
                        onClick={() => setStatusFilter(k)}
                        className={`flex-1 rounded-md border px-2 py-1.5 text-xs font-semibold transition ${
                          statusFilter === k
                            ? k === "NC"
                              ? "border-destructive bg-destructive text-destructive-foreground"
                              : k === "C"
                                ? "border-success bg-success text-success-foreground"
                                : "border-primary bg-primary text-primary-foreground"
                            : "border-input bg-background text-muted-foreground hover:bg-muted"
                        }`}
                      >
                        {l}
                      </button>
                    ))}
                  </div>
                </div>
              </div>
              <p className="mt-3 text-xs text-muted-foreground">
                {filtered.length} registro(s) após filtros · exportações Excel/PDF respeitam os filtros aplicados.
              </p>
            </div>

            <div className="grid gap-4 md:grid-cols-3">
              <KpiCard
                title="Aderência à Frequência"
                value={`${aderenciaGeral.toFixed(1)}%`}
                sub={`${totalFeitos} de ${esperadoTotal} registros`}
                tone={aderenciaGeral >= 90 ? "good" : aderenciaGeral >= 70 ? "warn" : "bad"}
                onClick={() => openDrill(`Registros feitos — ${from} a ${to}`, filtered)}
              />
              <KpiCard
                title="Conformidade"
                value={`${conformidadeGeral.toFixed(1)}%`}
                sub={`${totalConformes} conformes / ${totalFeitos} registros`}
                tone={
                  conformidadeGeral >= 95 ? "good" : conformidadeGeral >= 80 ? "warn" : "bad"
                }
                onClick={() =>
                  openDrill(
                    "Registros conformes",
                    filtered.filter((r) => r.mangas.every((x) => x === "C")),
                  )
                }
              />
              <KpiCard
                title="Não Conformes"
                value={`${totalFeitos - totalConformes}`}
                sub={`no período (${totalDias} ${totalDias === 1 ? "dia" : "dias"})`}
                tone={totalFeitos - totalConformes === 0 ? "good" : "bad"}
                onClick={() =>
                  openDrill(
                    "Registros não conformes",
                    filtered.filter((r) => r.mangas.some((x) => x === "NC")),
                  )
                }
              />
            </div>

            <div className="overflow-x-auto rounded-xl border border-border bg-card">
              <table className="w-full text-sm">
                <thead className="bg-secondary text-secondary-foreground">
                  <tr>
                    <th className="px-4 py-2 text-left">Moinho</th>
                    <th className="px-4 py-2 text-left">Área</th>
                    <th className="px-4 py-2 text-center">Mangás</th>
                    <th className="px-4 py-2 text-center">Feitos / Esperado</th>
                    <th className="px-4 py-2 text-left">Aderência</th>
                    <th className="px-4 py-2 text-left">Conformidade</th>
                  </tr>
                </thead>
                <tbody>
                  {perMill.map((row) => (
                    <tr
                      key={row.mill.id}
                      className="cursor-pointer border-t border-border hover:bg-muted/60"
                      onClick={() =>
                        openDrill(
                          `Moinho ${row.mill.name} — ${row.feitos} registro(s)`,
                          filtered.filter((r) => r.millId === row.mill.id),
                        )
                      }
                      title="Ver registros deste moinho"
                    >
                      <td className="px-4 py-3 font-medium">{row.mill.name}</td>
                      <td className="px-4 py-3 text-muted-foreground">{row.mill.area}</td>
                      <td className="px-4 py-3 text-center">{row.mill.mangas}</td>
                      <td className="px-4 py-3 text-center">
                        {row.feitos} / {esperadoPorMoinho}
                      </td>
                      <td className="px-4 py-3">
                        <Bar value={row.aderencia} />
                      </td>
                      <td className="px-4 py-3">
                        <Bar value={row.conformidade} tone="conf" />
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            <div className="grid gap-4 lg:grid-cols-2">
              <ChartCard title="Tendência diária — Aderência & Conformidade">
                <ResponsiveContainer width="100%" height={260}>
                  <LineChart
                    data={perDate}
                    margin={{ top: 10, right: 16, bottom: 0, left: -10 }}
                    onClick={(e: unknown) => {
                      const label = (e as { activeLabel?: string })?.activeLabel;
                      if (!label) return;
                      const iso = `${from.slice(0, 4)}-${label}`;
                      openDrill(
                        `Registros de ${label}`,
                        filtered.filter((r) => r.date === iso),
                      );
                    }}
                  >
                    <CartesianGrid strokeDasharray="3 3" stroke="hsl(0 0% 90%)" />
                    <XAxis dataKey="date" fontSize={11} />
                    <YAxis domain={[0, 100]} fontSize={11} unit="%" />
                    <Tooltip />
                    <Legend wrapperStyle={{ fontSize: 12 }} />
                    <Line type="monotone" dataKey="aderencia" name="Aderência %" stroke="oklch(0.38 0.13 145)" strokeWidth={2} style={{ cursor: "pointer" }} />
                    <Line type="monotone" dataKey="conformidade" name="Conformidade %" stroke="oklch(0.7 0.18 85)" strokeWidth={2} style={{ cursor: "pointer" }} />
                  </LineChart>
                </ResponsiveContainer>
              </ChartCard>

              <ChartCard title="Registros por data (feitos vs. não conformes)">
                <ResponsiveContainer width="100%" height={260}>
                  <BarChart
                    data={perDate}
                    margin={{ top: 10, right: 16, bottom: 0, left: -10 }}
                    onClick={(e: unknown) => {
                      const label = (e as { activeLabel?: string })?.activeLabel;
                      if (!label) return;
                      const iso = `${from.slice(0, 4)}-${label}`;
                      openDrill(
                        `Registros de ${label}`,
                        filtered.filter((r) => r.date === iso),
                      );
                    }}
                  >
                    <CartesianGrid strokeDasharray="3 3" stroke="hsl(0 0% 90%)" />
                    <XAxis dataKey="date" fontSize={11} />
                    <YAxis fontSize={11} />
                    <Tooltip />
                    <Legend wrapperStyle={{ fontSize: 12 }} />
                    <RBar dataKey="conformes" name="Conformes" stackId="a" fill="oklch(0.62 0.17 145)" style={{ cursor: "pointer" }} />
                    <RBar dataKey="naoConformes" name="Não conformes" stackId="a" fill="oklch(0.6 0.22 27)" style={{ cursor: "pointer" }} />
                  </BarChart>
                </ResponsiveContainer>
              </ChartCard>

              <ChartCard title="Conformidade e aderência por moinho">
                <ResponsiveContainer width="100%" height={260}>
                  <BarChart
                    data={perMillChart}
                    margin={{ top: 10, right: 16, bottom: 0, left: -10 }}
                    onClick={(e: unknown) => {
                      const label = (e as { activeLabel?: string })?.activeLabel;
                      if (!label) return;
                      const mill = mills.find((m) => m.name === label);
                      if (!mill) return;
                      openDrill(
                        `Moinho ${mill.name}`,
                        filtered.filter((r) => r.millId === mill.id),
                      );
                    }}
                  >
                    <CartesianGrid strokeDasharray="3 3" stroke="hsl(0 0% 90%)" />
                    <XAxis dataKey="moinho" fontSize={11} />
                    <YAxis domain={[0, 100]} fontSize={11} unit="%" />
                    <Tooltip />
                    <Legend wrapperStyle={{ fontSize: 12 }} />
                    <RBar dataKey="aderencia" name="Aderência %" fill="oklch(0.38 0.13 145)" style={{ cursor: "pointer" }} />
                    <RBar dataKey="conformidade" name="Conformidade %" fill="oklch(0.7 0.18 85)" style={{ cursor: "pointer" }} />
                  </BarChart>
                </ResponsiveContainer>
              </ChartCard>

              <ChartCard title="Por turno (A / B / C)">
                <div className="grid grid-cols-2 gap-2">
                  <ResponsiveContainer width="100%" height={240}>
                    <BarChart
                      data={perShift}
                      margin={{ top: 10, right: 8, bottom: 0, left: -10 }}
                      onClick={(e: unknown) => {
                        const label = (e as { activeLabel?: string })?.activeLabel;
                        if (!label) return;
                        const s = label.replace("Turno ", "") as Shift;
                        openDrill(
                          `Turno ${s}`,
                          filtered.filter((r) => r.shift === s),
                        );
                      }}
                    >
                      <CartesianGrid strokeDasharray="3 3" stroke="hsl(0 0% 90%)" />
                      <XAxis dataKey="turno" fontSize={11} />
                      <YAxis domain={[0, 100]} fontSize={11} unit="%" />
                      <Tooltip />
                      <Legend wrapperStyle={{ fontSize: 11 }} />
                      <RBar dataKey="aderencia" name="Aderência %" fill="oklch(0.38 0.13 145)" style={{ cursor: "pointer" }} />
                      <RBar dataKey="conformidade" name="Conformidade %" fill="oklch(0.7 0.18 85)" style={{ cursor: "pointer" }} />
                    </BarChart>
                  </ResponsiveContainer>
                  <ResponsiveContainer width="100%" height={240}>
                    <PieChart>
                      <Tooltip />
                      <Legend wrapperStyle={{ fontSize: 11 }} />
                      <Pie
                        data={pieData}
                        dataKey="value"
                        nameKey="name"
                        outerRadius={80}
                        label
                        style={{ cursor: "pointer" }}
                        onClick={(d: unknown) => {
                          const name = (d as { name?: string })?.name;
                          if (!name) return;
                          const recs = pieDrill(name);
                          if (recs.length === 0 && name === "Não realizados") return;
                          openDrill(name, recs);
                        }}
                      >
                        {pieData.map((_, i) => (
                          <Cell
                            key={i}
                            fill={
                              ["oklch(0.62 0.17 145)", "oklch(0.6 0.22 27)", "oklch(0.75 0.05 130)"][i]
                            }
                          />
                        ))}
                      </Pie>
                    </PieChart>
                  </ResponsiveContainer>
                </div>
              </ChartCard>
            </div>
          </section>
        )}

        {tab === "novo" && <NewRecordForm mills={mills} onAdd={addRecord} />}

        {tab === "historico" && (
          <HistoryTable records={filtered} mills={mills} onDelete={removeRecord} />
        )}

        {tab === "pendencias" && (
          <PendenciasTab
            from={from}
            to={to}
            pendencias={pendencias}
            porTurno={pendenciasPorTurno}
            porMoinho={pendenciasPorMoinho}
            porDia={pendenciasPorDia}
          />
        )}

        {tab === "moinhos" && (
          <MillsManager mills={mills} onAdd={addMill} onDelete={removeMill} />
        )}
      </main>

      <footer className="mx-auto max-w-6xl px-6 py-6 text-center text-xs text-muted-foreground">
        © {new Date().getFullYear()} Nutrimilho - (Novaes Tech) | Todos os direitos reservados
      </footer>

      {drill && (
        <DrillModal
          title={drill.title}
          records={drill.recs}
          mills={mills}
          onClose={() => setDrill(null)}
        />
      )}
    </div>
  );
}

function KpiCard({
  title,
  value,
  sub,
  tone,
  onClick,
}: {
  title: string;
  value: string;
  sub: string;
  tone: "good" | "warn" | "bad";
  onClick?: () => void;
}) {
  const color =
    tone === "good"
      ? "text-success"
      : tone === "warn"
        ? "text-accent-foreground"
        : "text-destructive";
  const bg =
    tone === "good"
      ? "bg-success/10"
      : tone === "warn"
        ? "bg-accent/30"
        : "bg-destructive/10";
  return (
    <div
      className={`rounded-xl border border-border p-5 ${bg} ${onClick ? "cursor-pointer transition hover:shadow-md" : ""}`}
      onClick={onClick}
      role={onClick ? "button" : undefined}
      title={onClick ? "Ver registros" : undefined}
    >
      <p className="text-xs font-medium uppercase tracking-wide text-muted-foreground">
        {title}
      </p>
      <p className={`mt-2 text-3xl font-bold ${color}`}>{value}</p>
      <p className="mt-1 text-xs text-muted-foreground">{sub}</p>
      {onClick && (
        <p className="mt-2 text-[10px] uppercase tracking-wide text-muted-foreground/70">
          Clique para detalhar →
        </p>
      )}
    </div>
  );
}

function ChartCard({ title, children }: { title: string; children: React.ReactNode }) {
  return (
    <div className="rounded-xl border border-border bg-card p-4">
      <h3 className="mb-3 text-sm font-semibold text-foreground">{title}</h3>
      {children}
    </div>
  );
}

function Bar({ value, tone = "ader" }: { value: number; tone?: "ader" | "conf" }) {
  const v = Math.min(100, Math.max(0, value));
  const good = tone === "ader" ? v >= 90 : v >= 95;
  const mid = tone === "ader" ? v >= 70 : v >= 80;
  const color = good ? "bg-success" : mid ? "bg-accent" : "bg-destructive";
  return (
    <div className="flex items-center gap-2">
      <div className="h-2 flex-1 overflow-hidden rounded-full bg-muted">
        <div className={`h-full ${color}`} style={{ width: `${v}%` }} />
      </div>
      <span className="w-12 text-right text-xs font-medium tabular-nums">
        {v.toFixed(0)}%
      </span>
    </div>
  );
}

function NewRecordForm({
  mills,
  onAdd,
}: {
  mills: Mill[];
  onAdd: (r: Record) => void;
}) {
  const [millId, setMillId] = useState(mills[0]?.id ?? "");
  const mill = mills.find((m) => m.id === millId) ?? mills[0];
  const [date, setDate] = useState(todayISO);
  const [shift, setShift] = useState<Shift>("A");
  const [hour, setHour] = useState("");
  const [respLimpeza, setRespLimpeza] = useState("");
  const [respMonit, setRespMonit] = useState("");
  const [mangas, setMangas] = useState<("C" | "NC")[]>(() =>
    Array(mill?.mangas ?? 0).fill("C"),
  );
  const [saved, setSaved] = useState(false);

  useEffect(() => {
    setMangas(Array(mill?.mangas ?? 0).fill("C"));
  }, [mill?.mangas]);

  const submit = (e: React.FormEvent) => {
    e.preventDefault();
    if (!mill || !respLimpeza || !respMonit || !hour) return;
    onAdd({
      id: crypto.randomUUID(),
      millId: mill.id,
      date,
      shift,
      hour,
      responsavelLimpeza: respLimpeza,
      responsavelMonitoramento: respMonit,
      mangas,
      createdAt: new Date().toISOString(),
    });
    setHour("");
    setRespLimpeza("");
    setRespMonit("");
    setMangas(Array(mill.mangas).fill("C"));
    setSaved(true);
    setTimeout(() => setSaved(false), 2000);
  };

  const allC = mangas.every((x) => x === "C");

  if (!mill) {
    return (
      <div className="mx-auto max-w-3xl rounded-xl border border-dashed border-border bg-card p-10 text-center text-sm text-muted-foreground">
        Nenhum moinho cadastrado. Cadastre um moinho na aba "Moinhos" antes de
        registrar uma limpeza.
      </div>
    );
  }

  return (
    <form
      onSubmit={submit}
      className="mx-auto max-w-3xl space-y-6 rounded-xl border border-border bg-card p-6 shadow-sm"
    >
      <div>
        <h2 className="text-lg font-semibold">Novo registro de limpeza</h2>
        <p className="text-sm text-muted-foreground">
          1x por turno (cada final de turno). Marque cada mangá como C (conforme) ou NC.
        </p>
      </div>

      <div className="grid gap-4 md:grid-cols-2">
        <Field label="Moinho">
          <select
            value={millId}
            onChange={(e) => setMillId(e.target.value)}
            className="w-full rounded-md border border-input bg-background px-3 py-2 text-sm"
          >
            {mills.map((m) => (
              <option key={m.id} value={m.id}>
                {m.area} — {m.name} ({m.mangas} mangás)
              </option>
            ))}
          </select>
        </Field>
        <Field label="Data">
          <input
            type="date"
            value={date}
            onChange={(e) => setDate(e.target.value)}
            className="w-full rounded-md border border-input bg-background px-3 py-2 text-sm"
          />
        </Field>
        <Field label="Turno">
          <div className="flex gap-2">
            {SHIFTS.map((s) => (
              <button
                key={s}
                type="button"
                onClick={() => setShift(s)}
                className={`flex-1 rounded-md border px-3 py-2 text-sm font-medium transition ${
                  shift === s
                    ? "border-primary bg-primary text-primary-foreground"
                    : "border-input bg-background hover:bg-muted"
                }`}
              >
                Turno {s}
              </button>
            ))}
          </div>
        </Field>
        <Field label="Hora">
          <input
            type="time"
            value={hour}
            onChange={(e) => setHour(e.target.value)}
            required
            className="w-full rounded-md border border-input bg-background px-3 py-2 text-sm"
          />
        </Field>
        <Field label="Responsável pela limpeza">
          <input
            value={respLimpeza}
            onChange={(e) => setRespLimpeza(e.target.value)}
            required
            placeholder="Nome"
            className="w-full rounded-md border border-input bg-background px-3 py-2 text-sm"
          />
        </Field>
        <Field label="Responsável pelo monitoramento">
          <input
            value={respMonit}
            onChange={(e) => setRespMonit(e.target.value)}
            required
            placeholder="Nome"
            className="w-full rounded-md border border-input bg-background px-3 py-2 text-sm"
          />
        </Field>
      </div>

      <div>
        <div className="mb-2 flex items-center justify-between">
          <label className="text-sm font-medium">
            Status das {mill.mangas} mangás
          </label>
          <div className="flex gap-2 text-xs">
            <button
              type="button"
              onClick={() => setMangas(Array(mill.mangas).fill("C"))}
              className="rounded border border-input px-2 py-1 hover:bg-muted"
            >
              Todas C
            </button>
            <button
              type="button"
              onClick={() => setMangas(Array(mill.mangas).fill("NC"))}
              className="rounded border border-input px-2 py-1 hover:bg-muted"
            >
              Todas NC
            </button>
          </div>
        </div>
        <div className="grid grid-cols-2 gap-2 sm:grid-cols-5">
          {mangas.map((v, i) => (
            <button
              key={i}
              type="button"
              onClick={() =>
                setMangas((prev) =>
                  prev.map((x, idx) => (idx === i ? (x === "C" ? "NC" : "C") : x)),
                )
              }
              className={`rounded-lg border p-3 text-center text-sm font-semibold transition ${
                v === "C"
                  ? "border-success/40 bg-success/10 text-success"
                  : "border-destructive/40 bg-destructive/10 text-destructive"
              }`}
            >
              <div className="text-xs opacity-70">Mangá {i + 1}</div>
              <div className="mt-1 text-lg">{v}</div>
            </button>
          ))}
        </div>
        <p className="mt-2 text-xs text-muted-foreground">
          Este registro será: <strong>{allC ? "CONFORME" : "NÃO CONFORME"}</strong>
        </p>
      </div>

      <div className="flex items-center gap-3">
        <button
          type="submit"
          className="rounded-md bg-primary px-5 py-2.5 text-sm font-semibold text-primary-foreground shadow hover:opacity-90"
        >
          Salvar registro
        </button>
        {saved && <span className="text-sm text-success">✓ Registro salvo</span>}
      </div>
    </form>
  );
}

function Field({ label, children }: { label: string; children: React.ReactNode }) {
  return (
    <div>
      <label className="mb-1 block text-xs font-medium text-muted-foreground">
        {label}
      </label>
      {children}
    </div>
  );
}

function HistoryTable({
  records,
  mills,
  onDelete,
}: {
  records: Record[];
  mills: Mill[];
  onDelete: (id: string) => void;
}) {
  return <HistoryTableImpl records={records} mills={mills} onDelete={onDelete} />;
}

function DrillModal({
  title,
  records,
  mills,
  onClose,
}: {
  title: string;
  records: Record[];
  mills: Mill[];
  onClose: () => void;
}) {
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key === "Escape") onClose();
    };
    window.addEventListener("keydown", onKey);
    document.body.style.overflow = "hidden";
    return () => {
      window.removeEventListener("keydown", onKey);
      document.body.style.overflow = "";
    };
  }, [onClose]);

  const sorted = [...records].sort((a, b) =>
    (b.date + b.hour).localeCompare(a.date + a.hour),
  );
  const totalNC = sorted.filter((r) => r.mangas.some((x) => x === "NC")).length;

  return (
    <div
      className="fixed inset-0 z-50 flex items-center justify-center bg-black/50 p-4"
      onClick={onClose}
    >
      <div
        className="max-h-[90vh] w-full max-w-5xl overflow-hidden rounded-xl border border-border bg-card shadow-xl"
        onClick={(e) => e.stopPropagation()}
      >
        <div className="flex items-center justify-between border-b border-border bg-primary px-5 py-3 text-primary-foreground">
          <div>
            <h3 className="text-base font-semibold">{title}</h3>
            <p className="text-xs opacity-80">
              {sorted.length} registro(s) · {totalNC} não conforme(s)
            </p>
          </div>
          <button
            onClick={onClose}
            className="rounded-md bg-white/10 px-3 py-1.5 text-sm hover:bg-white/20"
          >
            Fechar ✕
          </button>
        </div>
        <div className="max-h-[calc(90vh-64px)] overflow-auto">
          {sorted.length === 0 ? (
            <div className="p-10 text-center text-sm text-muted-foreground">
              Nenhum registro encontrado.
            </div>
          ) : (
            <table className="w-full text-sm">
              <thead className="sticky top-0 bg-secondary text-secondary-foreground">
                <tr>
                  <th className="px-3 py-2 text-left">Data</th>
                  <th className="px-3 py-2 text-left">Turno</th>
                  <th className="px-3 py-2 text-left">Hora</th>
                  <th className="px-3 py-2 text-left">Moinho</th>
                  <th className="px-3 py-2 text-left">Mangás</th>
                  <th className="px-3 py-2 text-left">Status</th>
                  <th className="px-3 py-2 text-left">Limpeza</th>
                  <th className="px-3 py-2 text-left">Monitoramento</th>
                </tr>
              </thead>
              <tbody>
                {sorted.map((r) => {
                  const m = mills.find((x) => x.id === r.millId);
                  const ncIdx = r.mangas
                    .map((x, i) => (x === "NC" ? i + 1 : null))
                    .filter((x): x is number => x !== null);
                  const isC = ncIdx.length === 0;
                  return (
                    <tr key={r.id} className="border-t border-border align-top">
                      <td className="px-3 py-2">{r.date}</td>
                      <td className="px-3 py-2">{r.shift}</td>
                      <td className="px-3 py-2">{r.hour}</td>
                      <td className="px-3 py-2">
                        <div className="font-medium">{m?.name}</div>
                        <div className="text-xs text-muted-foreground">{m?.area}</div>
                      </td>
                      <td className="px-3 py-2">
                        <div className="flex flex-wrap gap-1">
                          {r.mangas.map((v, i) => (
                            <span
                              key={i}
                              title={`Mangá ${i + 1}: ${v}`}
                              className={`inline-flex h-6 min-w-[24px] items-center justify-center rounded px-1 text-[10px] font-semibold ${
                                v === "C"
                                  ? "bg-success/15 text-success"
                                  : "bg-destructive/15 text-destructive"
                              }`}
                            >
                              {i + 1}
                            </span>
                          ))}
                        </div>
                        {ncIdx.length > 0 && (
                          <div className="mt-1 text-[10px] text-destructive">
                            NC: mangá {ncIdx.join(", ")}
                          </div>
                        )}
                      </td>
                      <td className="px-3 py-2">
                        <span
                          className={`rounded px-2 py-0.5 text-xs font-semibold ${
                            isC
                              ? "bg-success/15 text-success"
                              : "bg-destructive/15 text-destructive"
                          }`}
                        >
                          {isC ? "C" : "NC"}
                        </span>
                      </td>
                      <td className="px-3 py-2">{r.responsavelLimpeza}</td>
                      <td className="px-3 py-2">{r.responsavelMonitoramento}</td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          )}
        </div>
      </div>
    </div>
  );
}

function HistoryTableImpl({
  records,
  mills,
  onDelete,
}: {
  records: Record[];
  mills: Mill[];
  onDelete: (id: string) => void;
}) {
  if (records.length === 0) {
    return (
      <div className="rounded-xl border border-dashed border-border bg-card p-10 text-center text-sm text-muted-foreground">
        Nenhum registro no período selecionado.
      </div>
    );
  }
  const sorted = [...records].sort((a, b) =>
    (b.date + b.hour).localeCompare(a.date + a.hour),
  );
  return (
    <div className="overflow-x-auto rounded-xl border border-border bg-card">
      <table className="w-full text-sm">
        <thead className="bg-secondary text-secondary-foreground">
          <tr>
            <th className="px-3 py-2 text-left">Data</th>
            <th className="px-3 py-2 text-left">Turno</th>
            <th className="px-3 py-2 text-left">Hora</th>
            <th className="px-3 py-2 text-left">Moinho</th>
            <th className="px-3 py-2 text-left">Limpeza</th>
            <th className="px-3 py-2 text-left">Monitoramento</th>
            <th className="px-3 py-2 text-center">Mangás</th>
            <th className="px-3 py-2 text-center">Status</th>
            <th className="px-3 py-2"></th>
          </tr>
        </thead>
        <tbody>
          {sorted.map((r) => {
            const mill = mills.find((m) => m.id === r.millId);
            const ncCount = r.mangas.filter((x) => x === "NC").length;
            const conforme = ncCount === 0;
            return (
              <tr key={r.id} className="border-t border-border">
                <td className="px-3 py-2">{r.date}</td>
                <td className="px-3 py-2">{r.shift}</td>
                <td className="px-3 py-2">{r.hour}</td>
                <td className="px-3 py-2">
                  {mill?.area} — {mill?.name}
                </td>
                <td className="px-3 py-2">{r.responsavelLimpeza}</td>
                <td className="px-3 py-2">{r.responsavelMonitoramento}</td>
                <td className="px-3 py-2 text-center text-xs">
                  {r.mangas.length - ncCount} C / {ncCount} NC
                </td>
                <td className="px-3 py-2 text-center">
                  <span
                    className={`inline-block rounded-full px-2 py-0.5 text-xs font-semibold ${
                      conforme
                        ? "bg-success/15 text-success"
                        : "bg-destructive/15 text-destructive"
                    }`}
                  >
                    {conforme ? "C" : "NC"}
                  </span>
                </td>
                <td className="px-3 py-2 text-right">
                  <button
                    onClick={() => onDelete(r.id)}
                    className="text-xs text-muted-foreground hover:text-destructive"
                  >
                    Excluir
                  </button>
                </td>
              </tr>
            );
          })}
        </tbody>
      </table>
    </div>
  );
}

function PendenciasTab({
  from,
  to,
  pendencias,
  porTurno,
  porMoinho,
  porDia,
}: {
  from: string;
  to: string;
  pendencias: { date: string; shift: Shift; mill: Mill }[];
  porTurno: { turno: string; pendencias: number }[];
  porMoinho: { moinho: string; pendencias: number }[];
  porDia: { date: string; count: number }[];
}) {
  return (
    <section className="space-y-6">
      <div>
        <h2 className="text-lg font-semibold">
          Pendências — limpezas não registradas
        </h2>
        <p className="text-sm text-muted-foreground">
          Combinações de dia, turno e moinho sem registro de limpeza entre{" "}
          {from} e {to}.
        </p>
      </div>

      <div className="grid gap-4 md:grid-cols-3">
        <KpiCard
          title="Total de Pendências"
          value={`${pendencias.length}`}
          sub="registros de limpeza faltando no período"
          tone={pendencias.length === 0 ? "good" : "bad"}
        />
        <div className="rounded-xl border border-border bg-card p-4">
          <p className="mb-2 text-xs font-medium text-muted-foreground">
            Por turno
          </p>
          <div className="space-y-1.5">
            {porTurno.map((t) => (
              <div key={t.turno} className="flex items-center justify-between text-sm">
                <span>{t.turno}</span>
                <span
                  className={`font-semibold ${t.pendencias === 0 ? "text-success" : "text-destructive"}`}
                >
                  {t.pendencias}
                </span>
              </div>
            ))}
          </div>
        </div>
        <div className="rounded-xl border border-border bg-card p-4">
          <p className="mb-2 text-xs font-medium text-muted-foreground">
            Por moinho
          </p>
          <div className="space-y-1.5">
            {porMoinho.map((m) => (
              <div key={m.moinho} className="flex items-center justify-between text-sm">
                <span>{m.moinho}</span>
                <span
                  className={`font-semibold ${m.pendencias === 0 ? "text-success" : "text-destructive"}`}
                >
                  {m.pendencias}
                </span>
              </div>
            ))}
          </div>
        </div>
      </div>

      {porDia.length > 0 && (
        <div className="rounded-xl border border-border bg-card p-4">
          <p className="mb-2 text-xs font-medium text-muted-foreground">
            Por dia
          </p>
          <div className="flex flex-wrap gap-2">
            {porDia.map((d) => (
              <span
                key={d.date}
                className="rounded-full border border-destructive/30 bg-destructive/10 px-2.5 py-1 text-xs font-medium text-destructive"
              >
                {d.date}: {d.count}
              </span>
            ))}
          </div>
        </div>
      )}

      <div className="overflow-x-auto rounded-xl border border-border bg-card">
        <table className="w-full text-sm">
          <thead className="bg-secondary text-secondary-foreground">
            <tr>
              <th className="px-4 py-2 text-left">Data</th>
              <th className="px-4 py-2 text-left">Turno</th>
              <th className="px-4 py-2 text-left">Moinho</th>
              <th className="px-4 py-2 text-left">Área</th>
            </tr>
          </thead>
          <tbody>
            {pendencias.length === 0 ? (
              <tr>
                <td colSpan={4} className="px-4 py-10 text-center text-muted-foreground">
                  Nenhuma pendência no período selecionado.
                </td>
              </tr>
            ) : (
              pendencias.map((p) => (
                <tr
                  key={`${p.date}|${p.shift}|${p.mill.id}`}
                  className="border-t border-border"
                >
                  <td className="px-4 py-2">{p.date}</td>
                  <td className="px-4 py-2">Turno {p.shift}</td>
                  <td className="px-4 py-2 font-medium">{p.mill.name}</td>
                  <td className="px-4 py-2 text-muted-foreground">{p.mill.area}</td>
                </tr>
              ))
            )}
          </tbody>
        </table>
      </div>
    </section>
  );
}

function MillsManager({
  mills,
  onAdd,
  onDelete,
}: {
  mills: Mill[];
  onAdd: (m: { name: string; area: string; mangas: number }) => Promise<void>;
  onDelete: (id: string) => Promise<void>;
}) {
  const [name, setName] = useState("");
  const [area, setArea] = useState("");
  const [mangas, setMangas] = useState("4");
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [deletingId, setDeletingId] = useState<string | null>(null);
  const [deleteError, setDeleteError] = useState<{ id: string; message: string } | null>(
    null,
  );

  const areas = Array.from(new Set(mills.map((m) => m.area))).sort();

  const submit = async (e: React.FormEvent) => {
    e.preventDefault();
    const qty = Number(mangas);
    if (!name.trim() || !area.trim() || !Number.isInteger(qty) || qty < 1) return;
    setSaving(true);
    setError(null);
    try {
      await onAdd({ name: name.trim(), area: area.trim(), mangas: qty });
      setName("");
      setArea("");
      setMangas("4");
    } catch (err) {
      setError(err instanceof Error ? err.message : "Não foi possível cadastrar o moinho.");
    } finally {
      setSaving(false);
    }
  };

  const remove = async (id: string) => {
    setDeletingId(id);
    setDeleteError(null);
    try {
      await onDelete(id);
    } catch (err) {
      setDeleteError({
        id,
        message:
          err instanceof Error ? err.message : "Não foi possível excluir o moinho.",
      });
    } finally {
      setDeletingId(null);
    }
  };

  return (
    <div className="mx-auto max-w-3xl space-y-6">
      <form
        onSubmit={submit}
        className="space-y-4 rounded-xl border border-border bg-card p-6 shadow-sm"
      >
        <div>
          <h2 className="text-lg font-semibold">Cadastrar moinho</h2>
          <p className="text-sm text-muted-foreground">
            Defina o nome, a área e a quantidade de mangás do moinho.
          </p>
        </div>

        <div className="grid gap-4 md:grid-cols-3">
          <Field label="Nome">
            <input
              value={name}
              onChange={(e) => setName(e.target.value)}
              required
              placeholder="Ex.: M9 680 A"
              className="w-full rounded-md border border-input bg-background px-3 py-2 text-sm"
            />
          </Field>
          <Field label="Área">
            <input
              value={area}
              onChange={(e) => setArea(e.target.value)}
              required
              placeholder="Ex.: Moagem"
              list="areas-moinho"
              className="w-full rounded-md border border-input bg-background px-3 py-2 text-sm"
            />
            <datalist id="areas-moinho">
              {areas.map((a) => (
                <option key={a} value={a} />
              ))}
            </datalist>
          </Field>
          <Field label="Quantidade de mangás">
            <input
              type="number"
              min={1}
              max={50}
              value={mangas}
              onChange={(e) => setMangas(e.target.value)}
              required
              className="w-full rounded-md border border-input bg-background px-3 py-2 text-sm"
            />
          </Field>
        </div>

        {error && <p className="text-sm text-destructive">{error}</p>}

        <button
          type="submit"
          disabled={saving}
          className="rounded-md bg-primary px-5 py-2.5 text-sm font-semibold text-primary-foreground shadow hover:opacity-90 disabled:opacity-60"
        >
          {saving ? "Salvando…" : "Cadastrar moinho"}
        </button>
      </form>

      <div className="overflow-hidden rounded-xl border border-border bg-card">
        <div className="border-b border-border px-4 py-3">
          <h3 className="text-sm font-semibold">
            Moinhos cadastrados ({mills.length})
          </h3>
        </div>
        {mills.length === 0 ? (
          <div className="p-10 text-center text-sm text-muted-foreground">
            Nenhum moinho cadastrado.
          </div>
        ) : (
          <ul className="divide-y divide-border">
            {mills.map((m) => (
              <li key={m.id} className="px-4 py-3">
                <div className="flex items-center justify-between gap-3">
                  <div>
                    <p className="text-sm font-medium">{m.name}</p>
                    <p className="text-xs text-muted-foreground">
                      {m.area} · {m.mangas} mangás
                    </p>
                  </div>
                  <button
                    onClick={() => remove(m.id)}
                    disabled={deletingId === m.id}
                    className="rounded-md border border-input px-3 py-1.5 text-xs font-medium text-muted-foreground transition hover:border-destructive hover:text-destructive disabled:opacity-60"
                  >
                    {deletingId === m.id ? "Excluindo…" : "Excluir"}
                  </button>
                </div>
                {deleteError?.id === m.id && (
                  <p className="mt-2 text-xs text-destructive">{deleteError.message}</p>
                )}
              </li>
            ))}
          </ul>
        )}
      </div>
    </div>
  );
}

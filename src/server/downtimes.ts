import { createServerFn } from "@tanstack/react-start";
import { neon } from "@neondatabase/serverless";
import { z } from "zod";

function sql() {
  const url = process.env.DATABASE_URL;
  if (!url) throw new Error("DATABASE_URL is not set");
  return neon(url);
}

// Dias/turnos em que o moinho ficou parado: não geram pendência de limpeza
// nem entram no esperado da aderência.
async function ensureTable() {
  await sql()`
    create table if not exists downtimes (
      id text primary key,
      mill_id text not null default '',
      date_from text not null,
      date_to text not null,
      shifts text not null,
      motivo text not null default '',
      created_at text not null
    )
  `;
}

const DowntimeInput = z.object({
  id: z.string(),
  // "" = todos os moinhos
  millId: z.string(),
  dateFrom: z.string().regex(/^\d{4}-\d{2}-\d{2}$/),
  dateTo: z.string().regex(/^\d{4}-\d{2}-\d{2}$/),
  shifts: z.array(z.enum(["A", "B", "C"])).min(1),
  motivo: z.string(),
  createdAt: z.string(),
});

export type DowntimeRow = z.infer<typeof DowntimeInput>;

export const getDowntimes = createServerFn({ method: "GET" }).handler(async () => {
  await ensureTable();
  const rows = await sql()`
    select id, mill_id as "millId", date_from as "dateFrom", date_to as "dateTo",
      shifts, motivo, created_at as "createdAt"
    from downtimes
    order by date_from desc
  `;
  return rows.map((r) => ({
    ...r,
    shifts: (r.shifts as string).split("").filter(Boolean),
  })) as DowntimeRow[];
});

export const addDowntime = createServerFn({ method: "POST" })
  .validator(DowntimeInput)
  .handler(async ({ data }) => {
    await ensureTable();
    await sql()`
      insert into downtimes (id, mill_id, date_from, date_to, shifts, motivo, created_at)
      values (
        ${data.id}, ${data.millId}, ${data.dateFrom}, ${data.dateTo},
        ${data.shifts.join("")}, ${data.motivo}, ${data.createdAt}
      )
    `;
    return data;
  });

export const deleteDowntime = createServerFn({ method: "POST" })
  .validator(z.object({ id: z.string() }))
  .handler(async ({ data }) => {
    await sql()`delete from downtimes where id = ${data.id}`;
    return { id: data.id };
  });

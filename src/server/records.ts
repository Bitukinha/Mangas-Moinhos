import { createServerFn } from "@tanstack/react-start";
import { neon } from "@neondatabase/serverless";
import { z } from "zod";

function sql() {
  const url = process.env.DATABASE_URL;
  if (!url) throw new Error("DATABASE_URL is not set");
  return neon(url);
}

async function ensureTable() {
  await sql()`
    create table if not exists records (
      id text primary key,
      mill_id text not null,
      date text not null,
      shift text not null,
      hour text not null,
      responsavel_limpeza text not null,
      responsavel_monitoramento text not null,
      mangas text not null,
      created_at text not null
    )
  `;
  await sql()`create index if not exists records_date_idx on records (date)`;
}

const RecordInput = z.object({
  id: z.string(),
  millId: z.string(),
  date: z.string(),
  shift: z.enum(["A", "B", "C"]),
  hour: z.string(),
  responsavelLimpeza: z.string(),
  responsavelMonitoramento: z.string(),
  mangas: z.array(z.enum(["C", "NC"])),
  createdAt: z.string(),
});

export type RecordRow = z.infer<typeof RecordInput>;

export const getRecords = createServerFn({ method: "GET" }).handler(async () => {
  await ensureTable();
  const rows = await sql()`
    select id, mill_id as "millId", date, shift, hour,
      responsavel_limpeza as "responsavelLimpeza",
      responsavel_monitoramento as "responsavelMonitoramento",
      mangas, created_at as "createdAt"
    from records
    order by created_at desc
  `;
  return rows.map((r) => ({
    ...r,
    mangas: JSON.parse(r.mangas as string),
  })) as RecordRow[];
});

export const addRecord = createServerFn({ method: "POST" })
  .validator(RecordInput)
  .handler(async ({ data }) => {
    await ensureTable();
    await sql()`
      insert into records (
        id, mill_id, date, shift, hour,
        responsavel_limpeza, responsavel_monitoramento, mangas, created_at
      ) values (
        ${data.id}, ${data.millId}, ${data.date}, ${data.shift}, ${data.hour},
        ${data.responsavelLimpeza}, ${data.responsavelMonitoramento},
        ${JSON.stringify(data.mangas)}, ${data.createdAt}
      )
    `;
    return data;
  });

export const deleteRecord = createServerFn({ method: "POST" })
  .validator(z.object({ id: z.string() }))
  .handler(async ({ data }) => {
    await sql()`delete from records where id = ${data.id}`;
    return { id: data.id };
  });

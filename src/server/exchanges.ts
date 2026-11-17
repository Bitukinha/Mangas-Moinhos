import { createServerFn } from "@tanstack/react-start";
import { neon } from "@neondatabase/serverless";
import { z } from "zod";

function sql() {
  const url = process.env.DATABASE_URL;
  if (!url) throw new Error("DATABASE_URL is not set");
  return neon(url);
}

// Trocas de mangas (pode ser lançada com data retroativa).
async function ensureTable() {
  await sql()`
    create table if not exists exchanges (
      id text primary key,
      mill_id text not null,
      date text not null,
      hour text not null default '',
      mangas text not null,
      responsavel text not null,
      motivo text not null default '',
      observacao text not null default '',
      created_at text not null
    )
  `;
  await sql()`create index if not exists exchanges_mill_date_idx on exchanges (mill_id, date)`;
}

const ExchangeInput = z.object({
  id: z.string(),
  millId: z.string(),
  date: z.string().regex(/^\d{4}-\d{2}-\d{2}$/),
  hour: z.string(),
  // Posições das mangas trocadas (1, 2, 3...)
  mangas: z.array(z.number().int().min(1)).min(1),
  responsavel: z.string().trim().min(1),
  motivo: z.string(),
  observacao: z.string(),
  createdAt: z.string(),
});

export type ExchangeRow = z.infer<typeof ExchangeInput>;

export const getExchanges = createServerFn({ method: "GET" }).handler(async () => {
  await ensureTable();
  const rows = await sql()`
    select id, mill_id as "millId", date, hour, mangas, responsavel, motivo, observacao,
      created_at as "createdAt"
    from exchanges
    order by date desc, hour desc
  `;
  return rows.map((r) => ({
    ...r,
    mangas: JSON.parse(r.mangas as string),
  })) as ExchangeRow[];
});

export const addExchange = createServerFn({ method: "POST" })
  .validator(ExchangeInput)
  .handler(async ({ data }) => {
    await ensureTable();
    await sql()`
      insert into exchanges (
        id, mill_id, date, hour, mangas, responsavel, motivo, observacao, created_at
      ) values (
        ${data.id}, ${data.millId}, ${data.date}, ${data.hour},
        ${JSON.stringify(data.mangas)}, ${data.responsavel}, ${data.motivo},
        ${data.observacao}, ${data.createdAt}
      )
    `;
    return data;
  });

export const deleteExchange = createServerFn({ method: "POST" })
  .validator(z.object({ id: z.string() }))
  .handler(async ({ data }) => {
    await sql()`delete from exchanges where id = ${data.id}`;
    return { id: data.id };
  });

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
    create table if not exists washes (
      id text primary key,
      mill_id text not null,
      date text not null,
      hour text not null,
      responsavel text not null,
      observacao text not null default '',
      created_at text not null
    )
  `;
  await sql()`create index if not exists washes_mill_date_idx on washes (mill_id, date)`;
}

const WashInput = z.object({
  id: z.string(),
  millId: z.string(),
  date: z.string(),
  hour: z.string(),
  responsavel: z.string().trim().min(1),
  observacao: z.string(),
  createdAt: z.string(),
});

export type WashRow = z.infer<typeof WashInput>;

export const getWashes = createServerFn({ method: "GET" }).handler(async () => {
  await ensureTable();
  const rows = await sql()`
    select id, mill_id as "millId", date, hour, responsavel, observacao,
      created_at as "createdAt"
    from washes
    order by date desc, hour desc
  `;
  return rows as WashRow[];
});

export const addWash = createServerFn({ method: "POST" })
  .validator(WashInput)
  .handler(async ({ data }) => {
    await ensureTable();
    await sql()`
      insert into washes (id, mill_id, date, hour, responsavel, observacao, created_at)
      values (
        ${data.id}, ${data.millId}, ${data.date}, ${data.hour},
        ${data.responsavel}, ${data.observacao}, ${data.createdAt}
      )
    `;
    return data;
  });

export const deleteWash = createServerFn({ method: "POST" })
  .validator(z.object({ id: z.string() }))
  .handler(async ({ data }) => {
    await sql()`delete from washes where id = ${data.id}`;
    return { id: data.id };
  });

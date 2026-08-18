import { createServerFn } from "@tanstack/react-start";
import { neon } from "@neondatabase/serverless";
import { z } from "zod";

function sql() {
  const url = process.env.DATABASE_URL;
  if (!url) throw new Error("DATABASE_URL is not set");
  return neon(url);
}

const DEFAULT_MILLS = [
  { id: "M-M2", name: "M2 680 A", area: "Moagem", mangas: 4 },
  { id: "M-M6", name: "M6 680 A", area: "Moagem", mangas: 4 },
  { id: "M-M8", name: "M8 680 A", area: "Moagem", mangas: 4 },
  { id: "E-M1", name: "M1 950", area: "Extrusora", mangas: 4 },
  { id: "E-M3", name: "M3 950", area: "Extrusora", mangas: 4 },
  { id: "E-M7", name: "M7 680 A", area: "Extrusora", mangas: 4 },
];

async function ensureTable() {
  await sql()`
    create table if not exists mills (
      id text primary key,
      name text not null,
      area text not null,
      mangas integer not null,
      created_at text not null
    )
  `;
  const [{ count }] = await sql()`select count(*)::int as count from mills`;
  if (count === 0) {
    for (const m of DEFAULT_MILLS) {
      await sql()`
        insert into mills (id, name, area, mangas, created_at)
        values (${m.id}, ${m.name}, ${m.area}, ${m.mangas}, ${new Date().toISOString()})
        on conflict (id) do nothing
      `;
    }
  }
}

const MillInput = z.object({
  name: z.string().trim().min(1),
  area: z.string().trim().min(1),
  mangas: z.number().int().min(1).max(50),
});

export type MillRow = { id: string; name: string; area: string; mangas: number };

export const getMills = createServerFn({ method: "GET" }).handler(async () => {
  await ensureTable();
  const rows = await sql()`select id, name, area, mangas from mills order by area, name`;
  return rows as MillRow[];
});

export const addMill = createServerFn({ method: "POST" })
  .validator(MillInput)
  .handler(async ({ data }) => {
    await ensureTable();
    const id = crypto.randomUUID();
    await sql()`
      insert into mills (id, name, area, mangas, created_at)
      values (${id}, ${data.name}, ${data.area}, ${data.mangas}, ${new Date().toISOString()})
    `;
    return { id, name: data.name, area: data.area, mangas: data.mangas } as MillRow;
  });

export const deleteMill = createServerFn({ method: "POST" })
  .validator(z.object({ id: z.string() }))
  .handler(async ({ data }) => {
    const [{ count }] = await sql()`
      select count(*)::int as count from records where mill_id = ${data.id}
    `;
    if (count > 0) {
      throw new Error(
        "Este moinho possui registros de limpeza e não pode ser excluído.",
      );
    }
    await sql()`delete from mills where id = ${data.id}`;
    return { id: data.id };
  });

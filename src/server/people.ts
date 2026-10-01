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
    create table if not exists people (
      id text primary key,
      name text not null,
      limpeza boolean not null,
      monitoramento boolean not null,
      created_at text not null
    )
  `;
}

const PersonInput = z
  .object({
    name: z.string().trim().min(1),
    limpeza: z.boolean(),
    monitoramento: z.boolean(),
  })
  .refine((p) => p.limpeza || p.monitoramento, {
    message: "Selecione ao menos uma função.",
  });

export type PersonRow = {
  id: string;
  name: string;
  limpeza: boolean;
  monitoramento: boolean;
};

export const getPeople = createServerFn({ method: "GET" }).handler(async () => {
  await ensureTable();
  const rows = await sql()`
    select id, name, limpeza, monitoramento from people order by name
  `;
  return rows as PersonRow[];
});

export const addPerson = createServerFn({ method: "POST" })
  .validator(PersonInput)
  .handler(async ({ data }) => {
    await ensureTable();
    const [{ count }] = await sql()`
      select count(*)::int as count from people where lower(name) = lower(${data.name})
    `;
    if (count > 0) throw new Error("Já existe um responsável com esse nome.");
    const id = crypto.randomUUID();
    await sql()`
      insert into people (id, name, limpeza, monitoramento, created_at)
      values (${id}, ${data.name}, ${data.limpeza}, ${data.monitoramento}, ${new Date().toISOString()})
    `;
    return { id, ...data } as PersonRow;
  });

export const deletePerson = createServerFn({ method: "POST" })
  .validator(z.object({ id: z.string() }))
  .handler(async ({ data }) => {
    await sql()`delete from people where id = ${data.id}`;
    return { id: data.id };
  });

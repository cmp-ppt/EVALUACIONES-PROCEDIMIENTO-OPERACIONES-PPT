----
-- Evaluaciones con nota 100:
--   Ricardo Kunz   -> los 16 de Filtrado 3.8.3.1 .. 3.8.3.16   fecha 24/09/2026
--   Matias Fuentes -> Filtrado 3.8.3.9 .. 3.8.3.16 (8) y
--                     Descarga 3.8.1.1 .. 3.8.1.7  (7)         fecha 30/09/2026
--
-- 3.8.1 tiene SIETE procedimientos (3.8.1.1 a 3.8.1.7), no ocho. Son los 7
-- que Fuentes tenia pendientes en ese proceso.
--
-- Los ids no se hardcodean: se resuelven por apellido completo + nombre.
-- "FUENTES %" no alcanza a "DE LA FUENTE CISTERNA", que es otra persona.
-- El '_' en MAT_AS es comodin de un caracter: calza con o sin tilde.
--
-- Guarda: ninguna evaluacion puede quedar fechada antes de la aprobacion de
-- su documento. Si eso ocurriera el bloque aborta sin escribir nada.
--
-- Solo se tocan ev, evNota y evDate. La difusion se conserva.

do $$
declare
  aprobacion constant jsonb := '{
    "3.8.1.1":"2026-04-20","3.8.1.2":"2026-04-20","3.8.1.3":"2026-04-20",
    "3.8.1.4":"2026-04-20","3.8.1.5":"2026-04-20","3.8.1.6":"2026-04-20",
    "3.8.1.7":"2026-04-20",
    "3.8.3.1":"2026-07-09","3.8.3.2":"2026-07-09","3.8.3.3":"2026-07-09",
    "3.8.3.4":"2026-07-09","3.8.3.5":"2026-07-09","3.8.3.6":"2026-07-09",
    "3.8.3.7":"2026-07-09","3.8.3.8":"2026-07-09",
    "3.8.3.9":"2026-07-14","3.8.3.10":"2026-07-14","3.8.3.11":"2026-07-14",
    "3.8.3.12":"2025-02-25","3.8.3.13":"2026-07-14","3.8.3.14":"2026-07-14",
    "3.8.3.15":"2026-07-14","3.8.3.16":"2026-08-05"
  }'::jsonb;
  d          jsonb;
  w          record;
  pid        text;
  n_match    int;
  code       text;
  eval_patch jsonb;
  n_cells    int := 0;
begin
  select data into d from dashboard_state where id = 'ops_ppt';
  if d is null then
    raise exception 'No existe la fila dashboard_state con id = ops_ppt';
  end if;

  for w in
    select * from (values
      ('KUNZ %',    'RICARDO%', '2026-09-24',
       array['3.8.3.1','3.8.3.2','3.8.3.3','3.8.3.4','3.8.3.5','3.8.3.6',
             '3.8.3.7','3.8.3.8','3.8.3.9','3.8.3.10','3.8.3.11','3.8.3.12',
             '3.8.3.13','3.8.3.14','3.8.3.15','3.8.3.16']),
      ('FUENTES %', 'MAT_AS%',  '2026-09-30',
       array['3.8.3.9','3.8.3.10','3.8.3.11','3.8.3.12','3.8.3.13','3.8.3.14',
             '3.8.3.15','3.8.3.16',
             '3.8.1.1','3.8.1.2','3.8.1.3','3.8.1.4','3.8.1.5','3.8.1.6','3.8.1.7'])
    ) as t(ap_like, nombre_like, eval_date, proc_codes)
  loop
    select count(*), min(p->>'id')
      into n_match, pid
    from jsonb_array_elements(d->'people') p
    where upper(regexp_replace(trim(coalesce(p->>'apPat','') || ' ' || coalesce(p->>'apMat','')),
                               '\s+', ' ', 'g')) like w.ap_like
      and upper(coalesce(p->>'nombre','')) like w.nombre_like;

    if n_match = 0 then
      raise exception 'No se encontro a "%" con nombre "%" en people', w.ap_like, w.nombre_like;
    elsif n_match > 1 then
      raise exception 'Ambiguo: % personas coinciden con "%" / "%"', n_match, w.ap_like, w.nombre_like;
    end if;

    eval_patch := jsonb_build_object('ev', 'ok', 'evNota', 100, 'evDate', w.eval_date);

    foreach code in array w.proc_codes loop
      if w.eval_date::date < (aprobacion ->> code)::date then
        raise exception 'La fecha % es anterior a la aprobacion de % (%). No se escribio nada.',
                        w.eval_date, code, aprobacion ->> code;
      end if;
      -- '||' fusiona: conserva dif y difDate, sobrescribe ev/evNota/evDate
      d := jsonb_set(d, array['cells', pid, code],
                     coalesce(d #> array['cells', pid, code], '{}'::jsonb) || eval_patch,
                     true);
      n_cells := n_cells + 1;
    end loop;

    raise notice 'OK  % (%) -> nota 100 en % evaluaciones con fecha %',
                 rpad(w.ap_like, 11), pid, array_length(w.proc_codes, 1), w.eval_date;
  end loop;

  update dashboard_state
     set data = d, updated_at = now()
   where id = 'ops_ppt';

  raise notice 'Listo: % evaluaciones actualizadas', n_cells;
end $$;

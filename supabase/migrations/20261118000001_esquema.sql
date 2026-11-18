-- Esquema inicial (plan paso 1). Varias escuelas en la misma base de datos: casi toda tabla lleva
-- escuela_id, y las tablas hijas lo amarran a su padre con llaves compuestas (id, escuela_id), así
-- un registro no puede quedar en una escuela distinta a la de su clase o grupo.
-- Los permisos (Row Level Security) están en la migración siguiente.

create type public.rol as enum ('superadmin', 'director', 'maestro');
create type public.paquete as enum ('individual', 'empresarial');
create type public.plantilla as enum ('asist', 'sino', 'trab', 'disc', 'num');
create type public.tipo_cierre as enum ('parcial', 'semestre');

-- ---------- Escuelas y cuentas ----------
create table public.escuelas (
  id uuid primary key default gen_random_uuid(),
  nombre text not null check (length(trim(nombre)) > 0),
  clave text not null unique check (clave ~ '^[A-Z0-9]{2,10}$'),     -- ESC01: va en la clave de cada alumno
  paquete public.paquete not null default 'empresarial',
  precio_mensual numeric(10, 2) check (precio_mensual >= 0),          -- por maestro; null = precio de lista
  cuentas_contratadas int not null default 1 check (cuentas_contratadas >= 1),
  vence_el date,
  activa boolean not null default true,
  creado_en timestamptz not null default now()
);

create table public.usuarios (
  id uuid primary key references auth.users (id) on delete cascade,
  escuela_id uuid references public.escuelas (id) on delete restrict,  -- null solo para superadmin
  rol public.rol not null,
  nombre text not null default '',
  correo text not null,
  activo boolean not null default true,
  creado_en timestamptz not null default now(),
  check (rol = 'superadmin' or escuela_id is not null)
);
create index on public.usuarios (escuela_id);

create table public.invitaciones (
  id uuid primary key default gen_random_uuid(),
  escuela_id uuid not null references public.escuelas (id) on delete cascade,
  correo text not null check (correo = lower(trim(correo)) and correo like '%@%'),
  rol public.rol not null default 'maestro' check (rol in ('director', 'maestro')),
  token uuid not null unique default gen_random_uuid(),
  creada_por uuid references public.usuarios (id) on delete set null,
  creada_en timestamptz not null default now(),
  vence_en timestamptz not null default now() + interval '7 days',
  usada_en timestamptz,
  usada_por uuid references public.usuarios (id) on delete set null
);
create index on public.invitaciones (escuela_id);

create table public.pagos (
  id uuid primary key default gen_random_uuid(),
  escuela_id uuid not null references public.escuelas (id) on delete cascade,
  periodo_inicio date not null,
  periodo_fin date not null check (periodo_fin >= periodo_inicio),
  monto numeric(10, 2) not null check (monto >= 0),
  metodo text not null default 'transferencia',
  factura text,
  registrado_por uuid references public.usuarios (id) on delete set null,
  creado_en timestamptz not null default now()
);
create index on public.pagos (escuela_id);

-- ---------- Periodos: semestre con 3 parciales ----------
create table public.semestres (
  id uuid primary key default gen_random_uuid(),
  escuela_id uuid not null references public.escuelas (id) on delete cascade,
  nombre text not null,                                               -- "Ago–Dic 2026"
  inicio date not null,
  fin date not null check (fin > inicio),
  unique (escuela_id, nombre),
  unique (id, escuela_id)
);

create table public.parciales (
  id uuid primary key default gen_random_uuid(),
  escuela_id uuid not null,
  semestre_id uuid not null,
  numero int not null check (numero between 1 and 6),
  inicio date not null,
  fin date not null check (fin >= inicio),
  unique (semestre_id, numero),
  foreign key (semestre_id, escuela_id) references public.semestres (id, escuela_id) on delete cascade
);

-- ---------- Alumnos y grupos ----------
-- El alumno existe aparte del grupo: cada semestre un grupo lo inscribe, así su historial lo sigue.
create table public.alumnos (
  id uuid primary key default gen_random_uuid(),
  escuela_id uuid not null references public.escuelas (id) on delete cascade,
  clave text not null,                                                -- ESC01-V205-001 (la asigna la app)
  matricula text,                                                     -- oficial, cuando la escuela la dé
  nombre text not null check (length(trim(nombre)) > 0),
  creado_en timestamptz not null default now(),
  unique (escuela_id, clave),
  unique (escuela_id, matricula),
  unique (id, escuela_id)
);

create table public.grupos (
  id uuid primary key default gen_random_uuid(),
  escuela_id uuid not null references public.escuelas (id) on delete cascade,
  semestre_id uuid,
  nombre text not null check (length(trim(nombre)) > 0),              -- V205
  unique nulls not distinct (escuela_id, semestre_id, nombre),
  unique (id, escuela_id),
  foreign key (semestre_id, escuela_id) references public.semestres (id, escuela_id) on delete set null (semestre_id)
);

create table public.inscripciones (
  id uuid primary key default gen_random_uuid(),
  escuela_id uuid not null,
  grupo_id uuid not null,
  alumno_id uuid not null,
  numero_lista text not null default '',
  alta date not null default current_date,
  baja date,
  unique (grupo_id, alumno_id),
  foreign key (grupo_id, escuela_id) references public.grupos (id, escuela_id) on delete cascade,
  foreign key (alumno_id, escuela_id) references public.alumnos (id, escuela_id) on delete cascade
);
create index on public.inscripciones (alumno_id);

-- ---------- Clase = grupo + materia + maestro ----------
create table public.clases (
  id uuid primary key default gen_random_uuid(),
  escuela_id uuid not null,
  grupo_id uuid not null,
  maestro_id uuid not null references public.usuarios (id) on delete restrict,
  materia text not null default '',
  retardo int not null default 100 check (retardo in (0, 50, 100)),   -- cuánto vale un retardo
  escala int not null default 10 check (escala in (10, 100)),
  config_semestral jsonb not null default '{}'::jsonb,                -- promedio de parciales o con final, exentos…
  creado_en timestamptz not null default now(),
  unique (id, escuela_id),
  foreign key (grupo_id, escuela_id) references public.grupos (id, escuela_id) on delete cascade
);
create index on public.clases (maestro_id);
create index on public.clases (grupo_id);

create table public.secciones (
  id uuid primary key default gen_random_uuid(),
  escuela_id uuid not null,
  clase_id uuid not null,
  nombre text not null check (length(trim(nombre)) > 0 and nombre !~ '[:\n]'),
  plantilla public.plantilla not null,
  porcentaje numeric(5, 2) not null default 0 check (porcentaje between 0 and 100),
  minimo int check (minimo >= 1),                                     -- solo Sí/No: "Sí" para el 100 %
  orden int not null default 0,
  unique (clase_id, nombre),
  unique (id, escuela_id),
  unique (id, clase_id),
  foreign key (clase_id, escuela_id) references public.clases (id, escuela_id) on delete cascade
);

-- Trabajos de un día (plantilla 0/50/100): varios por fecha, con título opcional.
create table public.trabajos (
  id uuid primary key default gen_random_uuid(),
  escuela_id uuid not null,
  seccion_id uuid not null,
  fecha date not null,
  titulo text not null default '',
  orden int not null default 0,
  unique (id, seccion_id),
  foreign key (seccion_id, escuela_id) references public.secciones (id, escuela_id) on delete cascade
);
create index on public.trabajos (seccion_id, fecha);

-- Cada dato capturado: un alumno, en una sección de una clase, en una fecha (y un trabajo, si aplica).
-- valor: A/F/R/J, S/N, 0/50/100, 0 (disciplina) o 0–100.
create table public.capturas (
  id uuid primary key default gen_random_uuid(),
  escuela_id uuid not null,
  clase_id uuid not null,
  seccion_id uuid not null,
  alumno_id uuid not null,
  fecha date not null,
  trabajo_id uuid,
  valor text not null check (valor ~ '^([AFRJSN]|[0-9]{1,3}(\.[0-9]+)?)$'),
  actualizado_en timestamptz not null default now(),
  actualizado_por uuid references public.usuarios (id) on delete set null,
  unique nulls not distinct (seccion_id, alumno_id, fecha, trabajo_id),
  foreign key (clase_id, escuela_id) references public.clases (id, escuela_id) on delete cascade,
  foreign key (seccion_id, clase_id) references public.secciones (id, clase_id) on delete cascade,
  foreign key (alumno_id, escuela_id) references public.alumnos (id, escuela_id) on delete cascade,
  foreign key (trabajo_id, seccion_id) references public.trabajos (id, seccion_id) on delete cascade
);
create index on public.capturas (clase_id, fecha);
create index on public.capturas (alumno_id);

-- Calificaciones definitivas al cerrar un parcial o el semestre. Para corregir se reabre (queda registro).
create table public.cierres (
  id uuid primary key default gen_random_uuid(),
  escuela_id uuid not null,
  clase_id uuid not null,
  tipo public.tipo_cierre not null,
  parcial_id uuid references public.parciales (id) on delete restrict,
  calificaciones jsonb not null,                                      -- { alumno_id: { final, partes… } }
  cerrado_por uuid references public.usuarios (id) on delete set null,
  cerrado_en timestamptz not null default now(),
  reabierto_por uuid references public.usuarios (id) on delete set null,
  reabierto_en timestamptz,
  check ((tipo = 'parcial') = (parcial_id is not null)),
  foreign key (clase_id, escuela_id) references public.clases (id, escuela_id) on delete cascade
);
create index on public.cierres (clase_id);

-- Cuadernillo PDF de una clase (el archivo vive en Supabase Storage).
create table public.cuadernillos (
  id uuid primary key default gen_random_uuid(),
  escuela_id uuid not null,
  clase_id uuid not null unique,
  nombre text not null,
  ruta text not null,
  tamano int not null check (tamano > 0),
  pagina int not null default 1,
  actualizado_en timestamptz not null default now(),
  foreign key (clase_id, escuela_id) references public.clases (id, escuela_id) on delete cascade
);

-- Historial de cambios: quién cambió qué y cuándo. Solo lo escriben los triggers; nadie lo edita.
create table public.historial (
  id bigint generated always as identity primary key,
  escuela_id uuid not null,
  tabla text not null,
  registro_id uuid not null,
  accion text not null check (accion in ('insert', 'update', 'delete')),
  antes jsonb,
  despues jsonb,
  usuario_id uuid,
  en timestamptz not null default now()
);
create index on public.historial (escuela_id, en desc);
create index on public.historial (registro_id);

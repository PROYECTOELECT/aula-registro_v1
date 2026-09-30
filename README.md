# Registro de Apertura de Aulas v3

Aplicación web responsive (celular y escritorio) — FastAPI + SQLAlchemy.

## ¿Listo para GitHub / Vercel?

**Sí.** El código ya está preparado:

| Ítem | Estado |
|------|--------|
| `vercel.json` | Configurado |
| `requirements.txt` | Incluye `psycopg2-binary` (PostgreSQL) |
| `DATABASE_URL` | Si existe → Postgres; si no → SQLite local |
| Tablas | Se crean **solas** al arrancar (`init_db`) |
| Celular | Bootstrap + CSS móvil (viewport, columnas, tabla scroll) |

**Importante:** GitHub **no** es la base de datos. Solo guarda el código.  
La base de datos es **PostgreSQL** (Neon, Supabase o Vercel Postgres).

---

## Tablas (se crean automáticamente)

No tienes que crearlas a mano en GitHub ni en SQL. Al primer arranque con `DATABASE_URL`, la app crea:

| Tabla | Contenido |
|-------|-----------|
| `usuarios` | Admin, master, general, cupos, foto, invitación |
| `edificios` | Edificios del Excel o agregados manualmente |
| `archivo_horario` | Hasta 5 Excel de horario |
| `aulas` | Aulas parseadas del Excel |
| `registros_apertura` | Aperturas (no se eliminan) |

Usuario inicial: **`admin` / `admin123`**

---

## Paso a paso: GitHub + Vercel

### Paso 1 — Subir código a GitHub

```bash
cd aula_pkg7
git init
git add .
git commit -m "Registro de aulas v3 - listo para Vercel"
git branch -M main
git remote add origin https://github.com/TU_USUARIO/aula-registro.git
git push -u origin main
```

(Crea el repo vacío antes en github.com → New repository)

### Paso 2 — Crear PostgreSQL gratis (Neon)

1. Entra a https://neon.tech y regístrate  
2. **Create project**  
3. Copia la **Connection string** (parece así):

```
postgresql://usuario:clave@ep-xxxx.us-east-2.aws.neon.tech/neondb?sslmode=require
```

### Paso 3 — Desplegar en Vercel

1. Entra a https://vercel.com → **Add New Project**  
2. **Import** el repositorio de GitHub  
3. Framework Preset: **Other**  
4. **Environment Variables** → agrega:

| Name | Value |
|------|--------|
| `DATABASE_URL` | (pega la connection string de Neon) |

5. **Deploy**

### Paso 4 — Verificar

1. Abre la URL que te da Vercel (ej. `https://aula-registro.vercel.app`)  
2. Login: `admin` / `admin123`  
3. Cambia la contraseña (botón llave 🔑)  
4. Sube un Excel de horario (como master/admin)  
5. Prueba en el celular abriendo la misma URL  

Las tablas se crean solas en Neon la primera vez que entra alguien.

### Paso 5 — (Opcional) Dominio propio

En Vercel → Project → Settings → Domains.

---

## Local (prueba antes de subir)

```bash
cd aula_pkg7
rm -f aula_registro.db   # solo si actualizas versión
pip install -r requirements.txt
python -m uvicorn main:app --host 0.0.0.0 --port 8000 --reload
```

Abre http://localhost:8000 (también desde el celular en la misma red WiFi: `http://IP-DE-TU-PC:8000`).

---

## Limitaciones en Vercel (serverless)

- Archivos Excel/fotos en disco pueden perderse entre despliegues (el **contenido de aulas** sí queda en Postgres).  
- Sesiones en memoria: si hay muchas instancias, a veces pide login de nuevo.  
- Para producción grande: Redis (sesiones) + Vercel Blob/S3 (archivos).

Para la mayoría de usos institucionales, Neon + Vercel es suficiente.

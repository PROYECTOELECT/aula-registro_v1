# Registro de Apertura de Aulas

Plataforma web (FastAPI) para registro de apertura de aulas físicas, horarios Excel y reservas.

## Funciones
- Roles: admin, master, general
- Carga de mapas horario Excel (hasta 5)
- Carga de reservas de aulas (Excel)
- Registro de apertura en tiempo real
- Reportes PDF / Excel
- PWA instalable en celular
- PostgreSQL (Neon) en la nube / SQLite en local

## Login inicial
- Usuario: `admin`
- Contraseña: `admin123`  
(Cámbiala al entrar)

## Local
```bash
pip install -r requirements.txt
python -m uvicorn main:app --host 0.0.0.0 --port 8000 --reload
```
Abrir: http://localhost:8000

## Deploy Vercel + Neon
1. Sube este repo a GitHub
2. Neon: crea proyecto y copia `DATABASE_URL`
3. Vercel: Import repo → Environment Variable `DATABASE_URL` → Deploy
4. Framework: FastAPI (o Other). Entrypoint: `main.py`

## Archivos principales
| Archivo | Uso |
|---------|-----|
| main.py | App FastAPI |
| database.py | Modelos SQLAlchemy |
| parser_excel.py | Horarios por sala |
| parser_reservas.py | Reservas de aulas |
| templates/ | HTML |
| static/ | PWA (manifest, icons, sw) |
| vercel.json | Config Vercel |
| requirements.txt | Dependencias |

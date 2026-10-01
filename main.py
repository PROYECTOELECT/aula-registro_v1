from fastapi import FastAPI, Request, Depends, UploadFile, File, Form, HTTPException, Query
from fastapi.responses import HTMLResponse, StreamingResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from sqlalchemy.orm import Session
from sqlalchemy import func
from datetime import datetime, date, timedelta
from typing import Optional
import os
import shutil
import io
import json
import secrets
import hmac
import base64
import hashlib

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch

from database import (
    init_db, get_db, Aula, RegistroApertura, ArchivoHorario, Usuario, Edificio,
    ArchivoReserva, ReservaAula,
    hash_password, verify_password, contar_generales_de_master, MAX_ARCHIVOS_EXCEL,
)
from parser_excel import parsear_aulas_desde_excel
from parser_reservas import parsear_reservas_desde_excel

app = FastAPI(title="Registro de Apertura de Aulas", version="3.0")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
# En Vercel el FS es de solo lectura; usar /tmp para uploads
_IS_VERCEL = bool(os.environ.get("VERCEL") or os.environ.get("VERCEL_ENV"))
if _IS_VERCEL:
    UPLOAD_DIR = "/tmp/uploads"
    FOTOS_DIR = "/tmp/uploads/fotos"
else:
    UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
    FOTOS_DIR = os.path.join(BASE_DIR, "uploads", "fotos")
try:
    os.makedirs(UPLOAD_DIR, exist_ok=True)
    os.makedirs(FOTOS_DIR, exist_ok=True)
except OSError as e:
    print(f"[warn] makedirs uploads: {e}")

static_dir = os.path.join(BASE_DIR, "static")
try:
    os.makedirs(static_dir, exist_ok=True)
except OSError:
    pass
try:
    if os.path.isdir(static_dir):
        app.mount("/static", StaticFiles(directory=static_dir), name="static")
    if os.path.isdir(UPLOAD_DIR):
        app.mount("/uploads", StaticFiles(directory=UPLOAD_DIR), name="uploads")
except Exception as e:
    print(f"[warn] mount static/uploads: {e}")
templates = Jinja2Templates(directory=os.path.join(BASE_DIR, "templates"))

# Sesión firmada (funciona en Vercel serverless; no usa memoria del proceso)
SESSION_SECRET = os.environ.get("SESSION_SECRET") or os.environ.get("DATABASE_URL") or "aula-registro-dev-secret-change-me"
SESSION_HOURS = 12


def _safe_init_db():
    try:
        init_db()
    except Exception as e:
        print(f"[warn] init_db: {e}")

_safe_init_db()


def _b64url_encode(raw: bytes) -> str:
    return base64.urlsafe_b64encode(raw).decode("ascii").rstrip("=")


def _b64url_decode(s: str) -> bytes:
    pad = "=" * (-len(s) % 4)
    return base64.urlsafe_b64decode(s + pad)


def crear_sesion(usuario: Usuario) -> str:
    """Token firmado con HMAC: payload.exp.signature (sin estado en servidor)."""
    payload = {
        "user_id": usuario.id,
        "username": usuario.username,
        "rol": usuario.rol,
        "nombre": usuario.nombre,
        "foto": usuario.foto,
        "creado_por": usuario.creado_por,
        "exp": (datetime.now() + timedelta(hours=SESSION_HOURS)).timestamp(),
    }
    body = _b64url_encode(json.dumps(payload, ensure_ascii=False, separators=(",", ":")).encode("utf-8"))
    sig = hmac.new(SESSION_SECRET.encode("utf-8"), body.encode("ascii"), hashlib.sha256).hexdigest()
    return f"{body}.{sig}"


def obtener_usuario_sesion(request: Request):
    token = request.cookies.get("session_token")
    if not token or "." not in token:
        return None
    try:
        body, sig = token.rsplit(".", 1)
        expect = hmac.new(SESSION_SECRET.encode("utf-8"), body.encode("ascii"), hashlib.sha256).hexdigest()
        if not hmac.compare_digest(sig, expect):
            return None
        payload = json.loads(_b64url_decode(body).decode("utf-8"))
        if float(payload.get("exp", 0)) < datetime.now().timestamp():
            return None
        return {
            "user_id": payload.get("user_id"),
            "username": payload.get("username"),
            "rol": payload.get("rol"),
            "nombre": payload.get("nombre"),
            "foto": payload.get("foto"),
            "creado_por": payload.get("creado_por"),
        }
    except Exception:
        return None


def _cookie_secure(request: Request) -> bool:
    # HTTPS en Vercel
    return bool(os.environ.get("VERCEL") or request.url.scheme == "https")


def set_session_cookie(resp, token: str, request: Request = None):
    secure = True if os.environ.get("VERCEL") else False
    resp.set_cookie(
        "session_token",
        token,
        httponly=True,
        max_age=SESSION_HOURS * 3600,
        samesite="lax",
        secure=secure,
        path="/",
    )


def requiere_login(request: Request):
    user = obtener_usuario_sesion(request)
    if not user:
        raise HTTPException(status_code=401, detail="No autenticado")
    return user


def requiere_master_o_admin(request: Request):
    user = requiere_login(request)
    if user["rol"] not in ("admin", "master"):
        raise HTTPException(status_code=403, detail="Solo administrador o master")
    return user


def requiere_admin(request: Request):
    user = requiere_login(request)
    if user["rol"] != "admin":
        raise HTTPException(status_code=403, detail="Solo administrador")
    return user


def nombre_master(db, master_id):
    if not master_id:
        return None
    m = db.query(Usuario).filter(Usuario.id == master_id).first()
    return m.nombre if m else None


# ==================== AUTH ====================

@app.get("/login", response_class=HTMLResponse)
async def login_page(request: Request):
    if obtener_usuario_sesion(request):
        return RedirectResponse("/", status_code=302)
    return templates.TemplateResponse("login.html", {"request": request, "error": None})


@app.post("/login")
async def login(request: Request, username: str = Form(...), password: str = Form(...), db: Session = Depends(get_db)):
    user = db.query(Usuario).filter(Usuario.username == username.strip()).first()
    if not user or not user.activo or not verify_password(password, user.password_hash, user.salt):
        return templates.TemplateResponse("login.html", {"request": request, "error": "Usuario o contraseña incorrectos"}, status_code=401)
    token = crear_sesion(user)
    resp = RedirectResponse("/", status_code=302)
    set_session_cookie(resp, token, request)
    return resp


@app.get("/logout")
async def logout(request: Request):
    resp = RedirectResponse("/login", status_code=302)
    resp.delete_cookie("session_token", path="/")
    return resp


# ==================== AUTOREGISTRO ====================

@app.get("/registro", response_class=HTMLResponse)
async def registro_page(request: Request, db: Session = Depends(get_db)):
    if obtener_usuario_sesion(request):
        return RedirectResponse("/", status_code=302)
    masters = db.query(Usuario).filter(Usuario.rol == "master", Usuario.activo == True).all()
    disponibles = []
    for m in masters:
        usados = contar_generales_de_master(db, m.id)
        cupo = m.cupo_max or 0
        if cupo > usados:
            disponibles.append({
                "id": m.id, "nombre": m.nombre, "username": m.username,
                "codigo": m.codigo_invitacion, "cupo_restante": cupo - usados, "cupo_max": cupo,
            })
    return templates.TemplateResponse("registro.html", {"request": request, "error": None, "ok": None, "masters": disponibles})


@app.post("/registro")
async def registro_post(
    request: Request,
    username: str = Form(...), password: str = Form(...), nombre: str = Form(...),
    codigo_invitacion: str = Form(...), foto: Optional[UploadFile] = File(None),
    db: Session = Depends(get_db),
):
    username = username.strip().lower()
    codigo = codigo_invitacion.strip().upper()
    nombre = nombre.strip()
    masters = db.query(Usuario).filter(Usuario.rol == "master", Usuario.activo == True).all()
    disponibles = []
    for m in masters:
        usados = contar_generales_de_master(db, m.id)
        cupo = m.cupo_max or 0
        if cupo > usados:
            disponibles.append({
                "id": m.id, "nombre": m.nombre, "username": m.username,
                "codigo": m.codigo_invitacion, "cupo_restante": cupo - usados, "cupo_max": cupo,
            })

    def err(msg):
        return templates.TemplateResponse("registro.html", {"request": request, "error": msg, "ok": None, "masters": disponibles}, status_code=400)

    if len(username) < 3 or len(password) < 4:
        return err("Usuario o contraseña inválidos")
    if db.query(Usuario).filter(Usuario.username == username).first():
        return err("Ese nombre de usuario ya existe")
    master = db.query(Usuario).filter(Usuario.rol == "master", Usuario.activo == True, Usuario.codigo_invitacion == codigo).first()
    if not master:
        return err("Código de invitación inválido")
    if contar_generales_de_master(db, master.id) >= (master.cupo_max or 0):
        return err(f"Cupo del master {master.nombre} agotado")

    foto_path = None
    if foto and foto.filename:
        ext = os.path.splitext(foto.filename)[1].lower()
        if ext in (".jpg", ".jpeg", ".png", ".gif", ".webp"):
            fname = f"{username}_{secrets.token_hex(4)}{ext}"
            with open(os.path.join(FOTOS_DIR, fname), "wb") as f:
                shutil.copyfileobj(foto.file, f)
            foto_path = f"fotos/{fname}"

    h, s = hash_password(password)
    db.add(Usuario(username=username, password_hash=h, salt=s, nombre=nombre, rol="general",
                   activo=True, creado_por=master.id, foto=foto_path))
    db.commit()
    return templates.TemplateResponse("registro.html", {
        "request": request, "error": None,
        "ok": f"Cuenta creada bajo «{master.nombre}». Inicia sesión como {username}.",
        "masters": disponibles,
    })


# ==================== PÁGINAS ====================

@app.get("/", response_class=HTMLResponse)
async def index(request: Request, db: Session = Depends(get_db)):
    user = obtener_usuario_sesion(request)
    if not user:
        return RedirectResponse("/login", status_code=302)

    archivos = db.query(ArchivoHorario).filter(ArchivoHorario.activo == True).order_by(ArchivoHorario.fecha_carga.desc()).all()
    edificios = db.query(Edificio).filter(Edificio.activo == True).order_by(Edificio.nombre).all()
    registros = db.query(RegistroApertura).order_by(RegistroApertura.fecha_hora.desc()).limit(100).all()

    stats_edificio = db.query(RegistroApertura.edificio, func.count(RegistroApertura.id)).group_by(RegistroApertura.edificio).all()
    indicadores = [{"edificio": e or "Sin edificio", "total": c} for e, c in stats_edificio]
    stats_usuario = (
        db.query(RegistroApertura.usuario_nombre, func.count(RegistroApertura.id))
        .group_by(RegistroApertura.usuario_nombre)
        .order_by(func.count(RegistroApertura.id).desc()).all()
    )
    indicadores_usuarios = [{"nombre": n or "Sin nombre", "total": c} for n, c in stats_usuario]
    stats_aula = (
        db.query(
            RegistroApertura.aula_codigo,
            RegistroApertura.aula_nombre,
            RegistroApertura.edificio,
            func.count(RegistroApertura.id).label("total"),
        )
        .group_by(RegistroApertura.aula_codigo, RegistroApertura.aula_nombre, RegistroApertura.edificio)
        .order_by(func.count(RegistroApertura.id).desc())
        .limit(15)
        .all()
    )
    indicadores_aulas = [
        {"codigo": a.aula_codigo, "nombre": a.aula_nombre or "", "edificio": a.edificio or "Sin edificio", "total": a.total}
        for a in stats_aula
    ]

    master_nombre = nombre_master(db, user.get("creado_por")) if user["rol"] == "general" else None
    lista_usuarios = []
    if user["rol"] in ("admin", "master"):
        lista_usuarios = db.query(Usuario).filter(Usuario.activo == True).order_by(Usuario.username).all()

    return templates.TemplateResponse("index.html", {
        "request": request, "user": user, "archivos": archivos,
        "max_archivos": MAX_ARCHIVOS_EXCEL, "total_archivos": len(archivos),
        "edificios": edificios, "registros": registros,
        "total_registros": db.query(RegistroApertura).count(),
        "indicadores": indicadores, "indicadores_usuarios": indicadores_usuarios,
        "indicadores_aulas": indicadores_aulas,
        "es_admin": user["rol"] == "admin",
        "es_master_o_admin": user["rol"] in ("admin", "master"),
        "master_nombre": master_nombre, "lista_usuarios": lista_usuarios,
    })


@app.get("/usuarios", response_class=HTMLResponse)
async def pagina_usuarios(request: Request, db: Session = Depends(get_db)):
    user = obtener_usuario_sesion(request)
    if not user:
        return RedirectResponse("/login", status_code=302)
    if user["rol"] not in ("admin", "master"):
        return RedirectResponse("/", status_code=302)
    if user["rol"] == "admin":
        usuarios = db.query(Usuario).order_by(Usuario.rol, Usuario.username).all()
    else:
        usuarios = db.query(Usuario).filter(
            (Usuario.creado_por == user["user_id"]) | (Usuario.id == user["user_id"])
        ).order_by(Usuario.username).all()
    cupos_info = {}
    for u in usuarios:
        if u.rol == "master":
            cupos_info[u.id] = {"usados": contar_generales_de_master(db, u.id), "max": u.cupo_max or 0}
    mi_cupo = None
    if user["rol"] == "master":
        yo = db.query(Usuario).filter(Usuario.id == user["user_id"]).first()
        mi_cupo = {"usados": contar_generales_de_master(db, user["user_id"]), "max": yo.cupo_max or 0, "codigo": yo.codigo_invitacion}
    return templates.TemplateResponse("usuarios.html", {
        "request": request, "user": user, "usuarios": usuarios,
        "es_admin": user["rol"] == "admin", "cupos_info": cupos_info, "mi_cupo": mi_cupo,
    })


@app.get("/rendimiento", response_class=HTMLResponse)
async def pagina_rendimiento(request: Request, db: Session = Depends(get_db)):
    user = obtener_usuario_sesion(request)
    if not user:
        return RedirectResponse("/login", status_code=302)
    if user["rol"] not in ("admin", "master"):
        return RedirectResponse("/", status_code=302)

    hoy = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    semana = hoy - timedelta(days=7)
    mes = hoy - timedelta(days=30)

    usuarios = db.query(Usuario).filter(Usuario.activo == True, Usuario.rol.in_(["general", "master", "admin"])).all()
    metricas = []
    for u in usuarios:
        total = db.query(RegistroApertura).filter(RegistroApertura.usuario_id == u.id).count()
        hoy_c = db.query(RegistroApertura).filter(RegistroApertura.usuario_id == u.id, RegistroApertura.fecha_hora >= hoy).count()
        sem_c = db.query(RegistroApertura).filter(RegistroApertura.usuario_id == u.id, RegistroApertura.fecha_hora >= semana).count()
        mes_c = db.query(RegistroApertura).filter(RegistroApertura.usuario_id == u.id, RegistroApertura.fecha_hora >= mes).count()
        if total == 0 and u.rol == "admin":
            continue  # opcional: ocultar admin sin actividad
        metricas.append({
            "id": u.id, "nombre": u.nombre, "username": u.username, "rol": u.rol,
            "total": total, "hoy": hoy_c, "semana": sem_c, "mes": mes_c,
        })
    metricas.sort(key=lambda x: x["total"], reverse=True)
    total_global = db.query(RegistroApertura).count()

    return templates.TemplateResponse("rendimiento.html", {
        "request": request, "user": user, "metricas": metricas,
        "total_global": total_global, "es_admin": user["rol"] == "admin",
        "es_master_o_admin": True,
    })


# ==================== API USUARIOS ====================

@app.post("/api/usuarios")
async def crear_usuario(
    request: Request, username: str = Form(...), password: str = Form(...), nombre: str = Form(...),
    rol: str = Form("general"), cupo_max: Optional[int] = Form(None),
    foto: Optional[UploadFile] = File(None), db: Session = Depends(get_db),
):
    actor = requiere_master_o_admin(request)
    username = username.strip().lower()
    if len(username) < 3 or len(password) < 4:
        raise HTTPException(400, "Usuario o contraseña muy cortos")
    if rol not in ("admin", "master", "general"):
        raise HTTPException(400, "Rol inválido")
    if actor["rol"] == "master":
        if rol != "general":
            raise HTTPException(403, "Master solo crea generales")
        yo = db.query(Usuario).filter(Usuario.id == actor["user_id"]).first()
        if contar_generales_de_master(db, actor["user_id"]) >= (yo.cupo_max or 0):
            raise HTTPException(400, "Cupo agotado")
    if db.query(Usuario).filter(Usuario.username == username).first():
        raise HTTPException(400, "Usuario ya existe")

    foto_path = None
    if foto and foto.filename:
        ext = os.path.splitext(foto.filename)[1].lower()
        if ext in (".jpg", ".jpeg", ".png", ".gif", ".webp"):
            fname = f"{username}_{secrets.token_hex(4)}{ext}"
            with open(os.path.join(FOTOS_DIR, fname), "wb") as f:
                shutil.copyfileobj(foto.file, f)
            foto_path = f"fotos/{fname}"

    codigo_inv = secrets.token_hex(4).upper() if rol == "master" else None
    cupo = max(0, int(cupo_max or 0)) if rol == "master" else None
    h, s = hash_password(password)
    db.add(Usuario(username=username, password_hash=h, salt=s, nombre=nombre.strip(), rol=rol,
                   activo=True, cupo_max=cupo, codigo_invitacion=codigo_inv,
                   creado_por=actor["user_id"], foto=foto_path))
    db.commit()
    msg = f"Usuario {username} creado ({rol})"
    if codigo_inv:
        msg += f" · Cupo: {cupo} · Código: {codigo_inv}"
    return {"ok": True, "mensaje": msg, "codigo_invitacion": codigo_inv}


@app.post("/api/usuarios/{user_id}/cupo")
async def actualizar_cupo(user_id: int, request: Request, cupo_max: int = Form(...), db: Session = Depends(get_db)):
    requiere_admin(request)
    u = db.query(Usuario).filter(Usuario.id == user_id, Usuario.rol == "master").first()
    if not u:
        raise HTTPException(404, "Master no encontrado")
    u.cupo_max = max(0, int(cupo_max))
    if not u.codigo_invitacion:
        u.codigo_invitacion = secrets.token_hex(4).upper()
    db.commit()
    return {"ok": True, "cupo_max": u.cupo_max, "codigo_invitacion": u.codigo_invitacion}


@app.post("/api/usuarios/{user_id}/toggle")
async def toggle_usuario(user_id: int, request: Request, db: Session = Depends(get_db)):
    actor = requiere_master_o_admin(request)
    u = db.query(Usuario).filter(Usuario.id == user_id).first()
    if not u or u.username == "admin" or u.id == actor["user_id"]:
        raise HTTPException(400, "Operación no permitida")
    if actor["rol"] == "master" and (u.creado_por != actor["user_id"] or u.rol != "general"):
        raise HTTPException(403, "Solo tus usuarios generales")
    u.activo = not u.activo
    db.commit()
    return {"ok": True, "activo": u.activo}


@app.post("/api/usuarios/{user_id}/eliminar")
async def eliminar_usuario(user_id: int, request: Request, db: Session = Depends(get_db)):
    requiere_admin(request)
    u = db.query(Usuario).filter(Usuario.id == user_id).first()
    if not u or u.username == "admin":
        raise HTTPException(400, "No se puede eliminar")
    db.delete(u)
    db.commit()
    return {"ok": True}



@app.post("/api/cambiar-password")
async def cambiar_password(
    request: Request,
    password_actual: str = Form(...),
    password_nueva: str = Form(...),
    password_confirmar: str = Form(...),
    db: Session = Depends(get_db),
):
    """Cambio de contraseña para admin, master y general (usuario logueado)."""
    user = requiere_login(request)
    if len(password_nueva) < 4:
        raise HTTPException(400, "La nueva contraseña debe tener al menos 4 caracteres")
    if password_nueva != password_confirmar:
        raise HTTPException(400, "La confirmación no coincide")
    u = db.query(Usuario).filter(Usuario.id == user["user_id"]).first()
    if not u:
        raise HTTPException(404, "Usuario no encontrado")
    if not verify_password(password_actual, u.password_hash, u.salt):
        raise HTTPException(400, "Contraseña actual incorrecta")
    h, s = hash_password(password_nueva)
    u.password_hash = h
    u.salt = s
    db.commit()
    return {"ok": True, "mensaje": "Contraseña actualizada correctamente"}

# ==================== EDIFICIOS ====================

@app.get("/api/edificios")
async def listar_edificios(request: Request, db: Session = Depends(get_db)):
    requiere_login(request)
    return [{"id": e.id, "nombre": e.nombre, "codigo": e.codigo, "manual": e.manual}
            for e in db.query(Edificio).filter(Edificio.activo == True).order_by(Edificio.nombre).all()]


@app.post("/api/edificios")
async def crear_edificio(request: Request, nombre: str = Form(...), codigo: Optional[str] = Form(None), db: Session = Depends(get_db)):
    requiere_login(request)
    nombre = nombre.strip()
    if not nombre:
        raise HTTPException(400, "Nombre obligatorio")
    if db.query(Edificio).filter(Edificio.nombre == nombre).first():
        raise HTTPException(400, "Ese edificio ya existe")
    e = Edificio(nombre=nombre, codigo=(codigo or "").strip() or None, manual=True, activo=True)
    db.add(e)
    db.commit()
    return {"ok": True, "edificio": {"id": e.id, "nombre": e.nombre}}


# ==================== ARCHIVOS EXCEL (máx 5) ====================

@app.post("/api/upload")
async def upload_archivo(request: Request, file: UploadFile = File(...), db: Session = Depends(get_db)):
    user = requiere_master_o_admin(request)
    if not file.filename.lower().endswith((".xlsx", ".xls")):
        raise HTTPException(400, "Solo Excel (.xlsx)")

    activos = db.query(ArchivoHorario).filter(ArchivoHorario.activo == True).count()
    if activos >= MAX_ARCHIVOS_EXCEL:
        raise HTTPException(400, f"Máximo {MAX_ARCHIVOS_EXCEL} archivos de horario. Elimina uno antes de subir otro.")

    # Nombre único en disco
    safe_name = f"{secrets.token_hex(4)}_{file.filename}"
    ruta = os.path.join(UPLOAD_DIR, safe_name)
    with open(ruta, "wb") as f:
        shutil.copyfileobj(file.file, f)

    try:
        aulas_data = parsear_aulas_desde_excel(ruta)
    except Exception as e:
        try:
            os.remove(ruta)
        except Exception:
            pass
        raise HTTPException(400, f"Error al leer Excel: {e}")

    if not aulas_data:
        try:
            os.remove(ruta)
        except Exception:
            pass
        raise HTTPException(400, "No se encontraron aulas en el archivo")

    arch = ArchivoHorario(
        nombre_archivo=file.filename,
        total_aulas=len(aulas_data),
        ruta=ruta,
        fecha_carga=datetime.now(),
        cargado_por=user["user_id"],
        activo=True,
    )
    db.add(arch)
    db.flush()

    edificios_vistos = set()
    for a in aulas_data:
        ed_nombre = a.get("edificio")
        if ed_nombre and ed_nombre not in edificios_vistos:
            edificios_vistos.add(ed_nombre)
            if not db.query(Edificio).filter(Edificio.nombre == ed_nombre).first():
                db.add(Edificio(nombre=ed_nombre, manual=False, activo=True))

        # Misma aula puede venir de varios archivos: se actualiza o se crea con archivo_id
        existente = db.query(Aula).filter(Aula.codigo == a["codigo"], Aula.archivo_id == arch.id).first()
        if not existente:
            existente = db.query(Aula).filter(Aula.codigo == a["codigo"], Aula.activa == True).first()

        if existente and existente.archivo_id == arch.id:
            existente.nombre = a["nombre"]
            existente.edificio = a["edificio"]
            existente.capacidad = a["capacidad"]
            existente.horario_json = json.dumps(a.get("horario", {}), ensure_ascii=False)
            existente.dias_habilitados_json = json.dumps(a.get("dias_habilitados", []), ensure_ascii=False)
        else:
            db.add(Aula(
                codigo=a["codigo"], nombre=a["nombre"], edificio=a["edificio"],
                capacidad=a["capacidad"], activa=True, manual=False,
                horario_json=json.dumps(a.get("horario", {}), ensure_ascii=False),
                dias_habilitados_json=json.dumps(a.get("dias_habilitados", []), ensure_ascii=False),
                archivo_id=arch.id,
            ))

    db.commit()
    return {
        "ok": True,
        "mensaje": f"Archivo cargado ({activos + 1}/{MAX_ARCHIVOS_EXCEL}). {len(aulas_data)} aulas.",
        "total_aulas": len(aulas_data),
        "archivo_id": arch.id,
    }


@app.post("/api/eliminar-archivo/{archivo_id}")
async def eliminar_archivo(archivo_id: int, request: Request, db: Session = Depends(get_db)):
    requiere_master_o_admin(request)
    archivo = db.query(ArchivoHorario).filter(ArchivoHorario.id == archivo_id).first()
    if not archivo:
        raise HTTPException(404, "Archivo no encontrado")
    if archivo.ruta and os.path.exists(archivo.ruta):
        try:
            os.remove(archivo.ruta)
        except Exception:
            pass
    db.query(Aula).filter(Aula.archivo_id == archivo_id).delete()
    db.delete(archivo)
    db.commit()
    return {"ok": True, "mensaje": "Archivo eliminado"}



# ==================== RESERVAS DE AULAS ====================

@app.post("/api/upload-reserva")
async def upload_reserva(request: Request, file: UploadFile = File(...), db: Session = Depends(get_db)):
    """Carga Excel de reservas (área, cédula, nombre, fechas, sede, aula)."""
    user = requiere_master_o_admin(request)
    name = (file.filename or "").lower()
    if not name.endswith((".xlsx", ".xls")):
        raise HTTPException(400, "Sube un Excel de reservas (.xlsx). Las imágenes aún no se procesan automáticamente.")

    safe_name = f"res_{secrets.token_hex(4)}_{file.filename}"
    ruta = os.path.join(UPLOAD_DIR, safe_name)
    with open(ruta, "wb") as f:
        shutil.copyfileobj(file.file, f)

    try:
        filas = parsear_reservas_desde_excel(ruta)
    except Exception as e:
        try:
            os.remove(ruta)
        except Exception:
            pass
        raise HTTPException(400, f"Error al leer reservas: {e}")

    if not filas:
        try:
            os.remove(ruta)
        except Exception:
            pass
        raise HTTPException(400, "No se encontraron reservas en el archivo")

    arch = ArchivoReserva(
        nombre_archivo=file.filename,
        total_reservas=len(filas),
        ruta=ruta,
        fecha_carga=datetime.now(),
        cargado_por=user["user_id"],
        activo=True,
    )
    db.add(arch)
    db.flush()

    for frow in filas:
        db.add(ReservaAula(
            area_solicitante=frow.get("area_solicitante") or None,
            cedula=frow.get("cedula") or None,
            nombre_persona=frow.get("nombre_persona") or None,
            dia=frow.get("dia") or None,
            fecha_ini=frow.get("fecha_ini") or None,
            hora_ini=frow.get("hora_ini") or None,
            fecha_fin=frow.get("fecha_fin") or None,
            hora_fin=frow.get("hora_fin") or None,
            sede=frow.get("sede") or None,
            aula_codigo=frow["aula_codigo"],
            archivo_id=arch.id,
            activo=True,
        ))
    db.commit()
    return {"ok": True, "mensaje": f"Reservas cargadas: {len(filas)}", "total": len(filas), "archivo_id": arch.id}


@app.get("/api/archivos-reserva")
async def listar_archivos_reserva(request: Request, db: Session = Depends(get_db)):
    requiere_login(request)
    archivos = db.query(ArchivoReserva).filter(ArchivoReserva.activo == True).order_by(ArchivoReserva.fecha_carga.desc()).all()
    return {
        "total": len(archivos),
        "archivos": [
            {"id": a.id, "nombre_archivo": a.nombre_archivo, "total_reservas": a.total_reservas,
             "fecha_carga": a.fecha_carga.strftime("%Y-%m-%d %H:%M") if a.fecha_carga else ""}
            for a in archivos
        ],
    }


@app.delete("/api/archivos-reserva/{archivo_id}")
async def eliminar_archivo_reserva(archivo_id: int, request: Request, db: Session = Depends(get_db)):
    requiere_master_o_admin(request)
    arch = db.query(ArchivoReserva).filter(ArchivoReserva.id == archivo_id).first()
    if not arch:
        raise HTTPException(404, "Archivo no encontrado")
    arch.activo = False
    db.query(ReservaAula).filter(ReservaAula.archivo_id == archivo_id).update({"activo": False})
    db.commit()
    return {"ok": True}


@app.get("/api/reservas")
async def buscar_reservas(
    request: Request,
    q: str = Query("", min_length=0),
    aula: Optional[str] = Query(None),
    cedula: Optional[str] = Query(None),
    limit: int = Query(50, ge=1, le=200),
    db: Session = Depends(get_db),
):
    """Busca reservas por aula, nombre, cédula o sede."""
    requiere_login(request)
    query = db.query(ReservaAula).filter(ReservaAula.activo == True)
    if aula:
        query = query.filter(ReservaAula.aula_codigo.ilike(f"%{aula}%"))
    if cedula:
        query = query.filter(ReservaAula.cedula.ilike(f"%{cedula}%"))
    if q:
        like = f"%{q}%"
        query = query.filter(
            (ReservaAula.aula_codigo.ilike(like))
            | (ReservaAula.nombre_persona.ilike(like))
            | (ReservaAula.cedula.ilike(like))
            | (ReservaAula.sede.ilike(like))
            | (ReservaAula.area_solicitante.ilike(like))
        )
    rows = query.order_by(ReservaAula.id.desc()).limit(limit).all()
    return [
        {
            "id": r.id,
            "area_solicitante": r.area_solicitante or "",
            "cedula": r.cedula or "",
            "nombre_persona": r.nombre_persona or "",
            "dia": r.dia or "",
            "fecha_ini": r.fecha_ini or "",
            "hora_ini": r.hora_ini or "",
            "fecha_fin": r.fecha_fin or "",
            "hora_fin": r.hora_fin or "",
            "sede": r.sede or "",
            "aula_codigo": r.aula_codigo,
        }
        for r in rows
    ]


@app.get("/api/archivos")
async def listar_archivos(request: Request, db: Session = Depends(get_db)):
    requiere_login(request)
    archivos = db.query(ArchivoHorario).filter(ArchivoHorario.activo == True).order_by(ArchivoHorario.fecha_carga.desc()).all()
    return {
        "max": MAX_ARCHIVOS_EXCEL,
        "total": len(archivos),
        "archivos": [
            {"id": a.id, "nombre": a.nombre_archivo, "total_aulas": a.total_aulas,
             "fecha": a.fecha_carga.strftime("%Y-%m-%d %H:%M") if a.fecha_carga else ""}
            for a in archivos
        ],
    }


# ==================== AULAS ====================

@app.get("/api/aulas")
async def listar_aulas(request: Request, q: Optional[str] = Query(None), edificio: Optional[str] = Query(None), db: Session = Depends(get_db)):
    requiere_login(request)
    query = db.query(Aula).filter(Aula.activa == True)
    if q:
        like = f"%{q.strip()}%"
        query = query.filter((Aula.codigo.ilike(like)) | (Aula.nombre.ilike(like)))
    if edificio:
        query = query.filter(Aula.edificio == edificio)
    # Dedupe por código (puede haber misma aula en varios archivos)
    seen = set()
    out = []
    for a in query.order_by(Aula.codigo).all():
        if a.codigo in seen:
            continue
        seen.add(a.codigo)
        out.append({
            "codigo": a.codigo, "nombre": a.nombre, "edificio": a.edificio,
            "capacidad": a.capacidad, "dias_habilitados": a.get_dias_habilitados(), "manual": a.manual,
        })
    return out


@app.get("/api/aula/{codigo}")
async def detalle_aula(codigo: str, request: Request, db: Session = Depends(get_db)):
    requiere_login(request)
    aula = db.query(Aula).filter(Aula.codigo == codigo, Aula.activa == True).first()
    if not aula:
        raise HTTPException(404, "Aula no encontrada")
    return {
        "codigo": aula.codigo, "nombre": aula.nombre, "edificio": aula.edificio,
        "capacidad": aula.capacidad, "dias_habilitados": aula.get_dias_habilitados(),
        "horario": aula.get_horario(), "manual": aula.manual,
    }


# ==================== REGISTRO APERTURA ====================

@app.post("/api/registrar")
async def registrar_apertura(
    request: Request,
    aula_codigo: str = Form(...), docente_nombre: str = Form(...), cedula: str = Form(...),
    registrador_nombre: str = Form(...), motivo: str = Form(...),
    observaciones: Optional[str] = Form(None), edificio: Optional[str] = Form(None),
    db: Session = Depends(get_db),
):
    user = requiere_login(request)
    aula_codigo = aula_codigo.strip()
    motivo = motivo.strip()
    if not motivo or not aula_codigo:
        raise HTTPException(400, "Aula y motivo son obligatorios")

    aula = db.query(Aula).filter(Aula.codigo == aula_codigo, Aula.activa == True).first()
    edificio_val = (edificio or "").strip() or None
    if aula:
        aula_nombre = aula.nombre
        # Preferir lo que escribió el usuario; si vacío, el del Excel
        if not edificio_val:
            edificio_val = aula.edificio
    else:
        # Aula no en Excel: solo registro de apertura (no se crea aula ni edificio en BD)
        aula_nombre = f"Aula {aula_codigo}"
        if not edificio_val:
            raise HTTPException(400, "Indica el número o nombre del edificio")

    registro = RegistroApertura(
        aula_codigo=aula_codigo, aula_nombre=aula_nombre, edificio=edificio_val,
        docente_nombre=docente_nombre.strip(), cedula=cedula.strip(),
        registrador_nombre=registrador_nombre.strip(), motivo=motivo,
        observaciones=(observaciones or "").strip() or None,
        fecha_hora=datetime.now(), usuario_id=user["user_id"],
        usuario_username=user["username"], usuario_nombre=user["nombre"],
    )
    db.add(registro)
    db.commit()
    db.refresh(registro)
    return {
        "ok": True, "mensaje": "Apertura registrada",
        "registro": {
            "id": registro.id, "aula_codigo": registro.aula_codigo, "aula_nombre": registro.aula_nombre,
            "edificio": registro.edificio, "docente_nombre": registro.docente_nombre,
            "motivo": registro.motivo, "usuario_nombre": registro.usuario_nombre,
            "fecha_hora": registro.fecha_hora.strftime("%Y-%m-%d %H:%M:%S"),
        },
    }


@app.get("/api/registros")
async def listar_registros(
    request: Request, limit: int = Query(100, ge=1, le=500),
    usuario_id: Optional[int] = Query(None), edificio: Optional[str] = Query(None),
    db: Session = Depends(get_db),
):
    user = requiere_login(request)
    q = db.query(RegistroApertura)
    if usuario_id and user["rol"] in ("admin", "master"):
        q = q.filter(RegistroApertura.usuario_id == usuario_id)
    if edificio:
        q = q.filter(RegistroApertura.edificio == edificio)
    registros = q.order_by(RegistroApertura.fecha_hora.desc()).limit(limit).all()
    return [
        {
            "id": r.id, "aula_codigo": r.aula_codigo, "aula_nombre": r.aula_nombre,
            "edificio": r.edificio or "", "docente_nombre": r.docente_nombre, "cedula": r.cedula,
            "registrador_nombre": r.registrador_nombre, "motivo": r.motivo or "",
            "observaciones": r.observaciones or "",
            "usuario_username": r.usuario_username or "",
            "usuario_nombre": r.usuario_nombre or r.registrador_nombre or "",
            "fecha_hora": r.fecha_hora.strftime("%Y-%m-%d %H:%M:%S"),
        }
        for r in registros
    ]


@app.get("/api/indicadores")
async def indicadores(request: Request, db: Session = Depends(get_db)):
    requiere_master_o_admin(request)
    por_edificio = db.query(RegistroApertura.edificio, func.count(RegistroApertura.id)).group_by(RegistroApertura.edificio).all()
    por_usuario = (
        db.query(RegistroApertura.usuario_nombre, func.count(RegistroApertura.id))
        .group_by(RegistroApertura.usuario_nombre)
        .order_by(func.count(RegistroApertura.id).desc()).all()
    )
    # Aulas con más aperturas (ranking)
    por_aula = (
        db.query(
            RegistroApertura.aula_codigo,
            RegistroApertura.aula_nombre,
            RegistroApertura.edificio,
            func.count(RegistroApertura.id).label("total"),
        )
        .group_by(RegistroApertura.aula_codigo, RegistroApertura.aula_nombre, RegistroApertura.edificio)
        .order_by(func.count(RegistroApertura.id).desc())
        .limit(15)
        .all()
    )
    return {
        "edificios": [{"edificio": e or "Sin edificio", "total": c} for e, c in por_edificio],
        "usuarios": [{"nombre": n or "Sin nombre", "total": c} for n, c in por_usuario],
        "aulas": [
            {
                "codigo": a.aula_codigo,
                "nombre": a.aula_nombre or "",
                "edificio": a.edificio or "Sin edificio",
                "total": a.total,
            }
            for a in por_aula
        ],
    }


@app.get("/api/rendimiento")
async def api_rendimiento(request: Request, db: Session = Depends(get_db)):
    requiere_master_o_admin(request)
    hoy = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    semana = hoy - timedelta(days=7)
    mes = hoy - timedelta(days=30)
    usuarios = db.query(Usuario).filter(Usuario.activo == True).all()
    out = []
    for u in usuarios:
        total = db.query(RegistroApertura).filter(RegistroApertura.usuario_id == u.id).count()
        out.append({
            "id": u.id, "nombre": u.nombre, "username": u.username, "rol": u.rol,
            "total": total,
            "hoy": db.query(RegistroApertura).filter(RegistroApertura.usuario_id == u.id, RegistroApertura.fecha_hora >= hoy).count(),
            "semana": db.query(RegistroApertura).filter(RegistroApertura.usuario_id == u.id, RegistroApertura.fecha_hora >= semana).count(),
            "mes": db.query(RegistroApertura).filter(RegistroApertura.usuario_id == u.id, RegistroApertura.fecha_hora >= mes).count(),
        })
    out.sort(key=lambda x: x["total"], reverse=True)
    return {"total_global": db.query(RegistroApertura).count(), "usuarios": out}


# ==================== REPORTES ====================

@app.get("/api/exportar/excel")
async def exportar_excel(
    request: Request, fecha_inicio: date = Query(...), fecha_fin: date = Query(...),
    usuario_id: Optional[int] = Query(None), edificio: Optional[str] = Query(None),
    db: Session = Depends(get_db),
):
    user = requiere_login(request)
    if fecha_inicio > fecha_fin:
        raise HTTPException(400, "Fechas inválidas")
    inicio = datetime.combine(fecha_inicio, datetime.min.time())
    fin = datetime.combine(fecha_fin, datetime.max.time())
    q = db.query(RegistroApertura).filter(RegistroApertura.fecha_hora >= inicio, RegistroApertura.fecha_hora <= fin)
    if user["rol"] in ("admin", "master"):
        if usuario_id:
            q = q.filter(RegistroApertura.usuario_id == usuario_id)
        if edificio:
            q = q.filter(RegistroApertura.edificio == edificio)
    else:
        q = q.filter(RegistroApertura.usuario_id == user["user_id"])
    registros = q.order_by(RegistroApertura.fecha_hora.asc()).all()

    wb = Workbook()
    ws = wb.active
    ws.title = "Registros"
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="1F4E79")
    thin = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
    headers = ["ID", "Código", "Aula", "Edificio", "Docente", "Cédula", "Registra", "Usuario (nombre)", "Motivo", "Observaciones", "Fecha"]
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = thin
    for ri, r in enumerate(registros, 2):
        for ci, v in enumerate([
            r.id, r.aula_codigo, r.aula_nombre or "", r.edificio or "", r.docente_nombre, r.cedula,
            r.registrador_nombre, r.usuario_nombre or r.usuario_username or "", r.motivo or "",
            r.observaciones or "", r.fecha_hora.strftime("%Y-%m-%d %H:%M:%S"),
        ], 1):
            cell = ws.cell(row=ri, column=ci, value=v)
            cell.border = thin
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return StreamingResponse(buffer, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="registros_{fecha_inicio}_{fecha_fin}.xlsx"'})


@app.get("/api/exportar/pdf")
async def exportar_pdf(
    request: Request, fecha_inicio: date = Query(...), fecha_fin: date = Query(...),
    usuario_id: Optional[int] = Query(None), edificio: Optional[str] = Query(None),
    db: Session = Depends(get_db),
):
    user = requiere_login(request)
    if fecha_inicio > fecha_fin:
        raise HTTPException(400, "Fechas inválidas")
    inicio = datetime.combine(fecha_inicio, datetime.min.time())
    fin = datetime.combine(fecha_fin, datetime.max.time())
    q = db.query(RegistroApertura).filter(RegistroApertura.fecha_hora >= inicio, RegistroApertura.fecha_hora <= fin)
    if user["rol"] in ("admin", "master"):
        if usuario_id:
            q = q.filter(RegistroApertura.usuario_id == usuario_id)
        if edificio:
            q = q.filter(RegistroApertura.edificio == edificio)
    else:
        q = q.filter(RegistroApertura.usuario_id == user["user_id"])
    registros = q.order_by(RegistroApertura.fecha_hora.asc()).all()

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(letter), leftMargin=0.3*inch, rightMargin=0.3*inch, topMargin=0.4*inch, bottomMargin=0.4*inch)
    styles = getSampleStyleSheet()
    elements = [
        Paragraph("Registro de Apertura de Aulas", ParagraphStyle("T", parent=styles["Heading1"], fontSize=13, alignment=1)),
        Paragraph(f"{fecha_inicio} — {fecha_fin} | Total: {len(registros)}", ParagraphStyle("S", parent=styles["Normal"], fontSize=9, alignment=1, spaceAfter=8)),
    ]
    data = [["ID", "Código", "Edificio", "Docente", "Usuario", "Motivo", "Fecha"]]
    for r in registros:
        data.append([str(r.id), r.aula_codigo, (r.edificio or "")[:20], r.docente_nombre[:16],
                     (r.usuario_nombre or r.usuario_username or "")[:16], (r.motivo or "")[:18],
                     r.fecha_hora.strftime("%Y-%m-%d %H:%M")])
    table = Table(data, repeatRows=1)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F4E79")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTSIZE", (0, 0), (-1, -1), 7),
        ("GRID", (0, 0), (-1, -1), 0.3, colors.grey),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F2F2F2")]),
    ]))
    elements.append(table)
    doc.build(elements)
    buffer.seek(0)
    return StreamingResponse(buffer, media_type="application/pdf",
        headers={"Content-Disposition": f'attachment; filename="registros_{fecha_inicio}_{fecha_fin}.pdf"'})


# Health check para Vercel
@app.get("/api/health")
async def health():
    return {"status": "ok", "version": "3.0"}


if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", host="0.0.0.0", port=int(os.environ.get("PORT", 8000)), reload=True)

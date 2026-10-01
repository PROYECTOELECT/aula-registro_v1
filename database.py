from sqlalchemy import create_engine, Column, Integer, String, DateTime, Boolean, Text, ForeignKey
from sqlalchemy.orm import declarative_base, sessionmaker
from datetime import datetime
import os
import json
import hashlib
import secrets

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# --- Base de datos: PostgreSQL en producción (Vercel/Neon) o SQLite en local ---
DATABASE_URL = os.environ.get("DATABASE_URL", "").strip()
if DATABASE_URL:
    # Neon/Vercel a veces entregan postgres:// — SQLAlchemy 2 prefiere postgresql://
    if DATABASE_URL.startswith("postgres://"):
        DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql://", 1)
    print("[db] PostgreSQL:", DATABASE_URL.split("@")[-1] if "@" in DATABASE_URL else "ok")
    from sqlalchemy.pool import NullPool
    engine = create_engine(
        DATABASE_URL,
        pool_pre_ping=True,
        poolclass=NullPool,  # mejor para serverless (Vercel)
    )
else:
    DB_PATH = os.path.join(BASE_DIR, "aula_registro.db")
    engine = create_engine(
        f"sqlite:///{DB_PATH}",
        connect_args={"check_same_thread": False},
    )

SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)
Base = declarative_base()

MAX_ARCHIVOS_EXCEL = 5


def hash_password(password: str, salt: str = None) -> tuple:
    if salt is None:
        salt = secrets.token_hex(16)
    h = hashlib.sha256((salt + password).encode("utf-8")).hexdigest()
    return h, salt


def verify_password(password: str, password_hash: str, salt: str) -> bool:
    h, _ = hash_password(password, salt)
    return h == password_hash


class Usuario(Base):
    __tablename__ = "usuarios"
    id = Column(Integer, primary_key=True, index=True)
    username = Column(String(80), unique=True, index=True, nullable=False)
    password_hash = Column(String(128), nullable=False)
    salt = Column(String(64), nullable=False)
    nombre = Column(String(150), nullable=False)
    rol = Column(String(20), nullable=False, default="general")
    activo = Column(Boolean, default=True)
    cupo_max = Column(Integer, nullable=True, default=0)
    creado_por = Column(Integer, nullable=True)
    codigo_invitacion = Column(String(32), nullable=True, unique=True, index=True)
    foto = Column(String(500), nullable=True)
    creado_en = Column(DateTime, default=datetime.now)


class Edificio(Base):
    __tablename__ = "edificios"
    id = Column(Integer, primary_key=True, index=True)
    nombre = Column(String(255), unique=True, nullable=False, index=True)
    codigo = Column(String(50), nullable=True)
    manual = Column(Boolean, default=False)
    activo = Column(Boolean, default=True)


class ArchivoHorario(Base):
    """Hasta MAX_ARCHIVOS_EXCEL archivos de horario activos."""
    __tablename__ = "archivo_horario"
    id = Column(Integer, primary_key=True)
    nombre_archivo = Column(String(255), nullable=False)
    fecha_carga = Column(DateTime, default=datetime.now)
    total_aulas = Column(Integer, default=0)
    ruta = Column(String(500), nullable=True)
    cargado_por = Column(Integer, nullable=True)
    activo = Column(Boolean, default=True)


class Aula(Base):
    __tablename__ = "aulas"
    id = Column(Integer, primary_key=True, index=True)
    codigo = Column(String(50), index=True, nullable=False)
    nombre = Column(String(255), nullable=False)
    edificio = Column(String(255), nullable=True)
    capacidad = Column(Integer, nullable=True)
    activa = Column(Boolean, default=True)
    horario_json = Column(Text, nullable=True)
    dias_habilitados_json = Column(Text, nullable=True)
    manual = Column(Boolean, default=False)
    archivo_id = Column(Integer, nullable=True, index=True)

    def get_horario(self):
        if not self.horario_json:
            return {}
        try:
            return json.loads(self.horario_json)
        except Exception:
            return {}

    def get_dias_habilitados(self):
        if not self.dias_habilitados_json:
            return []
        try:
            return json.loads(self.dias_habilitados_json)
        except Exception:
            return []


class RegistroApertura(Base):
    __tablename__ = "registros_apertura"
    id = Column(Integer, primary_key=True, index=True)
    aula_codigo = Column(String(50), nullable=False, index=True)
    aula_nombre = Column(String(255), nullable=True)
    edificio = Column(String(255), nullable=True, index=True)
    docente_nombre = Column(String(255), nullable=False)
    cedula = Column(String(50), nullable=False)
    registrador_nombre = Column(String(255), nullable=False)
    motivo = Column(String(255), nullable=False, default="")
    observaciones = Column(Text, nullable=True)
    fecha_hora = Column(DateTime, default=datetime.now, index=True)
    usuario_id = Column(Integer, nullable=True, index=True)
    usuario_username = Column(String(80), nullable=True)
    usuario_nombre = Column(String(150), nullable=True)



class ArchivoReserva(Base):
    """Archivos Excel de reservas de aulas."""
    __tablename__ = "archivo_reserva"
    id = Column(Integer, primary_key=True)
    nombre_archivo = Column(String(255), nullable=False)
    fecha_carga = Column(DateTime, default=datetime.now)
    total_reservas = Column(Integer, default=0)
    ruta = Column(String(500), nullable=True)
    cargado_por = Column(Integer, nullable=True)
    activo = Column(Boolean, default=True)


class ReservaAula(Base):
    """Reserva de aula cargada desde Excel de reservas."""
    __tablename__ = "reservas_aula"
    id = Column(Integer, primary_key=True, index=True)
    area_solicitante = Column(String(255), nullable=True)
    cedula = Column(String(50), nullable=True, index=True)
    nombre_persona = Column(String(255), nullable=True, index=True)
    dia = Column(String(30), nullable=True)
    fecha_ini = Column(String(30), nullable=True)
    hora_ini = Column(String(20), nullable=True)
    fecha_fin = Column(String(30), nullable=True)
    hora_fin = Column(String(20), nullable=True)
    sede = Column(String(50), nullable=True, index=True)
    aula_codigo = Column(String(50), nullable=False, index=True)
    archivo_id = Column(Integer, nullable=True, index=True)
    activo = Column(Boolean, default=True)


def init_db():
    Base.metadata.create_all(bind=engine)
    db = SessionLocal()
    try:
        if db.query(Usuario).count() == 0:
            h, s = hash_password("admin123")
            db.add(Usuario(
                username="admin",
                password_hash=h,
                salt=s,
                nombre="Administrador",
                rol="admin",
                activo=True,
            ))
            db.commit()
            print("Usuario admin creado: admin / admin123")
    finally:
        db.close()


def contar_generales_de_master(db, master_id: int) -> int:
    return db.query(Usuario).filter(
        Usuario.creado_por == master_id,
        Usuario.rol == "general",
    ).count()


def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()

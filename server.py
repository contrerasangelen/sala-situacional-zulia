import http.server
import socketserver
import socket
import sys
import json
import sqlite3
import os
import glob
import re
import urllib.parse
from datetime import datetime, timedelta, timezone

def get_vet_time():
    """Retorna la fecha y hora oficial de Venezuela (UTC-4) para servidores en la nube como Render."""
    return datetime.now(timezone.utc) - timedelta(hours=4)
from collections import defaultdict
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

PORT = int(os.environ.get("PORT", 8000))
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

DB_PATH = os.environ.get("SALA_DB", os.path.join(BASE_DIR, "sala_situacional.db"))
PUBLIC_DIR = os.path.join(BASE_DIR, "public")
DOWNLOADS_DIR = os.path.join(os.path.expanduser("~"), "Downloads")
WORKSPACE_DIR = os.path.join(os.path.expanduser("~"), "Desktop")
if not os.path.exists(WORKSPACE_DIR):
    WORKSPACE_DIR = BASE_DIR

def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def init_db_schema():
    conn = get_db()
    cur = conn.cursor()
    try:
        cols = [r[1] for r in cur.execute("PRAGMA table_info(casos)").fetchall()]
        if "consecutivo_linea" not in cols:
            cur.execute("ALTER TABLE casos ADD COLUMN consecutivo_linea INTEGER DEFAULT 0")
        if "consecutivo_impreso" not in cols:
            cur.execute("ALTER TABLE casos ADD COLUMN consecutivo_impreso INTEGER DEFAULT 0")
        if "consecutivo_fecha_impresion" not in cols:
            cur.execute("ALTER TABLE casos ADD COLUMN consecutivo_fecha_impresion TEXT")
        if "desestimado_aprobado" not in cols:
            cur.execute("ALTER TABLE casos ADD COLUMN desestimado_aprobado INTEGER DEFAULT 0")
        if "inconsistencia_aprobada" not in cols:
            cur.execute("ALTER TABLE casos ADD COLUMN inconsistencia_aprobada INTEGER DEFAULT 0")

        if "operador_nombre" not in cols:
            cur.execute("ALTER TABLE casos ADD COLUMN operador_nombre TEXT")

        # Column 'grupo' in operadores
        op_cols = [r[1] for r in cur.execute("PRAGMA table_info(operadores)").fetchall()]
        if "grupo" not in op_cols:
            cur.execute("ALTER TABLE operadores ADD COLUMN grupo TEXT DEFAULT 'PRINCIPAL'")
        cur.execute("UPDATE operadores SET grupo = 'PRINCIPAL' WHERE id <= 15")
        cur.execute("UPDATE operadores SET grupo = 'HISTORICO' WHERE id = 99")
        for i in range(16, 26):
            num_str = f"{i-15:02d}"
            cur.execute("""
                INSERT INTO operadores (id, nombre, pestaña, pin, activo, grupo)
                VALUES (?, ?, ?, '1234', 1, 'APOYO')
                ON CONFLICT(id) DO UPDATE SET grupo = 'APOYO'
            """, (i, f"APOYO {num_str}", f"AP{num_str} - APOYO {i-15}"))
        conn.commit()
    except Exception as e:
        print("Schema init notice:", e)
    finally:
        conn.close()

init_db_schema()


# Formateo nacional de teléfonos
def format_phone(val):
    if not val:
        return ""
    s = ''.join(c for c in str(val) if c.isdigit())
    if not s:
        return ""
    if s.startswith('58') and len(s) == 12:
        s = s[2:]
    elif s.startswith('58') and len(s) in [2, 3]:
        return ""
    if len(s) == 10:
        return f"0{s[0:3]}-{s[3:6]}-{s[6:8]}-{s[8:10]}"
    if len(s) == 11 and s.startswith('0'):
        return f"{s[0:4]}-{s[4:7]}-{s[7:9]}-{s[9:11]}"
    if len(s) == 7:
        return f"0261-{s[0:3]}-{s[3:5]}-{s[5:7]}"
    return str(val).strip()

# Motor de Auditoría y Control de Calidad
def audit_case_coherence(cat_hom, estatus, persiste, requiere_campo, obs):
    obs_clean = (obs or "").lower().strip()
    cat_clean = (cat_hom or "").upper().strip()
    est_clean = (estatus or "").upper().strip()
    per_clean = (persiste or "").upper().strip()
    req_clean = (requiere_campo or "").upper().strip()
    
    warnings = []
    
    # 1. Caso de TX pero menciona poste / pica y poda sin recategorizar
    if "TX" in cat_clean or "TRANSFORMADOR" in cat_clean:
        if any(w in obs_clean for w in ["poste corroido", "poste quebrado", "poste caido", "poste oxidado", "cambio de poste", "poste chocado"]):
            if est_clean != "RECATEGORIZAR":
                warnings.append("El caso original es de Transformador (TX) pero la novedad reportada es daño físico en Poste. Debe marcarse estatus RECATEGORIZAR.")
        elif any(w in obs_clean for w in ["pica y poda", "arbol sobre cable", "ramas sobre linea"]):
            if est_clean not in ["RECATEGORIZAR", "SOLICITA PICA Y PODA"]:
                warnings.append("El caso original es de Transformador (TX) pero la novedad es Pica y Poda. Debe marcarse estatus RECATEGORIZAR o SOLICITA PICA Y PODA.")
                
    # 2. Marcado como ATENDIDO pero dice que sigue la avería
    if est_clean == "ATENDIDO":
        if any(w in obs_clean for w in ["sin luz", "sin electricidad", "no han venido", "sigue la falla", "exploto", "sin servicio", "aun persiste"]):
            warnings.append("Incongruencia: Marcado como ATENDIDO, pero la observación dice que la avería aún persiste o siguen sin servicio.")
        if per_clean in ["POR INSPECCIONAR", "INSPECCIONADO", "POR MATERIAL"]:
            warnings.append("Incongruencia: Marcado como ATENDIDO, pero se indicó que el caso persiste o está pendiente de inspección.")
            
    # 3. Marcado como SIN COMUNICACION pero hay relato del usuario
    if est_clean in ["SIN COMUNICACIÓN", "SIN COMUNICACION"]:
        if len(obs_clean) > 25 and any(w in obs_clean for w in ["dice", "indica", "informo", "manifiesta", "hablo", "atendio el vecino", "senor"]):
            warnings.append("Incongruencia: Marcado como SIN COMUNICACIÓN, pero la observación contiene testimonios o declaraciones del usuario.")
            
    # 4. Falla NO persiste pero solicita campo
    if per_clean in ["NO", "RESUELTO", "ATENDIDO"] and req_clean == "SI":
        warnings.append("Incongruencia: Se indicó que la avería no persiste, pero se marcó que REQUIERE CAMPO (Cuadrilla).")
        
    return warnings

class SituationalRoomHandler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=PUBLIC_DIR, **kwargs)

    def do_GET(self):
        url = urllib.parse.urlparse(self.path)
        path = url.path
        query = urllib.parse.parse_qs(url.query)

        if path.startswith("/api/"):
            self.handle_api_get(path, query)
        else:
            filename_map = {
                "/": "monitor.html",
                "/monitor": "monitor.html",
                "/monitor.html": "monitor.html",
                "/operador": "operador.html",
                "/operador.html": "operador.html",
                "/supervisor": "supervisor.html",
                "/supervisor.html": "supervisor.html",
                "/accesos": "accesos.html",
                "/accesos.html": "accesos.html",
            }
            if path in filename_map:
                fname = filename_map[path]
                target_file = os.path.join(BASE_DIR, fname)
                if not os.path.exists(target_file):
                    target_file = os.path.join(PUBLIC_DIR, fname)
                if os.path.exists(target_file):
                    with open(target_file, "rb") as f:
                        data = f.read()
                    self.send_response(200)
                    self.send_header('Content-Type', 'text/html; charset=utf-8')
                    self.send_header('Cache-Control', 'no-cache, no-store, must-revalidate')
                    self.send_header('Pragma', 'no-cache')
                    self.send_header('Expires', '0')
                    self.send_header('Content-Length', str(len(data)))
                    self.end_headers()
                    self.wfile.write(data)
                    return
            super().do_GET()

    def do_POST(self):
        url = urllib.parse.urlparse(self.path)
        path = url.path
        content_len = int(self.headers.get('Content-Length', 0))
        post_body = self.rfile.read(content_len) if content_len > 0 else b'{}'
        
        try:
            data = json.loads(post_body.decode('utf-8'))
        except Exception:
            data = {}

        if path.startswith("/api/"):
            self.handle_api_post(path, data)
        else:
            self.send_error(404, "Ruta no encontrada")

    def send_json(self, data, status=200):
        res = json.dumps(data, ensure_ascii=False)
        self.send_response(status)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Cache-Control', 'no-cache, no-store, must-revalidate')
        self.send_header('Pragma', 'no-cache')
        self.send_header('Expires', '0')
        self.end_headers()
        self.wfile.write(res.encode('utf-8'))

    def handle_api_get(self, path, query):
        conn = get_db()
        cur = conn.cursor()

        # Filtro de rango de fechas
        periodo = query.get("periodo", ["historico"])[0]
        f_desde = query.get("desde", [""])[0]
        f_hasta = query.get("hasta", [""])[0]

        now = get_vet_time()
        date_filter_sql = ""
        date_params = []

        if periodo == "diario":
            today_str = now.strftime("%Y-%m-%d")
            date_filter_sql = " AND (fecha_contacto = ? OR actualizado_en LIKE ?)"
            date_params = [today_str, f"{today_str}%"]
        elif periodo == "semanal":
            start_week = (now - timedelta(days=7)).strftime("%Y-%m-%d")
            date_filter_sql = " AND (fecha_contacto >= ? OR actualizado_en >= ?)"
            date_params = [start_week, start_week]
        elif periodo == "mensual":
            start_month = (now - timedelta(days=30)).strftime("%Y-%m-%d")
            date_filter_sql = " AND (fecha_contacto >= ? OR actualizado_en >= ?)"
            date_params = [start_month, start_month]
        elif f_desde and f_hasta:
            date_filter_sql = " AND ((fecha_contacto BETWEEN ? AND ?) OR (actualizado_en BETWEEN ? AND ?))"
            date_params = [f_desde, f_hasta, f_desde + " 00:00:00", f_hasta + " 23:59:59"]

        # 1. Estadísticas Globales
        if path == "/api/stats":
            cur.execute("SELECT COUNT(*) FROM casos")
            tot_casos = cur.fetchone()[0]

            cur.execute("SELECT COUNT(DISTINCT id_llamada) FROM clientes")
            tot_clientes = cur.fetchone()[0]

            cur.execute(f"SELECT COUNT(*) FROM casos WHERE estatus_llamada IS NOT NULL AND estatus_llamada != ''{date_filter_sql}", date_params)
            casos_auditados = cur.fetchone()[0]

            cur.execute(f"SELECT COUNT(DISTINCT id_llamada) FROM casos WHERE estatus_llamada IS NOT NULL AND estatus_llamada != ''{date_filter_sql}", date_params)
            clientes_contactados = cur.fetchone()[0]

            cur.execute(f"SELECT COUNT(*) FROM casos WHERE estatus_llamada IN ('ATENDIDO', 'ATENDIDO POR TERCEROS'){date_filter_sql}", date_params)
            casos_atendidos = cur.fetchone()[0]

            cur.execute(f"SELECT COUNT(*) FROM casos WHERE estatus_llamada IN ('SIN COMUNICACIÓN', 'SIN INFORMACION'){date_filter_sql}", date_params)
            casos_sin_com = cur.fetchone()[0]

            cur.execute(f"SELECT COUNT(*) FROM casos WHERE estatus_llamada = 'POR ATENDER'{date_filter_sql}", date_params)
            casos_por_atender = cur.fetchone()[0]

            cur.execute(f"SELECT COUNT(*) FROM casos WHERE requiere_campo = 'SI'{date_filter_sql}", date_params)
            casos_requieren_campo = cur.fetchone()[0]

            cur.execute(f"SELECT COUNT(*) FROM casos WHERE inconsistencia_flag = 1{date_filter_sql}", date_params)
            casos_inconsistentes = cur.fetchone()[0]

            cur.execute(f"SELECT COUNT(*) FROM casos WHERE devuelto_flag = 1{date_filter_sql}", date_params)
            casos_devueltos = cur.fetchone()[0]

            cur.execute(f"SELECT COUNT(*) FROM casos WHERE (estatus_llamada = 'DESESTIMAR' OR persiste = 'DESESTIMAR'){date_filter_sql}", date_params)
            casos_desestimados = cur.fetchone()[0]

            pct_avance = round(casos_auditados / tot_casos * 100, 1) if tot_casos else 0
            pct_efectividad = round(casos_atendidos / casos_auditados * 100, 1) if casos_auditados else 0

            conn.close()
            self.send_json({
                "periodo_activo": periodo,
                "tot_casos": tot_casos,
                "tot_clientes": tot_clientes,
                "casos_auditados": casos_auditados,
                "casos_pendientes": tot_casos - casos_auditados,
                "clientes_contactados": clientes_contactados,
                "casos_atendidos": casos_atendidos,
                "casos_sin_com": casos_sin_com,
                "casos_por_atender": casos_por_atender,
                "casos_requieren_campo": casos_requieren_campo,
                "casos_inconsistentes": casos_inconsistentes,
                "casos_devueltos": casos_devueltos,
                "casos_desestimados": casos_desestimados,
                "pct_avance": pct_avance,
                "pct_efectividad": pct_efectividad
            })

        # 2. Resumen de Operadores por Período
        elif path == "/api/operators_stats":
            sub_date = ""
            if date_filter_sql:
                sub_date = date_filter_sql.replace("fecha_contacto", "c.fecha_contacto").replace("actualizado_en", "c.actualizado_en")

            grupo_param = (query.get("grupo", [None])[0] or "").lower()
            where_ops = "WHERE o.id != 99"
            if grupo_param == "principal":
                where_ops += " AND o.id <= 15"
            elif grupo_param == "apoyo":
                where_ops += " AND o.id > 15 AND o.id != 99"

            cur.execute(f"""
            SELECT 
                o.id, o.nombre, o.pestaña, o.ultima_actividad, o.activo,
                COALESCE(o.grupo, CASE WHEN o.id <= 15 THEN 'PRINCIPAL' ELSE 'APOYO' END) as grupo,
                COUNT(CASE WHEN c.estado_auditoria = 'PENDIENTE' OR (c.estado_auditoria = 'GESTIONADO' AND DATE(c.actualizado_en) = DATE('now', '-4 hours')) THEN 1 END) as tot_casos,
                COUNT(DISTINCT CASE WHEN c.estado_auditoria = 'PENDIENTE' OR (c.estado_auditoria = 'GESTIONADO' AND DATE(c.actualizado_en) = DATE('now', '-4 hours')) THEN c.id_llamada END) as tot_clientes,
                SUM(CASE WHEN c.estatus_llamada IS NOT NULL AND c.estatus_llamada != '' {sub_date} THEN 1 ELSE 0 END) as auditados,
                SUM(CASE WHEN c.estatus_llamada IN ('ATENDIDO', 'ATENDIDO POR TERCEROS') {sub_date} THEN 1 ELSE 0 END) as atendidos,
                SUM(CASE WHEN c.estatus_llamada IN ('SIN COMUNICACIÓN', 'SIN INFORMACION') {sub_date} THEN 1 ELSE 0 END) as sin_com,
                SUM(CASE WHEN c.estatus_llamada = 'POR ATENDER' {sub_date} THEN 1 ELSE 0 END) as por_atender,
                SUM(CASE WHEN c.requiere_campo = 'SI' {sub_date} THEN 1 ELSE 0 END) as campo,
                SUM(CASE WHEN c.inconsistencia_flag = 1 {sub_date} THEN 1 ELSE 0 END) as inconsistencias,
                SUM(CASE WHEN c.devuelto_flag = 1 THEN 1 ELSE 0 END) as devueltos
            FROM operadores o
            LEFT JOIN casos c ON o.id = c.operador_id
            {where_ops}
            GROUP BY o.id
            ORDER BY o.id
            """, date_params * 6 if sub_date else ())
            
            rows = [dict(r) for r in cur.fetchall()]
            
            for r in rows:
                t_c = r["tot_casos"] or 0
                aud = r["auditados"] or 0
                r["pct_avance"] = round(aud / t_c * 100, 1) if t_c else 0
                r["pct_efectividad"] = round((r["atendidos"] or 0) / aud * 100, 1) if aud else 0
                
                is_active = False
                if r["ultima_actividad"]:
                    try:
                        last_t = datetime.strptime(r["ultima_actividad"], "%Y-%m-%d %H:%M:%S")
                        if (now - last_t).total_seconds() < 600:
                            is_active = True
                    except Exception:
                        pass
                r["is_active"] = is_active
                
            conn.close()
            self.send_json(rows)

        # 3. Lista de Archivos Disponibles para Cruce (Downloads y Directorio)
        elif path == "/api/supervisor/available_files":
            found_files = []
            
            # Buscar en Downloads
            if os.path.exists(DOWNLOADS_DIR):
                for f in glob.glob(os.path.join(DOWNLOADS_DIR, "*.xlsx")):
                    found_files.append({"nombre": os.path.basename(f), "ruta": f, "carpeta": "Descargas"})

            # Buscar en Workspace
            if os.path.exists(WORKSPACE_DIR):
                for f in glob.glob(os.path.join(WORKSPACE_DIR, "*.xlsx")):
                    if "SISTEMA" not in f.upper():
                        found_files.append({"nombre": os.path.basename(f), "ruta": f, "carpeta": "Directorio Proyecto"})

            conn.close()
            self.send_json(found_files)

        # 4. Parroquias Críticas
        elif path == "/api/parish_stats":
            cur.execute("""
            SELECT 
                parroquia, municipio,
                COUNT(*) as total,
                SUM(CASE WHEN estatus_llamada IN ('ATENDIDO', 'ATENDIDO POR TERCEROS') THEN 1 ELSE 0 END) as atendidos,
                SUM(CASE WHEN estatus_llamada = 'SIN COMUNICACIÓN' THEN 1 ELSE 0 END) as sin_com,
                SUM(CASE WHEN categoria_homologada LIKE '%TX%' OR categoria_homologada LIKE '%TRANSFORMADOR%' THEN 1 ELSE 0 END) as casos_tx
            FROM casos
            WHERE parroquia IS NOT NULL AND parroquia != ''
            GROUP BY parroquia, municipio
            ORDER BY total DESC
            LIMIT 10
            """)
            rows = [dict(r) for r in cur.fetchall()]
            conn.close()
            self.send_json(rows)

        # 5. Ticker de logs recientes
        elif path == "/api/recent_logs":
            cur.execute("""
            SELECT l.timestamp, l.detalle, o.nombre as operador_nombre
            FROM logs l
            LEFT JOIN operadores o ON l.operador_id = o.id
            ORDER BY l.id DESC
            LIMIT 8
            """)
            rows = [dict(r) for r in cur.fetchall()]
            conn.close()
            self.send_json(rows)

        # 6. Lista de operadores
        elif path == "/api/operators_list":
            grupo_param = (query.get("grupo", [None])[0] or "").lower()
            where_ops = "WHERE id != 99"
            if grupo_param == "principal":
                where_ops += " AND id <= 15"
            elif grupo_param == "apoyo":
                where_ops += " AND id > 15 AND id != 99"
            cur.execute(f"SELECT id, nombre, pestaña, activo, COALESCE(grupo, CASE WHEN id <= 15 THEN 'PRINCIPAL' ELSE 'APOYO' END) as grupo FROM operadores {where_ops} ORDER BY id")
            rows = [dict(r) for r in cur.fetchall()]
            conn.close()
            self.send_json(rows)

        # 7. Cola de clientes para operador
        elif path == "/api/operator_queue":
            op_id = query.get("operator_id", [None])[0]
            if not op_id:
                conn.close()
                self.send_json({"error": "Falta operator_id"}, 400)
                return

            cur.execute("""
            SELECT 
                cli.id_llamada, cli.denunciante, cli.cedula, cli.telefono_1, cli.telefono_2,
                cli.municipio, cli.parroquia, cli.sector, cli.direccion, 
                COALESCE(MAX(c.tipo_contacto), cli.tipo_contacto) as tipo_contacto,
                MAX(c.fecha_generado) as fecha_aviso,
                COUNT(c.id) as num_casos,
                SUM(CASE WHEN c.estado_auditoria = 'GESTIONADO' AND c.devuelto_flag = 0 THEN 1 ELSE 0 END) as auditados_en_grupo,
                MAX(c.devuelto_flag) as has_devueltos,
                SUM(CASE WHEN c.devuelto_flag = 1 THEN 1 ELSE 0 END) as devueltos_count,
                MAX(CASE WHEN c.devuelto_flag = 1 THEN c.codigo_caso ELSE NULL END) as devuelto_codigo,
                MAX(CASE WHEN c.devuelto_flag = 1 THEN c.devuelto_motivo ELSE NULL END) as devuelto_motivo
            FROM clientes cli
            JOIN casos c ON cli.id_llamada = c.id_llamada
            WHERE c.operador_id = ? AND (
                c.devuelto_flag = 1
                OR c.estado_auditoria IN ('PENDIENTE', 'DEVUELTO_CORRECCION')
                OR (c.estado_auditoria = 'GESTIONADO' AND DATE(c.actualizado_en) = DATE('now', '-4 hours'))
            )
            GROUP BY cli.id_llamada
            ORDER BY 
                has_devueltos DESC,
                (num_casos = auditados_en_grupo) ASC,
                num_casos DESC,
                fecha_aviso DESC,
                cli.id_llamada
            """, (op_id,))
            clientes = [dict(r) for r in cur.fetchall()]
            conn.close()
            self.send_json(clientes)

        # 8. Casos de un cliente (Casos devueltos primero, luego los más recientes)
        elif path == "/api/client_cases":
            id_llamada = query.get("id_llamada", [None])[0]
            if not id_llamada:
                conn.close()
                self.send_json({"error": "Falta id_llamada"}, 400)
                return

            cur.execute("""
            SELECT * FROM casos 
            WHERE id_llamada = ? 
            ORDER BY devuelto_flag DESC, fecha_generado DESC, id DESC
            """, (id_llamada,))
            casos = [dict(r) for r in cur.fetchall()]
            conn.close()
            self.send_json(casos)

        # Lista detallada de operadores para el módulo de Reasignación Rápida
        elif path == "/api/supervisor/operators_reassign_info":
            cur.execute("""
            SELECT 
                o.id, 
                o.nombre,
                COALESCE(o.grupo, CASE WHEN o.id <= 15 THEN 'PRINCIPAL' ELSE 'APOYO' END) as grupo,
                COUNT(DISTINCT CASE WHEN (c.estatus_llamada IS NULL OR c.estatus_llamada = '') THEN c.id_llamada END) as pendientes_llamadas,
                COUNT(CASE WHEN (c.estatus_llamada IS NULL OR c.estatus_llamada = '') THEN 1 END) as pendientes_casos
            FROM operadores o
            LEFT JOIN casos c ON o.id = c.operador_id
            GROUP BY o.id, o.nombre
            ORDER BY o.id
            """)
            ops = [dict(r) for r in cur.fetchall()]
            conn.close()
            self.send_json(ops)

        # 9. Inconsistencias para supervisor
        elif path == "/api/supervisor/inconsistencias":
            cur.execute("""
            SELECT c.*, o.nombre as operador_nombre
            FROM casos c
            JOIN operadores o ON c.operador_id = o.id
            WHERE (c.inconsistencia_flag = 1 OR c.estatus_llamada = 'RECATEGORIZAR')
              AND c.devuelto_flag = 0
              AND (c.inconsistencia_aprobada = 0 OR c.inconsistencia_aprobada IS NULL)
            ORDER BY c.actualizado_en DESC
            """)
            rows = [dict(r) for r in cur.fetchall()]
            conn.close()
            self.send_json(rows)

        # Listar casos de Consecutivos en Línea / Pegasus
        elif path == "/api/consecutivos_list":
            cur.execute("""
            SELECT c.*, o.nombre as operador_nombre
            FROM casos c
            JOIN operadores o ON c.operador_id = o.id
            WHERE c.consecutivo_linea = 1
            ORDER BY c.consecutivo_impreso ASC, c.actualizado_en DESC, c.id DESC
            """)
            rows = [dict(r) for r in cur.fetchall()]
            conn.close()
            self.send_json(rows)

        # Consecutivos en Línea: Contador de Pendientes
        elif path == "/api/consecutivos_pendientes_count":
            cur.execute("SELECT COUNT(*) FROM casos WHERE consecutivo_linea = 1 AND (consecutivo_impreso = 0 OR consecutivo_impreso IS NULL)")
            p = cur.fetchone()[0]
            cur.execute("SELECT COUNT(*) FROM casos WHERE consecutivo_linea = 1")
            tot = cur.fetchone()[0]
            conn.close()
            self.send_json({"pendientes": p, "total": tot})

        # Exportar Consecutivos en Línea a Excel (Descarga directa y marca impresos)
        elif path == "/api/export_consecutivos_linea":
            import io
            tipo = query.get("tipo", ["pendientes"])[0]
            now_dt = datetime.now()
            fecha_str = now_dt.strftime("%Y-%m-%d %H:%M:%S")
            fecha_tag = now_dt.strftime("%Y%m%d_%H%M%S")

            where_sql = "WHERE c.consecutivo_linea = 1"
            if tipo == "pendientes":
                where_sql += " AND (c.consecutivo_impreso = 0 OR c.consecutivo_impreso IS NULL)"

            cur.execute(f"""
            SELECT c.*, o.nombre as operador_nombre
            FROM casos c
            JOIN operadores o ON c.operador_id = o.id
            {where_sql}
            ORDER BY c.actualizado_en DESC, c.municipio, c.parroquia
            """)
            rows = cur.fetchall()

            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Consecutivos Campo"
            ws.views.sheetView[0].showGridLines = True

            headers = [
                "N°", "ID_CASO", "CÓDIGO CASO", "OPERADOR", "MUNICIPIO", "PARROQUIA",
                "SECTOR", "DIRECCIÓN", "POSTE CONFIRMADO", "POSTE DETECTADO",
                "DENUNCIANTE", "CÉDULA", "TELÉFONO 1", "TELÉFONO 2",
                "ESTATUS LLAMADA", "¿PERSISTE?", "REQUIERE CAMPO", "FECHA CONTACTO",
                "CATEGORÍA", "OBSERVACIÓN DETALLADA / NOVEDAD", "FECHA DESPACHO"
            ]
            ws.append(headers)

            font_head = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
            fill_head = PatternFill(start_color="D97706", end_color="D97706", fill_type="solid")
            for c_i in range(1, len(headers) + 1):
                c = ws.cell(row=1, column=c_i)
                c.font = font_head
                c.fill = fill_head
                c.alignment = Alignment(horizontal="center", vertical="center")

            case_ids_to_mark = []
            for idx, r in enumerate(rows, start=1):
                case_ids_to_mark.append(r["id"])
                ws.append([
                    idx,
                    r["id"],
                    r["codigo_caso"],
                    r["operador_nombre"],
                    r["municipio"],
                    r["parroquia"],
                    r["sector"],
                    r["direccion"],
                    r["poste_confirmado"] or "",
                    r["poste_detectado"] or "",
                    r["denunciante"],
                    r["cedula"],
                    r["telefono_1"],
                    r["telefono_2"],
                    r["estatus_llamada"],
                    r["persiste"],
                    r["requiere_campo"],
                    r["fecha_contacto"],
                    r["categoria_homologada"],
                    r["observaciones"],
                    r["consecutivo_fecha_impresion"] or fecha_str
                ])

            for col in ws.columns:
                max_len = max(len(str(cell.value or '')) for cell in col)
                col_letter = get_column_letter(col[0].column)
                ws.column_dimensions[col_letter].width = min(max(max_len + 3, 11), 60)

            if tipo == "pendientes" and case_ids_to_mark:
                placeholders = ",".join(["?"] * len(case_ids_to_mark))
                cur.execute(f"""
                UPDATE casos SET 
                    consecutivo_impreso = 1,
                    consecutivo_fecha_impresion = ?
                WHERE id IN ({placeholders})
                """, [fecha_str] + case_ids_to_mark)
                conn.commit()

            conn.close()

            buf = io.BytesIO()
            wb.save(buf)
            buf.seek(0)
            data_bytes = buf.getvalue()
            wb.close()

            filename = f"CONSECUTIVOS_LINEA_{fecha_tag}.xlsx"
            self.send_response(200)
            self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            self.send_header('Content-Disposition', f'attachment; filename="{filename}"')
            self.send_header('Content-Length', str(len(data_bytes)))
            self.send_header('Access-Control-Allow-Origin', '*')
            self.send_header('Cache-Control', 'no-cache, no-store, must-revalidate')
            self.end_headers()
            self.wfile.write(data_bytes)
            return

        # 10. Buscador de casos
        elif path == "/api/search_case":
            term = query.get("q", [""])[0].strip()
            if not term:
                conn.close()
                self.send_json([])
                return
            q_like = f"%{term}%"
            cur.execute("""
            SELECT c.*, o.nombre as operador_nombre
            FROM casos c
            JOIN operadores o ON c.operador_id = o.id
            WHERE c.codigo_caso LIKE ? OR c.cedula LIKE ? OR c.denunciante LIKE ? OR c.telefono_1 LIKE ?
            LIMIT 50
            """, (q_like, q_like, q_like, q_like))
            rows = [dict(r) for r in cur.fetchall()]
            conn.close()
            self.send_json(rows)

        # Exportar Consolidado General (.xlsx) con descarga directa al navegador
        elif path == "/api/export_excel":
            import io
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Casos y Auditoría"
            ws.views.sheetView[0].showGridLines = True

            cur.execute("""
            SELECT c.*, o.nombre as operador_nombre
            FROM casos c
            JOIN operadores o ON c.operador_id = o.id
            ORDER BY c.operador_id, c.id_llamada
            """)
            rows = cur.fetchall()

            headers = [
                "OPERADOR", "ESTADO_AUDITORIA", "ID_LLAMADA", "CASO_NRO", "MUNICIPIO", "PARROQUIA",
                "CEDULA", "DENUNCIANTE", "TELEFONO_1", "TELEFONO_2", "SECTOR", "DIRECCION",
                "POSTE_DETECTADO", "CODIGO_CASO", "FECHA_GENERADO", "CATEGORIA_HOMOLOGADA",
                "SUBCATEGORIA", "DESCRIPCION_REPORTE", "ESTATUS_LLAMADA", "¿ESTE_CASO_PERSISTE?",
                "REQUIERE_CAMPO", "FECHA_CONTACTO", "OBSERVACIONES", "POSTE_CONFIRMADO",
                "INCONSISTENCIA_ALERTA", "MOTIVO_INCONSISTENCIA", "CONSECUTIVO_LINEA"
            ]
            ws.append(headers)

            font_head = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
            fill_head = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
            for c_i in range(1, len(headers) + 1):
                c = ws.cell(row=1, column=c_i)
                c.font = font_head
                c.fill = fill_head
                c.alignment = Alignment(horizontal="center", vertical="center")

            for r in rows:
                ws.append([
                    r["operador_nombre"], r["estado_auditoria"], r["id_llamada"], r["caso_nro"],
                    r["municipio"], r["parroquia"], r["cedula"], r["denunciante"], r["telefono_1"],
                    r["telefono_2"], r["sector"], r["direccion"], r["poste_detectado"], r["codigo_caso"],
                    r["fecha_generado"], r["categoria_homologada"], r["subcategoria"], r["descripcion_reporte"],
                    r["estatus_llamada"], r["persiste"], r["requiere_campo"], r["fecha_contacto"],
                    r["observaciones"], r["poste_confirmado"], "SÍ" if r["inconsistencia_flag"] else "NO",
                    r["inconsistencia_motivo"], "SÍ" if r["consecutivo_linea"] else "NO"
                ])

            for col in ws.columns:
                max_len = max(len(str(cell.value or '')) for cell in col)
                col_letter = get_column_letter(col[0].column)
                ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 60)

            conn.close()

            buf = io.BytesIO()
            wb.save(buf)
            buf.seek(0)
            data_bytes = buf.getvalue()
            wb.close()

            fecha_tag = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"CONSOLIDADO_SALA_SITUACIONAL_{fecha_tag}.xlsx"
            self.send_response(200)
            self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            self.send_header('Content-Disposition', f'attachment; filename="{filename}"')
            self.send_header('Content-Length', str(len(data_bytes)))
            self.send_header('Access-Control-Allow-Origin', '*')
            self.send_header('Cache-Control', 'no-cache, no-store, must-revalidate')
            self.end_headers()
            self.wfile.write(data_bytes)
            return

        # 11. Exportar Casos para Recategorizar (.xlsx)
        elif path == "/api/export_recategorizacion":
            out_file = os.path.join(WORKSPACE_DIR, "INFORME_CASOS_PARA_RECATEGORIZAR.xlsx")
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Casos para Recategorizar"
            ws.views.sheetView[0].showGridLines = True

            cur.execute("""
            SELECT c.*, o.nombre as operador_nombre
            FROM casos c
            JOIN operadores o ON c.operador_id = o.id
            WHERE c.estatus_llamada = 'RECATEGORIZAR' OR c.inconsistencia_flag = 1
            ORDER BY c.municipio, c.parroquia
            """)
            rows = cur.fetchall()

            headers = [
                "CODIGO CASO", "OPERADOR QUE AUDITO", "CEDULA", "DENUNCIANTE", "TELEFONO",
                "MUNICIPIO", "PARROQUIA", "DIRECCION", "POSTE", "CATEGORIA ORIGINAL",
                "ESTATUS MARCADO", "¿PERSISTE?", "¿REQUIERE CAMPO?", "OBSERVACIONES / TESTIMONIO",
                "MOTIVO INCONSISTENCIA DETECTADA"
            ]
            ws.append(headers)

            font_head = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
            fill_head = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
            for c_i in range(1, len(headers) + 1):
                c = ws.cell(row=1, column=c_i)
                c.font = font_head
                c.fill = fill_head
                c.alignment = Alignment(horizontal="center", vertical="center")

            for r in rows:
                ws.append([
                    r["codigo_caso"], r["operador_nombre"], r["cedula"], r["denunciante"], r["telefono_1"],
                    r["municipio"], r["parroquia"], r["direccion"], r["poste_confirmado"] or r["poste_detectado"],
                    r["categoria_homologada"], r["estatus_llamada"], r["persiste"], r["requiere_campo"],
                    r["observaciones"], r["inconsistencia_motivo"]
                ])

            for col in ws.columns:
                max_len = max(len(str(cell.value or '')) for cell in col)
                col_letter = get_column_letter(col[0].column)
                ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

            import io
            buf = io.BytesIO()
            wb.save(buf)
            buf.seek(0)
            data_bytes = buf.getvalue()
            wb.close()
            conn.close()

            fecha_tag = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"CASOS_PARA_RECATEGORIZAR_{fecha_tag}.xlsx"
            self.send_response(200)
            self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            self.send_header('Content-Disposition', f'attachment; filename="{filename}"')
            self.send_header('Content-Length', str(len(data_bytes)))
            self.send_header('Access-Control-Allow-Origin', '*')
            self.send_header('Cache-Control', 'no-cache, no-store, must-revalidate')
            self.end_headers()
            self.wfile.write(data_bytes)
            return

        # 12. Listar Casos Desestimados para Supervisor
        elif path == "/api/supervisor/desestimados":
            cur.execute("""
            SELECT c.*, o.nombre as operador_nombre
            FROM casos c
            JOIN operadores o ON c.operador_id = o.id
            WHERE (c.estatus_llamada = 'DESESTIMAR' OR c.persiste = 'DESESTIMAR')
              AND c.devuelto_flag = 0
              AND (c.desestimado_aprobado = 0 OR c.desestimado_aprobado IS NULL)
            ORDER BY c.actualizado_en DESC, c.id DESC
            """)
            rows = [dict(r) for r in cur.fetchall()]
            conn.close()
            self.send_json(rows)

        # 13. Exportar Casos Desestimados a Excel (.xlsx)
        elif path == "/api/export_desestimados":
            out_file = os.path.join(WORKSPACE_DIR, "INFORME_CASOS_DESESTIMADOS.xlsx")
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Casos Desestimados"
            ws.views.sheetView[0].showGridLines = True

            cur.execute("""
            SELECT c.*, o.nombre as operador_nombre
            FROM casos c
            JOIN operadores o ON c.operador_id = o.id
            WHERE c.estatus_llamada = 'DESESTIMAR' OR c.persiste = 'DESESTIMAR'
            ORDER BY c.actualizado_en DESC, c.municipio, c.parroquia
            """)
            rows = cur.fetchall()

            headers = [
                "CODIGO CASO", "OPERADOR QUE DESESTIMO", "FECHA CONTACTO", "CEDULA", "DENUNCIANTE",
                "TELEFONO 1", "TELEFONO 2", "MUNICIPIO", "PARROQUIA", "DIRECCION",
                "POSTE", "CATEGORIA ORIGINAL", "SUBCATEGORIA", "DESCRIPCION ORIGINAL",
                "ESTATUS", "¿PERSISTE?", "OBSERVACIONES / MOTIVO DESESTIMACION", "FECHA ACTUALIZADO"
            ]
            ws.append(headers)

            font_head = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
            fill_head = PatternFill(start_color="991B1B", end_color="991B1B", fill_type="solid")
            for c_i in range(1, len(headers) + 1):
                c = ws.cell(row=1, column=c_i)
                c.font = font_head
                c.fill = fill_head
                c.alignment = Alignment(horizontal="center", vertical="center")

            for r in rows:
                ws.append([
                    r["codigo_caso"], r["operador_nombre"], r["fecha_contacto"], r["cedula"], r["denunciante"],
                    r["telefono_1"], r["telefono_2"], r["municipio"], r["parroquia"], r["direccion"],
                    r["poste_confirmado"] or r["poste_detectado"], r["categoria_homologada"], r["subcategoria"],
                    r["descripcion_reporte"], r["estatus_llamada"], r["persiste"], r["observaciones"],
                    r["actualizado_en"]
                ])

            for col in ws.columns:
                max_len = max(len(str(cell.value or '')) for cell in col)
                col_letter = get_column_letter(col[0].column)
                ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

            import io
            buf = io.BytesIO()
            wb.save(buf)
            buf.seek(0)
            data_bytes = buf.getvalue()
            wb.close()
            conn.close()

            fecha_tag = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"CASOS_PARA_RECATEGORIZAR_{fecha_tag}.xlsx"
            self.send_response(200)
            self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            self.send_header('Content-Disposition', f'attachment; filename="{filename}"')
            self.send_header('Content-Length', str(len(data_bytes)))
            self.send_header('Access-Control-Allow-Origin', '*')
            self.send_header('Cache-Control', 'no-cache, no-store, must-revalidate')
            self.end_headers()
            self.wfile.write(data_bytes)
            return

        else:
            conn.close()
            self.send_json({"error": "Endpoint no encontrado"}, 404)

    def handle_api_post(self, path, data):
        conn = get_db()
        cur = conn.cursor()

        # Guardar llamada
        if path == "/api/save_call":
            op_id = data.get("operador_id")
            cases_updates = data.get("cases", [])
            now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            if not op_id or not cases_updates:
                conn.close()
                self.send_json({"error": "Datos incompletos"}, 400)
                return

            cur.execute("UPDATE operadores SET ultima_actividad = ? WHERE id = ?", (now_str, op_id))
            op_row = cur.execute("SELECT nombre FROM operadores WHERE id = ?", (op_id,)).fetchone()
            op_nombre_real = op_row[0] if op_row else f"OP_{op_id}"
            warnings_total = []

            for cu in cases_updates:
                c_id = cu.get("id")
                est = cu.get("estatus_llamada")
                persiste = cu.get("persiste")
                campo = cu.get("requiere_campo")
                obs = cu.get("observaciones", "")
                poste_conf = cu.get("poste_confirmado", "")
                cat_hom = cu.get("categoria_homologada", "")

                warns = audit_case_coherence(cat_hom, est, persiste, campo, obs)
                incons_flag = 1 if len(warns) > 0 else 0
                incons_motivo = "; ".join(warns) if warns else None

                consecutivo_val = int(cu.get("consecutivo_linea", 0))
                cur.execute("""
                UPDATE casos SET 
                    estatus_llamada = ?,
                    persiste = ?,
                    requiere_campo = ?,
                    fecha_contacto = ?,
                    observaciones = ?,
                    poste_confirmado = ?,
                    estado_auditoria = 'GESTIONADO',
                    inconsistencia_flag = ?,
                    inconsistencia_motivo = ?,
                    devuelto_flag = 0,
                    consecutivo_linea = ?,
                    operador_nombre = ?,
                    actualizado_en = ?
                WHERE id = ? AND operador_id = ?
                """, (est, persiste, campo, now_str[:10], obs, poste_conf, incons_flag, incons_motivo, consecutivo_val, op_nombre_real, now_str, c_id, op_id))

                if warns:
                    warnings_total.extend(warns)

            first_case_code = cases_updates[0].get("codigo_caso", "CASO")
            cur.execute("""
            INSERT INTO logs (operador_id, tipo_accion, detalle)
            VALUES (?, 'REGISTRO_LLAMADA', ?)
            """, (op_id, f"Gestionó {len(cases_updates)} caso(s) [{first_case_code} - {cases_updates[0].get('estatus_llamada')}]"))

            conn.commit()
            conn.close()
            self.send_json({"success": True, "updated": len(cases_updates), "warnings": warnings_total})

        # Devolver caso
        elif path == "/api/supervisor/devolver":
            caso_id = data.get("caso_id")
            motivo = data.get("motivo", "Verificar datos de llamada")
            cur.execute("""
            UPDATE casos SET 
                devuelto_flag = 1,
                devuelto_motivo = ?,
                inconsistencia_flag = 0,
                inconsistencia_aprobada = 0,
                estado_auditoria = 'DEVUELTO_CORRECCION'
            WHERE id = ?
            """, (motivo, caso_id))
            cur.execute("INSERT INTO logs (operador_id, tipo_accion, detalle) VALUES (0, 'CASO_DEVUELTO', ?)",
                        (f"Supervisor devolvió caso #{caso_id} por: {motivo}",))
            conn.commit()
            conn.close()
            self.send_json({"success": True})

        # Modificar nombre / estatus del operador
        elif path == "/api/supervisor/update_operator":
            op_id = data.get("operador_id")
            nombre = (data.get("nombre") or "").strip().upper()
            activo = 1 if data.get("activo", True) else 0

            if not op_id or not nombre:
                conn.close()
                self.send_json({"error": "Faltan datos obligatorios (operador_id, nombre)"}, 400)
                return

            pestaña = f"OP{str(op_id).zfill(2)} - {nombre}"

            cur.execute("""
            UPDATE operadores 
            SET nombre = ?, pestaña = ?, activo = ?
            WHERE id = ?
            """, (nombre, pestaña, activo, op_id))

            cur.execute("""
            INSERT INTO logs (operador_id, tipo_accion, detalle)
            VALUES (?, 'OPERADOR_MODIFICADO', ?)
            """, (op_id, f"Supervisor actualizó nombre del Operador #{op_id} a '{nombre}' (Activo: {activo})"))

            conn.commit()
            conn.close()
            self.send_json({"success": True, "operador_id": op_id, "nombre": nombre, "pestaña": pestaña, "activo": activo})

        # Limpiar todas las llamadas de prueba (DESHABILITADO POR SEGURIDAD)
        elif path == "/api/supervisor/reset_test_data":
            conn.close()
            self.send_json({"error": "Función de reseteo deshabilitada por seguridad de producción"}, 403)
            return
            self.send_json({"success": True, "message": "Base de datos restaurada a cero. Todos los casos listos en PENDIENTE."})

        # Reabrir caso desestimado (regresar al operador con aviso explícito de devolución)
        elif path == "/api/supervisor/reabrir_caso":
            caso_id = data.get("caso_id")
            motivo = (data.get("motivo") or "").strip()
            if not motivo:
                motivo = "Caso desestimado devuelto por Supervisor: Verificar y contactar nuevamente al denunciante"

            if not caso_id:
                conn.close()
                self.send_json({"error": "Falta caso_id"}, 400)
                return

            dev_texto = f"CASO DESESTIMADO DEVUELTO: {motivo}"

            cur.execute("""
            UPDATE casos SET 
                estatus_llamada = NULL,
                persiste = NULL,
                requiere_campo = NULL,
                estado_auditoria = 'DEVUELTO_CORRECCION',
                devuelto_flag = 1,
                devuelto_motivo = ?,
                inconsistencia_flag = 0,
                inconsistencia_motivo = NULL,
                observaciones = CASE 
                    WHEN observaciones IS NOT NULL AND observaciones != '' 
                    THEN observaciones || ' [Regresado por Supervisor: ' || ? || ']'
                    ELSE '[Regresado por Supervisor: ' || ? || ']'
                END,
                actualizado_en = NULL
            WHERE id = ?
            """, (dev_texto, motivo, motivo, caso_id))

            cur.execute("""
            INSERT INTO logs (operador_id, tipo_accion, detalle)
            VALUES (0, 'CASO_DEVUELTO', ?)
            """, (f"Supervisor regresó caso desestimado #{caso_id} al operario por: {motivo}",))

            conn.commit()
            conn.close()
            self.send_json({"success": True, "message": "Caso regresado al operario con aviso explícito de devolución."})

        # Aprobar / Ratificar caso desestimado
        elif path == "/api/supervisor/aprobar_desestimado":
            caso_id = data.get("caso_id")
            if not caso_id:
                conn.close()
                self.send_json({"error": "Falta caso_id"}, 400)
                return

            cur.execute("""
            UPDATE casos SET 
                desestimado_aprobado = 1,
                devuelto_flag = 0,
                estado_auditoria = 'GESTIONADO'
            WHERE id = ?
            """, (caso_id,))

            cur.execute("""
            INSERT INTO logs (operador_id, tipo_accion, detalle)
            VALUES (0, 'DESESTIMACION_APROBADA', ?)
            """, (f"Supervisor ratificó y aprobó desestimación del caso #{caso_id}",))

            conn.commit()
            conn.close()
            self.send_json({"success": True, "message": "Caso desestimado aprobado formalmente."})

        # Aprobar caso inconsistente
        elif path == "/api/supervisor/aprobar":
            caso_id = data.get("caso_id")
            cur.execute("""
            UPDATE casos SET 
                inconsistencia_flag = 0, 
                devuelto_flag = 0, 
                inconsistencia_aprobada = 1,
                inconsistencia_motivo = NULL,
                estado_auditoria = 'GESTIONADO'
            WHERE id = ?
            """, (caso_id,))
            conn.commit()
            conn.close()
            self.send_json({"success": True, "message": "Caso aprobado formalmente."})

        # Reasignación básica manual
        elif path == "/api/supervisor/reasignar":
            from_op = int(data.get("from_operador_id", 0))
            to_op = int(data.get("to_operador_id", 0))
            cant_clientes = int(data.get("cantidad_clientes", 10))

            cur.execute("""
            SELECT DISTINCT id_llamada
            FROM casos
            WHERE operador_id = ? AND estado_auditoria = 'PENDIENTE'
            LIMIT ?
            """, (from_op, cant_clientes))
            clientes_to_move = [r[0] for r in cur.fetchall() if r[0]]

            casos_movidos = 0
            if clientes_to_move:
                cur.execute("SELECT nombre FROM operadores WHERE id = ?", (to_op,))
                to_row = cur.fetchone()
                to_name = to_row[0] if to_row else f"OPERADOR {to_op}"

                placeholders = ",".join("?" for _ in clientes_to_move)
                cur.execute(f"SELECT COUNT(*) FROM casos WHERE id_llamada IN ({placeholders}) AND estado_auditoria = 'PENDIENTE'", clientes_to_move)
                casos_movidos = cur.fetchone()[0]

                cur.execute(f"UPDATE clientes SET operador_id = ? WHERE id_llamada IN ({placeholders})", [to_op] + clientes_to_move)
                cur.execute(f"UPDATE casos SET operador_id = ?, operador_nombre = ? WHERE id_llamada IN ({placeholders}) AND estado_auditoria = 'PENDIENTE'", [to_op, to_name] + clientes_to_move)
                cur.execute("INSERT INTO logs (operador_id, tipo_accion, detalle) VALUES (0, 'REASIGNACION', ?)",
                            (f"Supervisor reasignó {len(clientes_to_move)} clientes ({casos_movidos} casos) del Op #{from_op} al Op #{to_op} ({to_name})",))
                conn.commit()

            conn.close()
            self.send_json({
                "status": "ok",
                "success": True,
                "clientes_movidos": len(clientes_to_move),
                "reassigned_clients": len(clientes_to_move),
                "casos_movidos": casos_movidos
            })

        # MOTOR DE INGESTA, CRUCE Y ASIGNACIÓN INTELIGENTE CON MÚLTIPLES CRITERIOS
        elif path == "/api/supervisor/cross_and_assign":
            file_base64 = data.get("file_base64", "")
            file_path = data.get("file_path", "")
            active_ops = [int(x) for x in data.get("active_operator_ids", [])]
            criteria = data.get("criteria", "EQUILIBRADO") # EQUILIBRADO, POR_MUNICIPIO, POR_CATEGORIA_TX
            
            raw_bytes = None
            if file_base64:
                import base64
                if "," in file_base64:
                    file_base64 = file_base64.split(",", 1)[1]
                try:
                    raw_bytes = base64.b64decode(file_base64)
                except Exception as e:
                    conn.close()
                    self.send_json({"error": f"Error decodificando archivo Excel: {str(e)}"}, 400)
                    return
            elif file_path and os.path.exists(file_path):
                try:
                    with open(file_path, "rb") as f:
                        raw_bytes = f.read()
                except Exception:
                    pass

            if not raw_bytes:
                conn.close()
                self.send_json({"error": "No se recibió ningún archivo Excel (.xlsx) para cruzar."}, 400)
                return

            if not active_ops:
                conn.close()
                self.send_json({"error": "Debe seleccionar al menos un operador activo."}, 400)
                return

            try:
                import io
                wb_new = openpyxl.load_workbook(io.BytesIO(raw_bytes), read_only=True, data_only=True)
                # Seleccionar primera hoja o 'Hoja1' o 'Casos Pendientes'
                sname = wb_new.sheetnames[0]
                ws = wb_new[sname]
                it = ws.iter_rows(values_only=True)
                raw_headers = [str(c).strip().upper() if c else "" for c in next(it)]

                # Mapear columnas dinámicamente
                def get_col_idx(names):
                    for n in names:
                        if n in raw_headers: return raw_headers.index(n)
                    return -1

                idx_code = get_col_idx(["CODIGO CASO", "CODIGO_CASO", "CASO", "NRO CASO"])
                idx_ced = get_col_idx(["CEDULA", "CEDULA_IDENTIDAD", "CI"])
                idx_den = get_col_idx(["DENUNCIANTE", "NOMBRE_DENUNCIANTE", "SOLICITANTE"])
                idx_t1 = get_col_idx(["TELEFONO", "TELEFONO_1", "CELULAR"])
                idx_t2 = get_col_idx(["TELEFONO ALT", "TELEFONO_2", "TELEFONO_ALT"])
                idx_mun = get_col_idx(["MUNICIPIO"])
                idx_parr = get_col_idx(["PARROQ", "PARROQUIA"])
                idx_sec = get_col_idx(["COMUNIDAD", "SECTOR"])
                idx_dir = get_col_idx(["DIRECCION"])
                idx_desc = get_col_idx(["DESCRIP", "DESCRIPCION", "DESCRIPCION_REPORTE"])
                idx_cat = get_col_idx(["CATEGORIA HOMOLOGADA", "CATEGORIA_HOMOLOGADA", "CATEGORIA"])
                idx_sub = get_col_idx(["SUBCATEG", "SUBCATEGORIA"])
                idx_fgen = get_col_idx(["FECHA GENERADO", "FECHA_GENERADO", "FECHA"])

                if idx_code == -1:
                    wb_new.close()
                    conn.close()
                    self.send_json({"error": "No se encontró la columna de Código de Caso en el archivo."}, 400)
                    return

                # Obtener todos los códigos ya en DB y su estatus
                cur.execute("SELECT codigo_caso, estatus_llamada FROM casos")
                db_cases_map = {r[0]: (r[1] or '') for r in cur.fetchall()}

                # Casos pendientes no atendidos en el nuevo archivo
                cases_to_assign = []
                nuevos_count = 0
                no_atendidos_previos = 0

                poste_regex = re.compile(r'poste\s*(?:n[o°º\.]*|num(?:ero)?\.?)?\s*([a-zA-Z0-9\-_/]+)', re.IGNORECASE)

                for r in it:
                    code_val = str(r[idx_code]).strip() if r[idx_code] else ""
                    if not code_val or code_val in ["None", ""]:
                        continue

                    prev_estatus = db_cases_map.get(code_val, None)

                    # Si ya estaba en DB y fue ATENDIDO o SOLVENTADO, no se reasigna
                    if prev_estatus in ["ATENDIDO", "ATENDIDO POR TERCEROS", "DESESTIMAR"]:
                        continue

                    is_nuevo = (prev_estatus is None)
                    if is_nuevo: nuevos_count += 1
                    else: no_atendidos_previos += 1

                    mun_val = str(r[idx_mun]).strip().upper() if idx_mun != -1 and r[idx_mun] else ""
                    parr_val = str(r[idx_parr]).strip().upper() if idx_parr != -1 and r[idx_parr] else ""
                    ced_val = str(r[idx_ced]).strip().upper() if idx_ced != -1 and r[idx_ced] else ""
                    den_val = str(r[idx_den]).strip().upper() if idx_den != -1 and r[idx_den] else ""
                    t1_val = format_phone(r[idx_t1]) if idx_t1 != -1 and r[idx_t1] else ""
                    t2_val = format_phone(r[idx_t2]) if idx_t2 != -1 and r[idx_t2] else ""
                    sec_val = str(r[idx_sec]).strip() if idx_sec != -1 and r[idx_sec] else ""
                    dir_val = str(r[idx_dir]).strip() if idx_dir != -1 and r[idx_dir] else ""
                    desc_val = str(r[idx_desc]).strip() if idx_desc != -1 and r[idx_desc] else ""
                    cat_val = str(r[idx_cat]).strip().upper().replace("TRAFO", "TX").replace("TRAFOS", "TX") if idx_cat != -1 and r[idx_cat] else ""
                    sub_val = str(r[idx_sub]).strip() if idx_sub != -1 and r[idx_sub] else ""
                    fgen_val = str(r[idx_fgen])[:10] if idx_fgen != -1 and r[idx_fgen] else ""

                    # Poste
                    m_p = poste_regex.search(f"{dir_val} {desc_val}")
                    poste_val = f"Poste {m_p.group(1)}" if m_p else ""

                    tipo_c = "PENDIENTE (PRIMER CONTACTO)" if is_nuevo else "REINTENTO (SIN COMUNICACION PREVIA)"

                    # Clave de cliente para mantener casos agrupados
                    if ced_val and ced_val not in ["V0", "0", "V", "SIN CEDULA"]:
                        ckey = ("CED", ced_val)
                    elif len(t1_val) >= 10:
                        ckey = ("TEL", t1_val)
                    elif den_val:
                        ckey = ("DEN", den_val)
                    else:
                        ckey = ("CASO", code_val)

                    cases_to_assign.append({
                        "ckey": ckey,
                        "codigo_caso": code_val,
                        "tipo_contacto": tipo_c,
                        "municipio": mun_val,
                        "parroquia": parr_val,
                        "cedula": ced_val,
                        "denunciante": den_val,
                        "telefono_1": t1_val,
                        "telefono_2": t2_val,
                        "sector": sec_val,
                        "direccion": dir_val,
                        "poste_detectado": poste_val,
                        "fecha_generado": fgen_val,
                        "categoria_homologada": cat_val,
                        "subcategoria": sub_val,
                        "descripcion_reporte": desc_val,
                        "is_tx": ("TX" in cat_val or "TRANSFORMADOR" in cat_val)
                    })

                wb_new.close()

                # Agrupar por denunciante
                groups = defaultdict(list)
                for c in cases_to_assign:
                    groups[c["ckey"]].append(c)

                # DISTRIBUCIÓN SEGÚN CRITERIO
                num_active = len(active_ops)
                op_buckets = {op_id: [] for op_id in active_ops}

                if criteria == "POR_CATEGORIA_TX":
                    # Mitad de operadores activos a TX, la otra mitad al resto
                    half = max(1, num_active // 2)
                    tx_ops = active_ops[:half]
                    other_ops = active_ops[half:] if len(active_ops) > 1 else active_ops

                    for ckey, items in groups.items():
                        has_tx = any(it["is_tx"] for it in items)
                        target_pool = tx_ops if has_tx else other_ops
                        # Asignar al que tenga menos
                        min_op = min(target_pool, key=lambda o: len(op_buckets[o]))
                        op_buckets[min_op].extend(items)

                elif criteria == "POR_MUNICIPIO":
                    # Reparto territorial
                    mun_groups = defaultdict(list)
                    for ckey, items in groups.items():
                        m = items[0]["municipio"]
                        mun_groups[m].append(items)

                    for m, m_items in mun_groups.items():
                        for items in m_items:
                            min_op = min(active_ops, key=lambda o: len(op_buckets[o]))
                            op_buckets[min_op].extend(items)
                else:
                    # EQUILIBRADO BALANCEADO (Por defecto)
                    # Reintentos primero, luego por volumen de casos
                    sorted_groups = sorted(groups.items(), key=lambda g: (-len(g[1]), g[1][0]["municipio"]))
                    for ckey, items in sorted_groups:
                        min_op = min(active_ops, key=lambda o: len(op_buckets[o]))
                        op_buckets[min_op].extend(items)

                # GUARDAR EN BASE DE DATOS
                now_s = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                total_asignados = 0

                for op_id, assigned_items in op_buckets.items():
                    # Agrupar por ckey para generar ID_LLAMADA
                    op_client_groups = defaultdict(list)
                    for itm in assigned_items:
                        op_client_groups[itm["ckey"]].append(itm)

                    for cli_idx, (ckey, itms) in enumerate(op_client_groups.items(), 1):
                        id_llamada_new = f"CALL-OP{op_id:02d}-{cli_idx:04d}"
                        first_it = itms[0]
                        tot_in_g = len(itms)

                        # Insertar cliente
                        cur.execute("""
                        INSERT OR REPLACE INTO clientes (id_llamada, operador_id, denunciante, cedula, telefono_1, telefono_2, municipio, parroquia, sector, direccion, tipo_contacto)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """, (id_llamada_new, op_id, first_it["denunciante"], first_it["cedula"], first_it["telefono_1"], first_it["telefono_2"], first_it["municipio"], first_it["parroquia"], first_it["sector"], first_it["direccion"], first_it["tipo_contacto"]))

                        for case_pos, itm in enumerate(itms, 1):
                            caso_nro_str = f"Caso {case_pos} de {tot_in_g}"
                            cur.execute("""
                            INSERT OR REPLACE INTO casos (
                                codigo_caso, id_llamada, operador_id, caso_nro, tipo_contacto,
                                municipio, parroquia, cedula, denunciante, telefono_1, telefono_2,
                                sector, direccion, poste_detectado, fecha_generado, categoria_homologada,
                                subcategoria, descripcion_reporte, estatus_llamada, persiste, requiere_campo,
                                estado_auditoria, actualizado_en
                            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, NULL, NULL, NULL, 'PENDIENTE', ?)
                            """, (
                                itm["codigo_caso"], id_llamada_new, op_id, caso_nro_str, itm["tipo_contacto"],
                                itm["municipio"], itm["parroquia"], itm["cedula"], itm["denunciante"], itm["telefono_1"], itm["telefono_2"],
                                itm["sector"], itm["direccion"], itm["poste_detectado"], itm["fecha_generado"], itm["categoria_homologada"],
                                itm["subcategoria"], itm["descripcion_reporte"], now_s
                            ))
                            total_asignados += 1

                # Actualizar estado activo de operadores
                cur.execute("UPDATE operadores SET activo = 0")
                for op_id in active_ops:
                    cur.execute("UPDATE operadores SET activo = 1 WHERE id = ?", (op_id,))

                cur.execute("""
                INSERT INTO logs (operador_id, tipo_accion, detalle)
                VALUES (0, 'CRUCE_ASIGNACION', ?)
                """, (f"Supervisor cruzó archivo {os.path.basename(file_path)} y asignó {total_asignados} casos entre {num_active} operadores activos.",))

                conn.commit()
                conn.close()

                resumen_ops = {op_id: len(op_buckets[op_id]) for op_id in active_ops}
                self.send_json({
                    "success": True,
                    "nuevos_detectados": nuevos_count,
                    "reintentos_no_atendidos": no_atendidos_previos,
                    "total_asignados": total_asignados,
                    "operadores_activos": num_active,
                    "criterio": criteria,
                    "distribucion": resumen_ops
                })

            except Exception as e:
                conn.close()
                self.send_json({"error": f"Error procesando archivo: {str(e)}"}, 500)

        # Exportar Excel general
        elif path == "/api/export_excel":
            out_file = os.path.join(WORKSPACE_DIR, "CONSOLIDADO_SALA_SITUACIONAL_EN_VIVO.xlsx")
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Casos y Auditoría"
            ws.views.sheetView[0].showGridLines = True

            cur.execute("""
            SELECT c.*, o.nombre as operador_nombre
            FROM casos c
            JOIN operadores o ON c.operador_id = o.id
            ORDER BY c.operador_id, c.id_llamada
            """)
            rows = cur.fetchall()

            headers = [
                "OPERADOR", "ESTADO_AUDITORIA", "ID_LLAMADA", "CASO_NRO", "MUNICIPIO", "PARROQUIA",
                "CEDULA", "DENUNCIANTE", "TELEFONO_1", "TELEFONO_2", "SECTOR", "DIRECCION",
                "POSTE_DETECTADO", "CODIGO_CASO", "FECHA_GENERADO", "CATEGORIA_HOMOLOGADA",
                "SUBCATEGORIA", "DESCRIPCION_REPORTE", "ESTATUS_LLAMADA", "¿ESTE_CASO_PERSISTE?",
                "REQUIERE_CAMPO", "FECHA_CONTACTO", "OBSERVACIONES", "POSTE_CONFIRMADO",
                "INCONSISTENCIA_ALERTA", "MOTIVO_INCONSISTENCIA"
            ]
            ws.append(headers)

            font_head = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
            fill_head = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
            for c_i in range(1, len(headers) + 1):
                c = ws.cell(row=1, column=c_i)
                c.font = font_head
                c.fill = fill_head
                c.alignment = Alignment(horizontal="center", vertical="center")

            for r in rows:
                ws.append([
                    r["operador_nombre"], r["estado_auditoria"], r["id_llamada"], r["caso_nro"],
                    r["municipio"], r["parroquia"], r["cedula"], r["denunciante"], r["telefono_1"],
                    r["telefono_2"], r["sector"], r["direccion"], r["poste_detectado"], r["codigo_caso"],
                    r["fecha_generado"], r["categoria_homologada"], r["subcategoria"], r["descripcion_reporte"],
                    r["estatus_llamada"], r["persiste"], r["requiere_campo"], r["fecha_contacto"],
                    r["observaciones"], r["poste_confirmado"], "SÍ" if r["inconsistencia_flag"] else "NO",
                    r["inconsistencia_motivo"]
                ])

            for col in ws.columns:
                max_len = max(len(str(cell.value or '')) for cell in col)
                col_letter = get_column_letter(col[0].column)
                ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

            wb.save(out_file)
            wb.close()
            conn.close()
            self.send_json({"success": True, "file": out_file})

        else:
            conn.close()
            self.send_json({"error": "Endpoint POST no encontrado"}, 404)

if __name__ == "__main__":
    # Obtener IP local de la máquina automáticamente
    local_ips = []
    try:
        hostname = socket.gethostname()
        _, _, ips = socket.gethostbyname_ex(hostname)
        local_ips = [ip for ip in ips if not ip.startswith('127.')]
    except Exception:
        pass

    ip_display = local_ips[0] if local_ips else "IP_DE_ESTA_PC"

    socketserver.TCPServer.allow_reuse_address = True
    with socketserver.ThreadingTCPServer(("", PORT), SituationalRoomHandler) as httpd:
        print(f"===========================================================")
        print(f"  SALA SITUACIONAL CORPOELEC / 1x10 DEL BUEN GOBIERNO")
        print(f"  SERVIDOR ACTIVO EN LA RED CORPORATIVA Y WIFI")
        print(f"===========================================================")
        print(f"  DIRECCIÓN IP DE ESTA MÁQUINA SERVIDORA: {ip_display}")
        print(f"")
        print(f"  1. PARA LAS 15 MÁQUINAS Y LAPTOPS DE LOS OPERADORES:")
        print(f"     Abre en su navegador (Chrome/Edge):")
        print(f"     --> http://{ip_display}:{PORT}/operador")
        print(f"")
        print(f"  2. PARA EL MONITOR / PANTALLA GIGANTE:")
        print(f"     --> http://{ip_display}:{PORT}/monitor")
        print(f"")
        print(f"  3. PARA TU PANTALLA DE SUPERVISOR:")
        print(f"     --> http://localhost:{PORT}/supervisor")
        print(f"===========================================================")
        print(f"  Presiona Ctrl + C en esta ventana para apagar el servidor.")
        print(f"===========================================================")
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            print("\nServidor detenido por el usuario.")

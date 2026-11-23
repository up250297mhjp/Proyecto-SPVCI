import os
import mysql.connector
from dotenv import load_dotenv

load_dotenv()


class SPVCIRepository:
    def __init__(self):
        self.host = os.getenv("DB_HOST", "localhost")
        self.port = int(os.getenv("DB_PORT", 3306))
        self.user = os.getenv("DB_USER", "root")
        self.password = os.getenv("DB_PASSWORD", "")
        self.database = os.getenv("DB_NAME", "spvci")
        self._asegurar_columnas_esquema()

    def get_connection(self):
        return mysql.connector.connect(
            host=self.host,
            port=self.port,
            user=self.user,
            password=self.password,
            database=self.database
        )

    def _asegurar_columnas_esquema(self):
        """Migra columnas opcionales sin recrear tablas ni alterar datos existentes."""
        conn = None
        cursor = None
        try:
            conn = self.get_connection()
            cursor = conn.cursor()
            migraciones = {
                "Producto": [
                    ("stockMaximo", "INT NOT NULL DEFAULT 20 COMMENT 'Maximo de unidades en inventario' AFTER stockMinimo",
                     "UPDATE Producto SET stockMaximo = GREATEST(stockMinimo * 2, 10)"),
                ],
                "NotaRecibo": [
                    ("estado", "VARCHAR(20) NOT NULL DEFAULT 'Activa' COMMENT 'Estado de la nota: Activa o Cancelada'", None),
                ],
                "Proveedor": [
                    ("estado", "BOOLEAN NOT NULL DEFAULT TRUE COMMENT 'TRUE activo, FALSE baja logica'", None),
                ],
                "Cliente": [
                    ("estado", "BOOLEAN NOT NULL DEFAULT TRUE COMMENT 'TRUE activo, FALSE baja logica'", None),
                    ("numeroCalle", "VARCHAR(30) NULL COMMENT 'Numero exterior/interior' AFTER domicilio", None),
                    ("colonia", "VARCHAR(100) NULL COMMENT 'Colonia' AFTER numeroCalle", None),
                    ("ciudad", "VARCHAR(100) NULL COMMENT 'Ciudad' AFTER colonia", None),
                    ("codigoPostal", "VARCHAR(10) NULL COMMENT 'Codigo Postal' AFTER ciudad", None),
                ],
                "Ticket": [
                    ("estado", "BOOLEAN NOT NULL DEFAULT TRUE COMMENT 'TRUE activo/vigente, FALSE cancelado (baja logica)'", None),
                    ("motivoCancelacion", "VARCHAR(255) NULL COMMENT 'Motivo de cancelacion logica' AFTER estado", None),
                ],
            }
            for tabla, columnas in migraciones.items():
                for nombre, definicion, backfill in columnas:
                    cursor.execute(f"SHOW COLUMNS FROM `{tabla}` LIKE %s", (nombre,))
                    if not cursor.fetchone():
                        cursor.execute(f"ALTER TABLE `{tabla}` ADD COLUMN `{nombre}` {definicion}")
                        if backfill:
                            cursor.execute(backfill)
            conn.commit()
        except Exception:
            if conn and conn.is_connected():
                conn.rollback()
            raise
        finally:
            if conn and conn.is_connected():
                if cursor:
                    cursor.close()
                conn.close()

    # =========================================================================
    # MÓDULO 1: PRODUCTOS
    # =========================================================================
    def buscar_productos(self, criterio: str = "", categoria: str = "Todas",
                         solo_stock_bajo: bool = False, incluir_inactivos: bool = False) -> list[dict]:
        condiciones = ["1=1"]
        params = []

        if not incluir_inactivos:
            condiciones.append("p.estado = 1")

        if criterio.strip():
            condiciones.append("(p.nombre LIKE %s OR p.codigo LIKE %s OR p.categoria LIKE %s)")
            patron = f"%{criterio.strip()}%"
            params.extend([patron, patron, patron])

        if categoria and categoria != "Todas":
            condiciones.append("p.categoria = %s")
            params.append(categoria)

        if solo_stock_bajo:
            condiciones.append("p.existencia <= p.stockMinimo")

        where_sql = " AND ".join(condiciones)
        query = f"""
            SELECT 
                p.codigo,
                p.nombre,
                p.categoria,
                p.unidadMedida,
                p.existencia,
                p.stockMinimo,
                COALESCE(p.stockMaximo, p.stockMinimo * 2) AS stockMaximo,
                p.costo,
                p.precio,
                p.estado,
                (SELECT COUNT(*) FROM DetalleTicket dt WHERE dt.codigoProducto = p.codigo) AS ventas_registradas
            FROM Producto p
            WHERE {where_sql}
            ORDER BY p.codigo ASC
        """
        conn = None
        try:
            conn = self.get_connection()
            cursor = conn.cursor(dictionary=True)
            cursor.execute(query, tuple(params))
            rows = cursor.fetchall()
            for r in rows:
                r["costo"] = float(r["costo"]) if r["costo"] is not None else 0.0
                r["precio"] = float(r["precio"]) if r["precio"] is not None else 0.0
                r["estado"] = bool(r["estado"])
            return rows
        finally:
            if conn and conn.is_connected():
                cursor.close()
                conn.close()

    def obtener_producto(self, codigo: str) -> dict | None:
        productos = self.buscar_productos(criterio=codigo, incluir_inactivos=True)
        for p in productos:
            if p["codigo"] == codigo:
                return p
        return None

    def obtener_movimientos_producto(self, codigo: str) -> list[dict]:
        query = """
            SELECT 
                DATE_FORMAT(t.fecha, '%d/%m/%Y %H:%i') AS fecha_fmt,
                t.fecha AS fecha_orden,
                'Venta' AS tipo,
                CONCAT('V-', LPAD(t.folio, 6, '0')) AS documento,
                -CAST(dt.cantidad AS SIGNED) AS cantidad
            FROM DetalleTicket dt
            INNER JOIN Ticket t ON dt.folioTicket = t.folio
            WHERE dt.codigoProducto = %s

            UNION ALL

            SELECT 
                DATE_FORMAT(nr.fecha, '%d/%m/%Y') AS fecha_fmt,
                CAST(nr.fecha AS DATETIME) AS fecha_orden,
                CONCAT('Nota de recibo (', COALESCE(nr.estado, 'Activa'), ')') AS tipo,
                CONCAT('NR-', LPAD(nr.folio, 4, '0')) AS documento,
                CAST(dnr.cantidad AS SIGNED) AS cantidad
            FROM DetalleNotaRecibo dnr
            INNER JOIN NotaRecibo nr ON dnr.folioNota = nr.folio
            WHERE dnr.codigoProducto = %s

            ORDER BY fecha_orden DESC
            LIMIT 15
        """
        conn = None
        try:
            conn = self.get_connection()
            cursor = conn.cursor(dictionary=True)
            cursor.execute(query, (codigo, codigo))
            return cursor.fetchall()
        finally:
            if conn and conn.is_connected():
                cursor.close()
                conn.close()

    def agregar_producto(self, datos: dict):
        if self.obtener_producto(datos["codigo"]) is not None:
            raise ValueError(f"El código '{datos['codigo']}' ya está registrado.")
        if datos["stockMinimo"] <= 0:
            raise ValueError("El punto de reorden debe ser un número positivo (> 0).")
        if datos["costo"] <= 0 or datos["precio"] <= 0:
            raise ValueError("El costo y el precio deben ser números positivos (> 0).")

        query = """
            INSERT INTO Producto (
                codigo, nombre, categoria, unidadMedida,
                existencia, stockMinimo, stockMaximo, costo, precio, estado
            ) VALUES (%s, %s, %s, %s, 0, %s, %s, %s, %s, TRUE)
        """
        valores = (
            datos["codigo"],
            datos["nombre"],
            datos["categoria"],
            datos.get("unidadMedida", "pieza"),
            int(datos["stockMinimo"]),
            int(datos.get("stockMaximo", datos["stockMinimo"] * 2)),
            float(datos["costo"]),
            float(datos["precio"])
        )
        conn = None
        try:
            conn = self.get_connection()
            cursor = conn.cursor()
            cursor.execute(query, valores)
            conn.commit()
        finally:
            if conn and conn.is_connected():
                cursor.close()
                conn.close()

    def modificar_producto(self, codigo: str, datos: dict):
        if datos["stockMinimo"] <= 0:
            raise ValueError("El punto de reorden debe ser un número positivo (> 0).")
        if datos["costo"] <= 0 or datos["precio"] <= 0:
            raise ValueError("El costo y el precio deben ser mayores que 0.")
        if datos["existencia"] < 0:
            raise ValueError("La existencia debe ser mayor o igual a 0.")

        query = """
            UPDATE Producto
            SET categoria = %s,
                unidadMedida = %s,
                existencia = %s,
                stockMinimo = %s,
                stockMaximo = %s,
                costo = %s,
                precio = %s,
                estado = %s
            WHERE codigo = %s
        """
        valores = (
            datos["categoria"],
            datos["unidadMedida"],
            int(datos["existencia"]),
            int(datos["stockMinimo"]),
            int(datos.get("stockMaximo", datos["stockMinimo"] * 2)),
            float(datos["costo"]),
            float(datos["precio"]),
            bool(datos.get("estado", True)),
            codigo
        )
        conn = None
        try:
            conn = self.get_connection()
            cursor = conn.cursor()
            cursor.execute(query, valores)
            conn.commit()
        finally:
            if conn and conn.is_connected():
                cursor.close()
                conn.close()

    def eliminar_producto_logico(self, codigo: str):
        prod = self.obtener_producto(codigo)
        if not prod:
            raise ValueError("No se encontró el producto.")
        if prod["existencia"] != 0:
            raise ValueError(
                f"No se puede eliminar '{prod['nombre']}' porque su existencia actual es {prod['existencia']} "
                "(debe ser exactamente 0)."
            )

        query = "UPDATE Producto SET estado = FALSE WHERE codigo = %s"
        conn = None
        try:
            conn = self.get_connection()
            cursor = conn.cursor()
            cursor.execute(query, (codigo,))
            conn.commit()
        finally:
            if conn and conn.is_connected():
                cursor.close()
                conn.close()

    # =========================================================================
    # MÓDULO 2: PROVEEDORES (Con Eliminación Lógica y Validaciones)
    # =========================================================================
    def buscar_proveedores(self, criterio: str = "", incluir_inactivos: bool = False) -> list[dict]:
        condiciones = ["1=1"]
        params = []

        if not incluir_inactivos:
            condiciones.append("COALESCE(p.estado, 1) = 1")

        if criterio.strip():
            condiciones.append(
                "(p.empresa LIKE %s OR p.rfc LIKE %s OR p.nombre LIKE %s OR CAST(p.idProveedor AS CHAR) = %s)"
            )
            patron = f"%{criterio.strip()}%"
            params.extend([patron, patron, patron, criterio.strip()])

        where_sql = " AND ".join(condiciones)
        query = f"""
            SELECT 
                p.idProveedor,
                p.nombre,
                p.domicilio,
                p.telefono,
                COALESCE(p.correoE, '') AS correoE,
                COALESCE(p.paginaWeb, '') AS paginaWeb,
                p.empresa,
                p.rfc,
                COALESCE(p.estado, 1) AS estado,
                (SELECT COUNT(*) FROM NotaRecibo nr WHERE nr.idProveedor = p.idProveedor) AS notas_registradas
            FROM Proveedor p
            WHERE {where_sql}
            ORDER BY p.idProveedor ASC
        """
        conn = None
        try:
            conn = self.get_connection()
            cursor = conn.cursor(dictionary=True)
            cursor.execute(query, tuple(params))
            rows = cursor.fetchall()
            for r in rows:
                r["telefono"] = str(r.get("telefono", "") or "")
                r["estado"] = bool(r.get("estado", True))
            return rows
        finally:
            if conn and conn.is_connected():
                cursor.close()
                conn.close()

    def obtener_proveedor(self, id_proveedor: int) -> dict | None:
        query = """
            SELECT 
                p.idProveedor,
                p.nombre,
                p.domicilio,
                p.telefono,
                COALESCE(p.correoE, '') AS correoE,
                COALESCE(p.paginaWeb, '') AS paginaWeb,
                p.empresa,
                p.rfc,
                COALESCE(p.estado, 1) AS estado,
                (SELECT COUNT(*) FROM NotaRecibo nr WHERE nr.idProveedor = p.idProveedor) AS notas_registradas
            FROM Proveedor p
            WHERE p.idProveedor = %s
        """
        conn = None
        try:
            conn = self.get_connection()
            cursor = conn.cursor(dictionary=True)
            cursor.execute(query, (id_proveedor,))
            r = cursor.fetchone()
            if r:
                r["telefono"] = str(r.get("telefono", "") or "")
                r["estado"] = bool(r.get("estado", True))
            return r
        finally:
            if conn and conn.is_connected():
                cursor.close()
                conn.close()

    def _validar_datos_proveedor(self, datos: dict):
        rfc = str(datos.get("rfc", "")).strip().upper()
        if len(rfc) not in (12, 13):
            raise ValueError("El RFC debe contener 12 caracteres (persona moral) o 13 caracteres (persona física).")

        telefono = str(datos.get("telefono", "")).strip()
        if not telefono.isdigit() or len(telefono) < 7 or len(telefono) > 15:
            raise ValueError(
                "El número de teléfono solo debe contener dígitos numéricos (sin letras ni espacios, de 7 a 15 dígitos)."
            )

        correo = str(datos.get("correoE", "")).strip()
        if not correo or "@" not in correo:
            raise ValueError("El correo electrónico es inválido: debe incluir obligatoriamente el símbolo '@'.")

    def agregar_proveedor(self, datos: dict) -> int:
        self._validar_datos_proveedor(datos)
        rfc = str(datos["rfc"]).strip().upper()
        telefono_str = str(datos["telefono"]).strip()

        conn = None
        try:
            conn = self.get_connection()
            cursor = conn.cursor()
            cursor.execute("SELECT idProveedor FROM Proveedor WHERE rfc = %s", (rfc,))
            if cursor.fetchone():
                raise ValueError(f"El RFC '{rfc}' ya está registrado para otro proveedor.")

            query = """
                INSERT INTO Proveedor (nombre, domicilio, telefono, correoE, paginaWeb, empresa, rfc, estado)
                VALUES (%s, %s, %s, %s, %s, %s, %s, TRUE)
            """
            valores = (
                str(datos["nombre"]).strip(),
                str(datos["domicilio"]).strip(),
                telefono_str,  # Interpretado y guardado como string (VARCHAR)
                str(datos.get("correoE", "")).strip(),
                str(datos.get("paginaWeb", "")).strip() or None,
                str(datos["empresa"]).strip(),
                rfc
            )
            cursor.execute(query, valores)
            conn.commit()
            return cursor.lastrowid
        finally:
            if conn and conn.is_connected():
                cursor.close()
                conn.close()

    def modificar_proveedor(self, id_proveedor: int, datos: dict):
        self._validar_datos_proveedor(datos)
        rfc = str(datos["rfc"]).strip().upper()
        telefono_str = str(datos["telefono"]).strip()

        conn = None
        try:
            conn = self.get_connection()
            cursor = conn.cursor()
            cursor.execute(
                "SELECT idProveedor FROM Proveedor WHERE rfc = %s AND idProveedor != %s",
                (rfc, id_proveedor)
            )
            if cursor.fetchone():
                raise ValueError(f"El RFC '{rfc}' ya pertenece a otro proveedor.")

            query = """
                UPDATE Proveedor
                SET nombre = %s,
                    domicilio = %s,
                    telefono = %s,
                    correoE = %s,
                    paginaWeb = %s,
                    empresa = %s,
                    rfc = %s
                WHERE idProveedor = %s
            """
            valores = (
                str(datos["nombre"]).strip(),
                str(datos["domicilio"]).strip(),
                telefono_str,  # Interpretado y guardado como string (VARCHAR)
                str(datos.get("correoE", "")).strip(),
                str(datos.get("paginaWeb", "")).strip() or None,
                str(datos["empresa"]).strip(),
                rfc,
                id_proveedor
            )
            cursor.execute(query, valores)
            conn.commit()
        finally:
            if conn and conn.is_connected():
                cursor.close()
                conn.close()

    def eliminar_proveedor_logico(self, id_proveedor: int):
        """
        Eliminación lógica del proveedor (estado = FALSE).
        Evita conflictos de llaves foráneas y conserva intactas las notas de recibo registradas.
        """
        conn = None
        try:
            conn = self.get_connection()
            cursor = conn.cursor()
            cursor.execute("UPDATE Proveedor SET estado = FALSE WHERE idProveedor = %s", (id_proveedor,))
            conn.commit()
        finally:
            if conn and conn.is_connected():
                cursor.close()
                conn.close()

    def obtener_notas_por_proveedor(self, id_proveedor: int) -> list[dict]:
        query = """
            SELECT 
                nr.folio,
                DATE_FORMAT(nr.fecha, '%d/%m/%Y') AS fecha,
                COUNT(dnr.codigoProducto) AS partidas,
                COALESCE(SUM(dnr.cantidad), 0) AS unidades,
                nr.total AS costo_total,
                COALESCE(nr.estado, 'Activa') AS estado
            FROM NotaRecibo nr
            LEFT JOIN DetalleNotaRecibo dnr ON nr.folio = dnr.folioNota
            WHERE nr.idProveedor = %s
            GROUP BY nr.folio, nr.fecha, nr.total, nr.estado
            ORDER BY nr.folio DESC
        """
        conn = None
        try:
            conn = self.get_connection()
            cursor = conn.cursor(dictionary=True)
            cursor.execute(query, (id_proveedor,))
            rows = cursor.fetchall()
            for r in rows:
                r["costo_total"] = float(r["costo_total"] or 0.0)
                r["unidades"] = int(r["unidades"] or 0)
            return rows
        finally:
            if conn and conn.is_connected():
                cursor.close()
                conn.close()

    # =========================================================================
    # MÓDULO 3: NOTAS DE RECIBO
    # =========================================================================
    def obtener_siguiente_folio_nota(self) -> int:
        conn = None
        try:
            conn = self.get_connection()
            cursor = conn.cursor()
            cursor.execute("SELECT COALESCE(MAX(folio), 0) + 1 FROM NotaRecibo")
            r = cursor.fetchone()
            return int(r[0]) if r else 1
        finally:
            if conn and conn.is_connected():
                cursor.close()
                conn.close()

    def listar_notas_recibo(self, filtro: str = "") -> list[dict]:
        condicion = ""
        params = []
        if filtro.strip():
            condicion = "WHERE pr.empresa LIKE %s OR CAST(nr.folio AS CHAR) LIKE %s OR nr.estado LIKE %s"
            p = f"%{filtro.strip()}%"
            params = [p, p, p]

        query = f"""
            SELECT 
                nr.folio,
                DATE_FORMAT(nr.fecha, '%d/%m/%Y') AS fecha_fmt,
                nr.fecha AS fecha_raw,
                nr.idProveedor,
                pr.empresa AS proveedor,
                pr.rfc AS rfc_proveedor,
                COUNT(dnr.codigoProducto) AS partidas,
                COALESCE(SUM(dnr.cantidad), 0) AS unidades,
                nr.total,
                COALESCE(nr.estado, 'Activa') AS estado
            FROM NotaRecibo nr
            INNER JOIN Proveedor pr ON nr.idProveedor = pr.idProveedor
            LEFT JOIN DetalleNotaRecibo dnr ON nr.folio = dnr.folioNota
            {condicion}
            GROUP BY nr.folio, nr.fecha, nr.idProveedor, pr.empresa, pr.rfc, nr.total, nr.estado
            ORDER BY nr.folio DESC
        """
        conn = None
        try:
            conn = self.get_connection()
            cursor = conn.cursor(dictionary=True)
            cursor.execute(query, tuple(params))
            rows = cursor.fetchall()
            for r in rows:
                r["total"] = float(r["total"] or 0.0)
                r["unidades"] = int(r["unidades"] or 0)
            return rows
        finally:
            if conn and conn.is_connected():
                cursor.close()
                conn.close()

    def obtener_nota_completa(self, folio: int) -> dict | None:
        query_cab = """
            SELECT 
                nr.folio,
                DATE_FORMAT(nr.fecha, '%d/%m/%Y') AS fecha_fmt,
                nr.fecha AS fecha_raw,
                nr.idProveedor,
                pr.empresa AS proveedor,
                pr.rfc AS rfc_proveedor,
                pr.nombre AS contacto_proveedor,
                nr.total,
                COALESCE(nr.estado, 'Activa') AS estado
            FROM NotaRecibo nr
            INNER JOIN Proveedor pr ON nr.idProveedor = pr.idProveedor
            WHERE nr.folio = %s
        """
        query_det = """
            SELECT 
                dnr.codigoProducto AS codigo,
                p.nombre AS nombre,
                p.unidadMedida AS unidadMedida,
                dnr.cantidad AS cantidad,
                dnr.costoUnitario AS costo,
                (dnr.cantidad * dnr.costoUnitario) AS subtotal,
                p.existencia AS existencia_actual
            FROM DetalleNotaRecibo dnr
            INNER JOIN Producto p ON dnr.codigoProducto = p.codigo
            WHERE dnr.folioNota = %s
        """
        conn = None
        try:
            conn = self.get_connection()
            cursor = conn.cursor(dictionary=True)
            cursor.execute(query_cab, (folio,))
            cab = cursor.fetchone()
            if not cab:
                return None
            cab["total"] = float(cab["total"] or 0.0)

            cursor.execute(query_det, (folio,))
            detalles = cursor.fetchall()
            for d in detalles:
                d["cantidad"] = int(d["cantidad"])
                d["costo"] = float(d["costo"])
                d["subtotal"] = float(d["subtotal"])
                d["existencia_actual"] = int(d["existencia_actual"])
            cab["detalles"] = detalles
            return cab
        finally:
            if conn and conn.is_connected():
                cursor.close()
                conn.close()

    def agregar_nota_recibo(self, folio: int, fecha_str: str, id_proveedor: int, detalles: list[dict]):
        if not detalles:
            raise ValueError("La nota de recibo debe incluir al menos un producto recibido.")

        conn = None
        try:
            conn = self.get_connection()
            cursor = conn.cursor(dictionary=True)

            cursor.execute("SELECT folio FROM NotaRecibo WHERE folio = %s", (folio,))
            if cursor.fetchone():
                raise ValueError(f"El folio de nota NR-{folio:04d} ya existe.")

            total_nota = 0.0
            for item in detalles:
                cod = item["codigo"]
                cant = int(item["cantidad"])
                costo = float(item["costo"])
                if cant <= 0 or costo <= 0:
                    raise ValueError("La cantidad y el costo de cada producto deben ser mayores a 0.")

                cursor.execute("SELECT codigo, nombre FROM Producto WHERE codigo = %s", (cod,))
                if not cursor.fetchone():
                    raise ValueError(
                        f"El producto con código '{cod}' no está registrado en el sistema. "
                        "Debe registrarlo antes en el módulo de Productos."
                    )
                total_nota += cant * costo

            cursor.execute(
                """
                INSERT INTO NotaRecibo (folio, fecha, total, idProveedor, estado)
                VALUES (%s, %s, %s, %s, 'Activa')
                """,
                (folio, fecha_str, total_nota, id_proveedor)
            )

            for item in detalles:
                cod = item["codigo"]
                cant = int(item["cantidad"])
                costo = float(item["costo"])

                cursor.execute(
                    """
                    INSERT INTO DetalleNotaRecibo (folioNota, codigoProducto, cantidad, costoUnitario)
                    VALUES (%s, %s, %s, %s)
                    """,
                    (folio, cod, cant, costo)
                )

                cursor.execute(
                    """
                    UPDATE Producto
                    SET existencia = existencia + %s,
                        costo = %s
                    WHERE codigo = %s
                    """,
                    (cant, costo, cod)
                )

            conn.commit()
        except Exception:
            if conn and conn.is_connected():
                conn.rollback()
            raise
        finally:
            if conn and conn.is_connected():
                cursor.close()
                conn.close()

    def modificar_nota_recibo(self, folio: int, nuevos_detalles: list[dict]):
        if not nuevos_detalles:
            raise ValueError("La nota de recibo debe conservar al menos un producto. Si desea anularla, cámbiela a Cancelada.")

        conn = None
        try:
            conn = self.get_connection()
            cursor = conn.cursor(dictionary=True)

            cursor.execute("SELECT estado FROM NotaRecibo WHERE folio = %s", (folio,))
            nota_actual = cursor.fetchone()
            if not nota_actual:
                raise ValueError("No se encontró la nota de recibo.")
            if nota_actual["estado"] != "Activa":
                raise ValueError("Solo se pueden modificar notas de recibo con estado 'Activa'.")

            cursor.execute(
                "SELECT codigoProducto, cantidad, costoUnitario FROM DetalleNotaRecibo WHERE folioNota = %s",
                (folio,)
            )
            detalles_previos = cursor.fetchall()

            for prev in detalles_previos:
                cod_prev = prev["codigoProducto"]
                cant_prev = int(prev["cantidad"])
                cursor.execute("SELECT nombre, existencia FROM Producto WHERE codigo = %s", (cod_prev,))
                prod_row = cursor.fetchone()
                if prod_row:
                    nueva_cant_para_este = sum(
                        int(x["cantidad"]) for x in nuevos_detalles if x["codigo"] == cod_prev
                    )
                    existencia_final = int(prod_row["existencia"]) - cant_prev + nueva_cant_para_este
                    if existencia_final < 0:
                        raise ValueError(
                            f"No se puede reducir o eliminar '{prod_row['nombre']}' porque dejaría la existencia "
                            f"en negativo ({existencia_final})."
                        )
                    cursor.execute(
                        "UPDATE Producto SET existencia = existencia - %s WHERE codigo = %s",
                        (cant_prev, cod_prev)
                    )

            cursor.execute("DELETE FROM DetalleNotaRecibo WHERE folioNota = %s", (folio,))

            nuevo_total = 0.0
            for item in nuevos_detalles:
                cod = item["codigo"]
                cant = int(item["cantidad"])
                costo = float(item["costo"])
                if cant <= 0 or costo <= 0:
                    raise ValueError("Las cantidades y costos deben ser mayores a 0.")

                cursor.execute(
                    """
                    INSERT INTO DetalleNotaRecibo (folioNota, codigoProducto, cantidad, costoUnitario)
                    VALUES (%s, %s, %s, %s)
                    """,
                    (folio, cod, cant, costo)
                )
                cursor.execute(
                    """
                    UPDATE Producto
                    SET existencia = existencia + %s,
                        costo = %s
                    WHERE codigo = %s
                    """,
                    (cant, costo, cod)
                )
                nuevo_total += cant * costo

            cursor.execute(
                "UPDATE NotaRecibo SET total = %s WHERE folio = %s",
                (nuevo_total, folio)
            )

            conn.commit()
        except Exception:
            if conn and conn.is_connected():
                conn.rollback()
            raise
        finally:
            if conn and conn.is_connected():
                cursor.close()
                conn.close()

    def cambiar_estado_nota_recibo(self, folio: int, nuevo_estado: str):
        conn = None
        try:
            conn = self.get_connection()
            cursor = conn.cursor(dictionary=True)

            cursor.execute("SELECT estado FROM NotaRecibo WHERE folio = %s", (folio,))
            row = cursor.fetchone()
            if not row:
                raise ValueError("La nota de recibo no existe.")
            estado_actual = row["estado"]
            if estado_actual == nuevo_estado:
                return

            cursor.execute(
                "SELECT codigoProducto, cantidad FROM DetalleNotaRecibo WHERE folioNota = %s",
                (folio,)
            )
            detalles = cursor.fetchall()

            if estado_actual == "Activa" and nuevo_estado == "Cancelada":
                for d in detalles:
                    cursor.execute(
                        "SELECT nombre, existencia FROM Producto WHERE codigo = %s",
                        (d["codigoProducto"],)
                    )
                    p = cursor.fetchone()
                    if p and int(p["existencia"]) < int(d["cantidad"]):
                        raise ValueError(
                            f"No se puede cancelar la nota porque el producto '{p['nombre']}' solo tiene "
                            f"{p['existencia']} en existencia y se descontarían {d['cantidad']}."
                        )
                    cursor.execute(
                        "UPDATE Producto SET existencia = existencia - %s WHERE codigo = %s",
                        (int(d["cantidad"]), d["codigoProducto"])
                    )
            elif estado_actual == "Cancelada" and nuevo_estado == "Activa":
                for d in detalles:
                    cursor.execute(
                        "UPDATE Producto SET existencia = existencia + %s WHERE codigo = %s",
                        (int(d["cantidad"]), d["codigoProducto"])
                    )

            cursor.execute("UPDATE NotaRecibo SET estado = %s WHERE folio = %s", (nuevo_estado, folio))
            conn.commit()
        except Exception:
            if conn and conn.is_connected():
                conn.rollback()
            raise
        finally:
            if conn and conn.is_connected():
                cursor.close()
                conn.close()

    # =========================================================================
    # REPORTE DE RECIBO DE MERCANCÍAS
    # =========================================================================
    def reporte_recibo_mercancias_detalle(self, fecha_ini: str, fecha_fin: str, id_proveedor: int | None = None) -> list[dict]:
        condiciones = ["nr.fecha BETWEEN %s AND %s", "COALESCE(nr.estado, 'Activa') = 'Activa'"]
        params = [fecha_ini, fecha_fin]

        if id_proveedor is not None:
            condiciones.append("nr.idProveedor = %s")
            params.append(id_proveedor)

        where_sql = " AND ".join(condiciones)
        query = f"""
            SELECT 
                nr.folio AS folio_nota,
                DATE_FORMAT(nr.fecha, '%d/%m/%Y') AS fecha_nota,
                pr.empresa AS proveedor,
                dnr.codigoProducto AS codigo,
                p.nombre AS nombre_producto,
                dnr.cantidad AS cantidad,
                dnr.costoUnitario AS costo_unitario,
                (dnr.cantidad * dnr.costoUnitario) AS costo_recibo
            FROM DetalleNotaRecibo dnr
            INNER JOIN NotaRecibo nr ON dnr.folioNota = nr.folio
            INNER JOIN Proveedor pr ON nr.idProveedor = pr.idProveedor
            INNER JOIN Producto p ON dnr.codigoProducto = p.codigo
            WHERE {where_sql}
            ORDER BY nr.fecha DESC, nr.folio DESC
        """
        conn = None
        try:
            conn = self.get_connection()
            cursor = conn.cursor(dictionary=True)
            cursor.execute(query, tuple(params))
            rows = cursor.fetchall()
            for r in rows:
                r["cantidad"] = int(r["cantidad"] or 0)
                r["costo_unitario"] = float(r["costo_unitario"] or 0.0)
                r["costo_recibo"] = float(r["costo_recibo"] or 0.0)
            return rows
        finally:
            if conn and conn.is_connected():
                cursor.close()
                conn.close()


    # =========================================================================
    # MÓDULO 4: CLIENTES
    # =========================================================================
    def listar_clientes(self, criterio: str = "", incluir_inactivos: bool = False) -> list[dict]:
        condiciones, params = ["1=1"], []
        if not incluir_inactivos:
            condiciones.append("c.estado = 1")
        if criterio.strip():
            patron = f"%{criterio.strip()}%"
            condiciones.append("(c.nombre LIKE %s OR CAST(c.idCliente AS CHAR) LIKE %s OR c.telefono LIKE %s OR c.rfc LIKE %s)")
            params.extend([patron, patron, patron, patron])
        sql = f"""SELECT c.idCliente,c.nombre,c.domicilio,c.numeroCalle,c.colonia,c.ciudad,
                         c.codigoPostal,c.telefono,c.correoE,c.paginaWeb,c.rfc,c.estado
                  FROM Cliente c WHERE {' AND '.join(condiciones)} ORDER BY c.idCliente DESC"""
        conn = self.get_connection()
        try:
            cur = conn.cursor(dictionary=True)
            cur.execute(sql, tuple(params))
            rows = cur.fetchall()
            for row in rows:
                row["telefono"] = str(row.get("telefono") or "")
                row["estado"] = bool(row.get("estado", True))
            return rows
        finally:
            cur.close()
            conn.close()

    def obtener_cliente(self, id_cliente: int) -> dict | None:
        filas = self.listar_clientes(str(id_cliente), incluir_inactivos=True)
        return next((x for x in filas if int(x["idCliente"]) == int(id_cliente)), None)

    def _validar_cliente(self, datos: dict, id_excluir: int | None = None):
        import re
        from datetime import date
        nombre = str(datos.get("nombre", "")).strip()
        if not nombre:
            raise ValueError("El nombre del cliente es obligatorio.")
        for campo, etiqueta in (("domicilio", "Domicilio / Calle"), ("numeroCalle", "Número de calle"),
                                ("colonia", "Colonia"), ("ciudad", "Ciudad"), ("codigoPostal", "Código Postal")):
            if not str(datos.get(campo, "")).strip():
                raise ValueError(f"{etiqueta} es obligatorio.")
        cp = str(datos.get("codigoPostal", "")).strip()
        if not re.fullmatch(r"[0-9]{5}", cp):
            raise ValueError("El Código Postal debe contener exactamente 5 dígitos.")
        tel = str(datos.get("telefono", "")).strip()
        if not re.fullmatch(r"[0-9]{10}", tel):
            raise ValueError("El teléfono debe contener exactamente 10 dígitos numéricos.")
        correo = str(datos.get("correoE", "")).strip()
        if "$" in correo or not re.fullmatch(r"[^\s@]+@[^\s@]+\.[^\s@]+", correo):
            raise ValueError("Correo inválido. Use usuario@dominio.ext y no incluya el símbolo $.")
        rfc = str(datos.get("rfc", "")).strip().upper()
        if not re.fullmatch(r"[A-ZÑ&]{4}\d{6}[A-Z0-9]{3}", rfc):
            raise ValueError("El RFC debe tener 13 caracteres: 4 letras mayúsculas, 6 dígitos YYMMDD y 3 caracteres alfanuméricos.")
        yy, mm, dd = int(rfc[4:6]), int(rfc[6:8]), int(rfc[8:10])
        # El RFC sólo aporta año de dos dígitos. Se comprueba calendario usando año bisiesto de referencia
        # y se acepta 19xx/20xx según la convención de RFC (sin imponer edad).
        year = 2000 + yy if yy <= (date.today().year % 100) else 1900 + yy
        try:
            date(year, mm, dd)
        except ValueError:
            raise ValueError("La fecha de nacimiento incluida en el RFC no es válida (YYMMDD).")
        conn = self.get_connection()
        try:
            cur = conn.cursor()
            cur.execute("SELECT idCliente, telefono, rfc FROM Cliente WHERE (telefono=%s OR rfc=%s)", (tel, rfc))
            for found in cur.fetchall():
                if id_excluir is None or int(found[0]) != int(id_excluir):
                    if str(found[1]) == tel:
                        raise ValueError("El teléfono ya está registrado en otro cliente.")
                    if str(found[2]).upper() == rfc:
                        raise ValueError("El RFC ya está registrado en otro cliente.")
        finally:
            cur.close()
            conn.close()
        return rfc, tel

    def agregar_cliente(self, datos: dict) -> int:
        rfc, tel = self._validar_cliente(datos)
        sql = """INSERT INTO Cliente(nombre,domicilio,numeroCalle,colonia,ciudad,codigoPostal,
                 telefono,correoE,paginaWeb,rfc,estado)
                 VALUES(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,TRUE)"""
        conn = self.get_connection()
        try:
            cur = conn.cursor()
            cur.execute(sql, (datos["nombre"].strip(), datos["domicilio"].strip(),
                datos["numeroCalle"].strip(), datos["colonia"].strip(), datos["ciudad"].strip(),
                str(datos["codigoPostal"]).strip(), tel, datos["correoE"].strip(),
                datos.get("paginaWeb") or None, rfc))
            conn.commit()
            return int(cur.lastrowid)
        except Exception:
            conn.rollback()
            raise
        finally:
            cur.close()
            conn.close()

    def modificar_cliente(self, id_cliente: int, datos: dict):
        rfc, tel = self._validar_cliente(datos, id_excluir=id_cliente)
        conn = self.get_connection()
        try:
            cur = conn.cursor()
            cur.execute("""UPDATE Cliente SET nombre=%s,domicilio=%s,numeroCalle=%s,colonia=%s,
                ciudad=%s,codigoPostal=%s,telefono=%s,correoE=%s,rfc=%s WHERE idCliente=%s""",
                (datos["nombre"].strip(), datos["domicilio"].strip(), datos["numeroCalle"].strip(),
                 datos["colonia"].strip(), datos["ciudad"].strip(), str(datos["codigoPostal"]).strip(),
                 tel, datos["correoE"].strip(), rfc, id_cliente))
            conn.commit()
        except Exception:
            conn.rollback()
            raise
        finally:
            cur.close()
            conn.close()

    def eliminar_cliente_logico(self, id_cliente: int):
        conn = self.get_connection()
        try:
            cur = conn.cursor()
            cur.execute("UPDATE Cliente SET estado=FALSE WHERE idCliente=%s AND estado=TRUE", (id_cliente,))
            if cur.rowcount == 0:
                raise ValueError("El cliente no existe o ya está inactivo.")
            conn.commit()
        except Exception:
            conn.rollback()
            raise
        finally:
            cur.close()
            conn.close()

    def historial_compras_cliente(self, id_cliente: int) -> list[dict]:
        conn = self.get_connection()
        try:
            cur = conn.cursor(dictionary=True)
            cur.execute("""SELECT folio,DATE_FORMAT(fecha,'%d/%m/%Y %H:%i') fecha,
                metodoPago,total,IF(estado=1,'Activo','Cancelado') estado
                FROM Ticket WHERE idCliente=%s ORDER BY fecha DESC,folio DESC""", (id_cliente,))
            rows = cur.fetchall()
            for r in rows: r["total"] = float(r["total"] or 0)
            return rows
        finally:
            cur.close()
            conn.close()

    # =========================================================================
    # MÓDULO 5: PUNTO DE VENTA, TICKETS Y MOVIMIENTOS DE INVENTARIO
    # =========================================================================
    def productos_para_venta(self, criterio: str = "") -> list[dict]:
        conn = self.get_connection()
        try:
            cur = conn.cursor(dictionary=True)
            patron = f"%{criterio.strip()}%"
            cur.execute("""SELECT codigo,nombre,existencia,precio,unidadMedida FROM Producto
                WHERE estado=TRUE AND (%s='' OR codigo LIKE %s OR nombre LIKE %s)
                ORDER BY nombre LIMIT 100""", (criterio.strip(), patron, patron))
            rows = cur.fetchall()
            for r in rows:
                r["precio"] = float(r["precio"] or 0)
                r["existencia"] = int(r["existencia"] or 0)
            return rows
        finally:
            cur.close()
            conn.close()

    def obtener_siguiente_folio_ticket(self) -> int:
        conn = self.get_connection()
        try:
            cur = conn.cursor()
            cur.execute("SELECT COALESCE(MAX(folio),0)+1 FROM Ticket")
            return int(cur.fetchone()[0])
        finally:
            cur.close()
            conn.close()

    def registrar_ticket(self, id_cliente: int | None, metodo_pago: str, efectivo: float,
                         detalles: list[dict]) -> int:
        if not detalles:
            raise ValueError("Agregue al menos un producto al ticket.")
        if metodo_pago not in ("Efectivo", "Tarjeta"):
            raise ValueError("Método de pago inválido.")
        conn = self.get_connection()
        try:
            cur = conn.cursor(dictionary=True)
            total = 0.0
            normalizados = []
            for d in detalles:
                codigo, cantidad = str(d["codigo"]), int(d["cantidad"])
                if cantidad <= 0: raise ValueError("La cantidad debe ser mayor que cero.")
                cur.execute("SELECT codigo,nombre,existencia,precio FROM Producto WHERE codigo=%s AND estado=TRUE FOR UPDATE", (codigo,))
                p = cur.fetchone()
                if not p: raise ValueError(f"El producto {codigo} no existe o está inactivo.")
                if int(p["existencia"]) < cantidad:
                    raise ValueError(f"Stock insuficiente para {p['nombre']}. Disponible: {p['existencia']}.")
                precio = float(p["precio"])
                total += cantidad * precio
                normalizados.append((codigo, cantidad, precio))
            cambio = float(efectivo) - total if metodo_pago == "Efectivo" else 0.0
            if metodo_pago == "Efectivo" and cambio < -0.00001:
                raise ValueError("El efectivo recibido debe ser mayor o igual al total.")
            if id_cliente is not None:
                cur.execute("SELECT idCliente FROM Cliente WHERE idCliente=%s AND estado=TRUE", (id_cliente,))
                if not cur.fetchone(): raise ValueError("El cliente seleccionado no está activo.")
            cur.execute("""INSERT INTO Ticket(fecha,metodoPago,descuento,total,cambio,idCliente,estado,motivoCancelacion)
                VALUES(NOW(),%s,0.00,%s,%s,%s,TRUE,NULL)""",
                (metodo_pago, round(total,2), round(max(cambio,0),2), id_cliente))
            folio = int(cur.lastrowid)
            for codigo,cantidad,precio in normalizados:
                cur.execute("""INSERT INTO DetalleTicket(folioTicket,codigoProducto,cantidad,precioUnitario,descuento)
                    VALUES(%s,%s,%s,%s,0.00)""", (folio,codigo,cantidad,precio))
                cur.execute("UPDATE Producto SET existencia=existencia-%s WHERE codigo=%s", (cantidad,codigo))
            conn.commit()
            return folio
        except Exception:
            conn.rollback()
            raise
        finally:
            cur.close()
            conn.close()

    def listar_tickets(self, folio: str = "", fecha_desde: str | None = None,
                       fecha_hasta: str | None = None, metodo_pago: str = "Todos",
                       incluir_cancelados: bool = True) -> list[dict]:
        cond, params = ["1=1"], []
        f = folio.strip().upper().replace("V-", "").strip()
        if f:
            f = f.lstrip("0") or "0"
            cond.append("CAST(t.folio AS CHAR) LIKE %s")
            params.append(f"%{f}%")
        if fecha_desde:
            cond.append("DATE(t.fecha)>=%s"); params.append(fecha_desde)
        if fecha_hasta:
            cond.append("DATE(t.fecha)<=%s"); params.append(fecha_hasta)
        if metodo_pago != "Todos":
            cond.append("t.metodoPago=%s"); params.append(metodo_pago)
        if not incluir_cancelados: cond.append("t.estado=TRUE")
        sql = f"""SELECT t.folio,DATE_FORMAT(t.fecha,'%d/%m/%Y %H:%i') fecha,
            COALESCE(c.nombre,'Público en general') cliente,t.metodoPago,t.total,t.cambio,
            t.estado,t.motivoCancelacion,t.idCliente
            FROM Ticket t LEFT JOIN Cliente c ON c.idCliente=t.idCliente
            WHERE {' AND '.join(cond)} ORDER BY t.fecha DESC,t.folio DESC"""
        conn = self.get_connection()
        try:
            cur=conn.cursor(dictionary=True); cur.execute(sql,tuple(params)); rows=cur.fetchall()
            for r in rows: r["total"]=float(r["total"] or 0); r["cambio"]=float(r["cambio"] or 0); r["estado"]=bool(r["estado"])
            return rows
        finally: cur.close(); conn.close()

    def obtener_ticket_completo(self, folio: int) -> dict | None:
        conn=self.get_connection()
        try:
            cur=conn.cursor(dictionary=True)
            cur.execute("""SELECT t.folio,DATE_FORMAT(t.fecha,'%d/%m/%Y %H:%i') fecha,
                t.metodoPago,t.total,t.cambio,t.idCliente,t.estado,t.motivoCancelacion,
                COALESCE(c.nombre,'Público en general') cliente
                FROM Ticket t LEFT JOIN Cliente c ON c.idCliente=t.idCliente WHERE t.folio=%s""",(folio,))
            cab=cur.fetchone()
            if not cab: return None
            cab["total"]=float(cab["total"] or 0); cab["cambio"]=float(cab["cambio"] or 0); cab["estado"]=bool(cab["estado"])
            cur.execute("""SELECT d.codigoProducto codigo,p.nombre,d.cantidad,d.precioUnitario precio,
                (d.cantidad*d.precioUnitario) subtotal,p.unidadMedida
                FROM DetalleTicket d JOIN Producto p ON p.codigo=d.codigoProducto
                WHERE d.folioTicket=%s ORDER BY p.nombre""",(folio,))
            cab["detalles"]=cur.fetchall()
            for d in cab["detalles"]:
                d["cantidad"]=int(d["cantidad"]); d["precio"]=float(d["precio"]); d["subtotal"]=float(d["subtotal"])
            return cab
        finally: cur.close(); conn.close()

    def modificar_ticket_detalle(self, folio: int, detalles_nuevos: list[dict]):
        if not detalles_nuevos: raise ValueError("El ticket debe conservar al menos un producto.")
        conn=self.get_connection()
        try:
            cur=conn.cursor(dictionary=True)
            cur.execute("SELECT estado,total,cambio,metodoPago FROM Ticket WHERE folio=%s FOR UPDATE",(folio,))
            ticket=cur.fetchone()
            if not ticket: raise ValueError("Ticket no encontrado.")
            if not bool(ticket["estado"]): raise ValueError("No se puede modificar un ticket cancelado.")
            cur.execute("SELECT codigoProducto,cantidad,precioUnitario FROM DetalleTicket WHERE folioTicket=%s FOR UPDATE",(folio,))
            anteriores={r["codigoProducto"]:(int(r["cantidad"]),float(r["precioUnitario"])) for r in cur.fetchall()}
            nuevos={}
            for d in detalles_nuevos:
                cod=str(d["codigo"]); cant=int(d["cantidad"])
                if cant<=0: raise ValueError("Las cantidades deben ser mayores que cero.")
                if cod in nuevos: raise ValueError("No repita productos; aumente la cantidad del renglón.")
                nuevos[cod]=cant
            total=0.0
            for cod,cant_nuevo in nuevos.items():
                cur.execute("SELECT existencia,precio,nombre,estado FROM Producto WHERE codigo=%s FOR UPDATE",(cod,))
                prod=cur.fetchone()
                if not prod or (not bool(prod["estado"]) and cod not in anteriores):
                    raise ValueError(f"El producto {cod} no está activo.")
                cantidad_anterior=anteriores.get(cod,(0,0.0))[0]
                delta=cant_nuevo-cantidad_anterior
                if delta>int(prod["existencia"]):
                    raise ValueError(f"Stock insuficiente para {prod['nombre']}. Disponible: {prod['existencia']}.")
                cur.execute("UPDATE Producto SET existencia=existencia-%s WHERE codigo=%s",(delta,cod))
                precio=anteriores[cod][1] if cod in anteriores else float(prod["precio"])
                total += cant_nuevo*precio
                cur.execute("""INSERT INTO DetalleTicket(folioTicket,codigoProducto,cantidad,precioUnitario,descuento)
                    VALUES(%s,%s,%s,%s,0.00) ON DUPLICATE KEY UPDATE cantidad=VALUES(cantidad),precioUnitario=VALUES(precioUnitario),descuento=0.00""",
                    (folio,cod,cant_nuevo,precio))
            for cod,(cant_ant,precio_ant) in anteriores.items():
                if cod not in nuevos:
                    cur.execute("UPDATE Producto SET existencia=existencia+%s WHERE codigo=%s",(cant_ant,cod))
                    cur.execute("DELETE FROM DetalleTicket WHERE folioTicket=%s AND codigoProducto=%s",(folio,cod))
            nuevo_cambio=0.0
            if ticket["metodoPago"]=="Efectivo":
                efectivo_original=float(ticket["total"] or 0)+float(ticket["cambio"] or 0)
                if round(total,2)>round(efectivo_original,2):
                    raise ValueError("El nuevo total supera el efectivo recibido originalmente. No se puede modificar sin registrar un pago adicional.")
                nuevo_cambio=round(efectivo_original-total,2)
            cur.execute("UPDATE Ticket SET total=%s,cambio=%s,descuento=0.00 WHERE folio=%s",(round(total,2),nuevo_cambio,folio))
            conn.commit()
        except Exception:
            conn.rollback(); raise
        finally: cur.close(); conn.close()

    def cancelar_ticket(self, folio: int, motivo: str):
        motivo=str(motivo).strip()
        if not motivo: raise ValueError("El motivo de cancelación es obligatorio.")
        if len(motivo)>255: raise ValueError("El motivo no puede exceder 255 caracteres.")
        conn=self.get_connection()
        try:
            cur=conn.cursor(dictionary=True)
            cur.execute("SELECT estado FROM Ticket WHERE folio=%s FOR UPDATE",(folio,))
            t=cur.fetchone()
            if not t: raise ValueError("Ticket no encontrado.")
            if not bool(t["estado"]): raise ValueError("El ticket ya está cancelado.")
            cur.execute("SELECT codigoProducto,cantidad FROM DetalleTicket WHERE folioTicket=%s",(folio,))
            for d in cur.fetchall():
                cur.execute("UPDATE Producto SET existencia=existencia+%s WHERE codigo=%s",(int(d["cantidad"]),d["codigoProducto"]))
            cur.execute("UPDATE Ticket SET estado=FALSE,motivoCancelacion=%s WHERE folio=%s",(motivo,folio))
            conn.commit()
        except Exception:
            conn.rollback(); raise
        finally: cur.close(); conn.close()

    def reporte_ventas(self, fecha_ini: str, fecha_fin: str, metodo_pago: str = "Todos") -> list[dict]:
        cond=["t.estado=TRUE","DATE(t.fecha) BETWEEN %s AND %s"]; params=[fecha_ini,fecha_fin]
        if metodo_pago!="Todos": cond.append("t.metodoPago=%s"); params.append(metodo_pago)
        sql=f"""SELECT t.folio,DATE_FORMAT(t.fecha,'%d/%m/%Y %H:%i') fecha,
            COALESCE(c.nombre,'Público en general') cliente,t.metodoPago,t.total
            FROM Ticket t LEFT JOIN Cliente c ON c.idCliente=t.idCliente
            WHERE {' AND '.join(cond)} ORDER BY t.fecha,t.folio"""
        conn=self.get_connection()
        try:
            cur=conn.cursor(dictionary=True); cur.execute(sql,tuple(params)); rows=cur.fetchall()
            for r in rows:r["total"]=float(r["total"] or 0)
            return rows
        finally:cur.close();conn.close()

    def sembrar_datos_prueba_si_vacio(self):
        """Siembra catálogos y ejemplos de clientes/ventas de forma independiente."""
        conn = self.get_connection()
        cur = conn.cursor()
        try:
            cur.execute("SELECT COUNT(*) FROM Proveedor")
            if cur.fetchone()[0] == 0:
                cur.executemany("""INSERT INTO Proveedor(nombre,domicilio,telefono,correoE,paginaWeb,empresa,rfc,estado)
                    VALUES(%s,%s,%s,%s,%s,%s,%s,TRUE)""", [
                    ("Juan Pérez","Blvd. Zacatecas Norte 845, Aguascalientes","4499102030","ventas@tecnocel.mx","www.tecnocel.mx","Tecnocel Distribuciones S.A. de C.V.","TCE930512AB1"),
                    ("María Gómez","Av. López Mateos 1204, Aguascalientes","4491557788","pedidos@mac.com.mx","www.mayoreocentro.mx","Mayoreo Accesorios del Centro","MAC880101XY2"),
                    ("Roberto Díaz","Av. Tecnológico 310, Aguascalientes","4773201144","contacto@rmbajio.mx","www.refaccionesbajio.mx","Refacciones Móviles del Bajío","RMB050707QQ3")])
            cur.execute("SELECT COUNT(*) FROM Producto")
            if cur.fetchone()[0] == 0:
                cur.executemany("""INSERT INTO Producto(codigo,nombre,categoria,unidadMedida,existencia,stockMinimo,stockMaximo,costo,precio,estado)
                    VALUES(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)""", [
                    ("7501234500011","Mica cristal templado iPhone 13","Micas","pieza",24,10,30,35,120,True),
                    ("7501234500028","Cargador USB-C 20 W","Accesorios","pieza",14,8,16,140,280,True),
                    ("7501234500035","Funda silicón Galaxy A54","Accesorios","pieza",15,5,20,55,150,True),
                    ("7501234500042","Pantalla OLED iPhone 11","Refacciones","pieza",2,3,6,900,1450,True),
                    ("7501234500059","Batería Moto G32","Refacciones","pieza",6,4,12,230,450,True),
                    ("7501234500066","Cable Lightning 1 m","Accesorios","pieza",12,6,12,45,120,True),
                    ("7501234500073","Audífonos Bluetooth TWS","Accesorios","pieza",9,4,15,190,390,True)])
            cur.execute("SELECT COUNT(*) FROM NotaRecibo")
            if cur.fetchone()[0] == 0:
                cur.execute("SELECT idProveedor FROM Proveedor ORDER BY idProveedor LIMIT 1")
                proveedor=cur.fetchone()
                if proveedor:
                    codigos_nota=["7501234500028","7501234500011","7501234500073","7501234500066","7501234500042"]
                    cur.execute("SELECT codigo FROM Producto WHERE codigo IN (%s,%s,%s,%s,%s)",tuple(codigos_nota))
                    existentes={r[0] for r in cur.fetchall()}
                    if set(codigos_nota).issubset(existentes):
                        cur.execute("INSERT INTO NotaRecibo(folio,fecha,total,idProveedor,estado) VALUES(41,'2026-10-15',2300,%s,'Activa')",(proveedor[0],))
                        cur.execute("INSERT INTO NotaRecibo(folio,fecha,total,idProveedor,estado) VALUES(42,'2026-10-19',4640,%s,'Activa')",(proveedor[0],))
                        cur.executemany("INSERT INTO DetalleNotaRecibo(folioNota,codigoProducto,cantidad,costoUnitario) VALUES(%s,%s,%s,%s)",[
                            (41,"7501234500028",6,140),(41,"7501234500011",20,35),(41,"7501234500073",4,190),
                            (42,"7501234500028",10,140),(42,"7501234500066",12,45),(42,"7501234500042",3,900)])
            cur.execute("SELECT COUNT(*) FROM Cliente")
            if cur.fetchone()[0] == 0:
                cur.executemany("""INSERT INTO Cliente(nombre,domicilio,numeroCalle,colonia,ciudad,codigoPostal,
                    telefono,correoE,rfc,estado) VALUES(%s,%s,%s,%s,%s,%s,%s,%s,%s,TRUE)""", [
                    ("Ana López","Av. Universidad","120","Bosques","Aguascalientes","20100","4491234567","ana.lopez@example.com","LOPA900215AB1"),
                    ("Carlos Martínez","Calle Morelos","45","Centro","Aguascalientes","20000","4497654321","carlos.martinez@example.com","MARC850731CD2"),
                    ("Público frecuente","Av. Convención","300","Las Flores","Aguascalientes","20120","4491122334","publico@example.com","PUFR920101EF3")])
            cur.execute("SELECT COUNT(*) FROM DetalleTicket")
            if cur.fetchone()[0] == 0:
                cur.execute("SELECT codigo,precio FROM Producto WHERE estado=TRUE AND existencia>0 ORDER BY codigo LIMIT 1")
                producto=cur.fetchone()
                if producto:
                    cur.execute("SELECT idCliente FROM Cliente WHERE estado=TRUE ORDER BY idCliente LIMIT 1")
                    cliente=cur.fetchone()
                    precio=float(producto[1] or 0)
                    cur.execute("""INSERT INTO Ticket(fecha,metodoPago,descuento,total,cambio,idCliente,estado,motivoCancelacion)
                        VALUES(NOW() - INTERVAL 2 DAY,'Efectivo',0.00,%s,0.00,%s,TRUE,NULL)""",
                        (round(precio,2),cliente[0] if cliente else None))
                    folio=cur.lastrowid
                    cur.execute("""INSERT INTO DetalleTicket(folioTicket,codigoProducto,cantidad,precioUnitario,descuento)
                        VALUES(%s,%s,1,%s,0.00)""",(folio,producto[0],precio))
                    cur.execute("UPDATE Producto SET existencia=existencia-1 WHERE codigo=%s",(producto[0],))
            conn.commit()
        except Exception:
            conn.rollback()
            raise
        finally:
            cur.close()
            conn.close()

import sys
import unicodedata
from datetime import datetime
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QGridLayout, QFrame, QLabel, QLineEdit, QPushButton, QComboBox,
    QDoubleSpinBox, QSpinBox, QStackedWidget, QTableWidget,
    QTableWidgetItem, QHeaderView, QDialog, QMessageBox, QCheckBox,
    QTabWidget, QProgressBar, QDateEdit, QFileDialog, QCompleter, QInputDialog
)
from PySide6.QtCore import Qt, QDate, QMarginsF, QRegularExpression
from PySide6.QtGui import (
    QColor, QBrush, QTextDocument, QPageSize, QPageLayout,
    QPalette, QRegularExpressionValidator, QTextCharFormat
)
from PySide6.QtPrintSupport import QPrinter

from db_spvci import SPVCIRepository

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    OPENPYXL_DISPONIBLE = True
except ImportError:
    OPENPYXL_DISPONIBLE = False


NOMBRE_EMPRESA = "Los Expertos"
COLOR_TEXTO_OSCURO = QColor("#0F172A")

# =====================================================================
# HOJA DE ESTILOS (Soluciona el marco gris de Reportes y el Calendario oscuro)
# =====================================================================
STYLESHEET = """
QMainWindow, QDialog {
    background-color: #EEF2FF;
    font-family: 'Segoe UI', Arial, sans-serif;
    color: #0F172A;
}
QWidget {
    color: #0F172A;
}
QLabel {
    color: #0F172A;
    background-color: transparent;
}
QCheckBox {
    color: #0F172A;
    font-size: 13px;
    font-weight: 600;
    spacing: 6px;
    background-color: transparent;
}
QCheckBox::indicator {
    width: 16px;
    height: 16px;
    border: 1px solid #64748B;
    border-radius: 3px;
    background-color: #FFFFFF;
}
QCheckBox::indicator:checked {
    background-color: #4F46E5;
    border: 1px solid #3730A3;
}
/* Barra lateral */
QFrame#sidebar {
    background-color: #172554;
}
QLabel#page_title {
    font-size: 25px;
    font-weight: bold;
    color: #0F172A;
}
QLabel#logo_title {
    color: #FFFFFF;
    font-size: 22px;
    font-weight: bold;
}
QLabel#logo_subtitle {
    color: #94A3B8;
    font-size: 11px;
}
QPushButton.nav_btn {
    background-color: transparent;
    color: #C7D2E8;
    text-align: left;
    padding: 10px 16px;
    border: none;
    border-radius: 6px;
    font-size: 13px;
}
QPushButton.nav_btn:hover {
    background-color: #283B78;
    color: #FFFFFF;
}
QPushButton.nav_btn_active {
    background-color: #4F46E5;
    color: #FFFFFF;
    text-align: left;
    padding: 10px 16px;
    border: none;
    border-radius: 6px;
    font-size: 13px;
    font-weight: bold;
}
/* Contenedor de Pestañas de Reportes (Elimina el marco gris oscuro) */
QTabWidget::pane {
    background-color: #F7F8FF;
    border: 1px solid #C7D2E8;
    border-radius: 8px;
    padding: 10px;
}
QTabBar::tab {
    background-color: #E4E8F7;
    color: #0F172A;
    padding: 10px 18px;
    font-weight: bold;
    border: 1px solid #C7D2E8;
    border-bottom: none;
    border-top-left-radius: 6px;
    border-top-right-radius: 6px;
    margin-right: 4px;
}
QTabBar::tab:hover {
    background-color: #E0E7FF;
    color: #3730A3;
}
QTabBar::tab:selected {
    background-color: #4F46E5;
    color: #FFFFFF;
    border-color: #4F46E5;
}
/* Tarjetas blancas */
QFrame.card {
    background-color: #FFFFFF;
    border: 1px solid #C7D2E8;
    border-radius: 8px;
}
QLabel.card_title {
    color: #0F172A;
    font-size: 14px;
    font-weight: bold;
}
QLabel.field_label {
    color: #0F172A;
    font-size: 12px;
    font-weight: bold;
}
/* Inputs y desplegables */
QLineEdit, QComboBox, QDoubleSpinBox, QSpinBox, QDateEdit {
    background-color: #FFFFFF;
    border: 1px solid #C7D2E8;
    border-radius: 8px;
    padding: 7px 10px;
    color: #0F172A;
    font-size: 13px;
}
QLineEdit:focus, QComboBox:focus, QDoubleSpinBox:focus, QSpinBox:focus, QDateEdit:focus {
    border: 2px solid #4F46E5;
}
QComboBox QAbstractItemView, QListView {
    background-color: #FFFFFF;
    color: #0F172A;
    selection-background-color: #E0E7FF;
    selection-color: #3730A3;
    border: 1px solid #94A3B8;
    outline: 0px;
    padding: 4px;
}
QLineEdit:read-only, QSpinBox:disabled, QDoubleSpinBox:disabled, QDateEdit:disabled, QComboBox:disabled {
    background-color: #E4E8F7;
    color: #475569;
    font-weight: 600;
}
/* Calendario Desplegable (Alto contraste claro para fechas) */
QCalendarWidget QWidget {
    background-color: #FFFFFF;
    color: #0F172A;
}
QCalendarWidget QWidget#qt_calendar_navigationbar {
    background-color: #E0E7FF;
    padding: 4px;
}
QCalendarWidget QToolButton {
    background-color: #E0E7FF;
    color: #172554;
    font-weight: bold;
    font-size: 13px;
    border: none;
    border-radius: 4px;
    padding: 4px 8px;
}

QCalendarWidget QToolButton:hover {
    background-color: #C7D2FE;
    color: #3730A3;
}
QCalendarWidget QMenu {
    background-color: #FFFFFF;
    color: #0F172A;
    border: 1px solid #C7D2E8;
}
QCalendarWidget QSpinBox {
    background-color: #FFFFFF;
    color: #0F172A;
    font-weight: bold;
    border-radius: 4px;
}
QCalendarWidget QAbstractItemView:enabled {
    background-color: #FFFFFF;
    color: #0F172A;
    selection-background-color: #4F46E5;
    selection-color: #FFFFFF;
    font-size: 12px;
    font-weight: 600;
    outline: 0px;
}
QCalendarWidget QAbstractItemView:disabled {
    color: #94A3B8;
}
/* Botones */
QPushButton#btn_success:hover { background-color: #0F766E; }
QPushButton#btn_primary:hover { background-color: #3730A3; }
QPushButton#btn_danger:hover { background-color: #BE123C; }
QPushButton#btn_success {
    background-color: #0D9488; color: #FFFFFF; border: none; border-radius: 6px;
    padding: 8px 16px; font-weight: bold; font-size: 13px;
}
QPushButton#btn_primary {
    background-color: #4F46E5; color: #FFFFFF; border: none; border-radius: 6px;
    padding: 8px 16px; font-weight: bold; font-size: 13px;
}
QPushButton#btn_danger {
    background-color: #E11D48; color: #FFFFFF; border: none; border-radius: 6px;
    padding: 8px 16px; font-weight: bold; font-size: 13px;
}
QPushButton#btn_secondary {
    background-color: #FFFFFF; color: #0F172A; border: 1px solid #94A3B8;
    border-radius: 6px; padding: 8px 16px; font-size: 13px; font-weight: 600;
}
QPushButton#btn_secondary:hover { background-color: #EEF2FF; }
/* Controles numéricos: botones de flechas con contraste */
QSpinBox::up-button, QDoubleSpinBox::up-button, QDateEdit::up-button,
QSpinBox::down-button, QDoubleSpinBox::down-button, QDateEdit::down-button {
    background-color: #E4E8F7;
    border: 1px solid #94A3B8;
    width: 18px;
}
QSpinBox::up-button:hover, QDoubleSpinBox::up-button:hover,
QDateEdit::up-button:hover, QSpinBox::down-button:hover,
QDoubleSpinBox::down-button:hover, QDateEdit::down-button:hover {
    background-color: #C7D2E8;
}
QSpinBox::up-arrow, QDoubleSpinBox::up-arrow, QDateEdit::up-arrow,
QSpinBox::down-arrow, QDoubleSpinBox::down-arrow, QDateEdit::down-arrow {
    width: 8px; height: 8px;
}
QPushButton#btn_report_generate {
    background-color: #4F46E5; color: #FFFFFF; border: none;
    border-radius: 6px; padding: 8px 16px; font-weight: bold; font-size: 13px;
}
QPushButton#btn_report_generate:hover { background-color: #3730A3; }
/* Tablas */
QTableWidget {
    background-color: #FFFFFF; alternate-background-color: #F7F8FF;
    border: 1px solid #C7D2E8; border-radius: 8px; gridline-color: #E4E8F7;
    color: #0F172A; font-size: 12px;
}
QTableWidget::item {
    padding: 4px 8px;
}
QTableWidget::item:hover {
    background-color: #EEF2FF;
}
QTableWidget::item:selected {
    background-color: #E0E7FF; color: #3730A3;
}
QHeaderView::section {
    background-color: #E0E7FF; color: #33456F; font-weight: bold;
    padding: 10px; border: none; font-size: 12px;
}
/* Estados deshabilitados y barras de progreso */
QPushButton:disabled, QPushButton#btn_success:disabled,
QPushButton#btn_primary:disabled, QPushButton#btn_danger:disabled,
QPushButton#btn_secondary:disabled, QPushButton#btn_report_generate:disabled {
    background-color: #E4E8F7;
    color: #64748B;
    border: 1px solid #C7D2E8;
}
QProgressBar {
    background-color: #E4E8F7;
    border: none;
    border-radius: 6px;
    text-align: center;
}
QProgressBar::chunk {
    background-color: #4F46E5;
    border-radius: 6px;
}
"""


def normalizar_texto(texto: str) -> str:
    """Pasa a minúsculas y quita acentos, para que 'bateria' encuentre 'Batería'."""
    base = unicodedata.normalize("NFD", str(texto))
    return "".join(c for c in base if unicodedata.category(c) != "Mn").lower()


def crear_celda(texto: str, alineacion=Qt.AlignLeft | Qt.AlignVCenter, negrita: bool = False) -> QTableWidgetItem:
    item = QTableWidgetItem(str(texto))
    item.setForeground(QBrush(COLOR_TEXTO_OSCURO))
    item.setTextAlignment(alineacion)
    if negrita:
        f = item.font()
        f.setBold(True)
        item.setFont(f)
    return item


def configurar_date_edit_claro(dt_edit: QDateEdit):
    """Fuerza colores claros y legibles en el QCalendarWidget interno de un QDateEdit."""
    dt_edit.setCalendarPopup(True)
    cal = dt_edit.calendarWidget()
    if cal:
        cal.setGridVisible(True)
        fmt_header = QTextCharFormat()
        fmt_header.setForeground(QBrush(QColor("#0F172A")))
        fmt_header.setBackground(QBrush(QColor("#E4E8F7")))
        f = fmt_header.font()
        f.setBold(True)
        fmt_header.setFont(f)
        cal.setHeaderTextFormat(fmt_header)

        fmt_weekday = QTextCharFormat()
        fmt_weekday.setForeground(QBrush(QColor("#0F172A")))
        for day in (Qt.Monday, Qt.Tuesday, Qt.Wednesday, Qt.Thursday, Qt.Friday):
            cal.setWeekdayTextFormat(day, fmt_weekday)

        fmt_weekend = QTextCharFormat()
        fmt_weekend.setForeground(QBrush(QColor("#E11D48")))
        cal.setWeekdayTextFormat(Qt.Saturday, fmt_weekend)
        cal.setWeekdayTextFormat(Qt.Sunday, fmt_weekend)


def normalizar_producto(p: dict) -> dict:
    return {
        "codigo": str(p.get("codigo", p.get("CODIGO", ""))),
        "nombre": str(p.get("nombre", p.get("NOMBRE", ""))),
        "categoria": str(p.get("categoria", p.get("CATEGORIA", "Accesorios"))),
        "unidadMedida": str(p.get("unidadMedida", p.get("unidadmedida", p.get("UNIDADMEDIDA", "pieza")))),
        "existencia": int(p.get("existencia", p.get("EXISTENCIA", 0)) or 0),
        "stockMinimo": int(p.get("stockMinimo", p.get("stockminimo", p.get("STOCKMINIMO", 1))) or 1),
        "stockMaximo": int(p.get("stockMaximo", p.get("stockmaximo", p.get("STOCKMAXIMO", 10))) or 10),
        "costo": float(p.get("costo", p.get("COSTO", 0.0)) or 0.0),
        "precio": float(p.get("precio", p.get("PRECIO", 0.0)) or 0.0),
        "estado": bool(p.get("estado", p.get("ESTADO", True))),
        "ventas_registradas": int(p.get("ventas_registradas", 0) or 0),
    }


def normalizar_proveedor(pr: dict) -> dict:
    return {
        "idProveedor": int(pr.get("idProveedor", pr.get("idproveedor", 0)) or 0),
        "nombre": str(pr.get("nombre", "") or ""),
        "domicilio": str(pr.get("domicilio", "") or ""),
        "telefono": str(pr.get("telefono", "") or ""),
        "correoE": str(pr.get("correoE", pr.get("correoe", "")) or ""),
        "paginaWeb": str(pr.get("paginaWeb", pr.get("paginaweb", "")) or ""),
        "empresa": str(pr.get("empresa", "") or ""),
        "rfc": str(pr.get("rfc", "") or ""),
        "estado": bool(pr.get("estado", pr.get("ESTADO", True))),
        "notas_registradas": int(pr.get("notas_registradas", 0) or 0),
    }



def numero_a_letras_mx(monto: float) -> str:
    """Convierte importes de 0 a millones a texto monetario mexicano."""
    unidades = ("", "UN", "DOS", "TRES", "CUATRO", "CINCO", "SEIS", "SIETE", "OCHO", "NUEVE")
    especiales = {10:"DIEZ",11:"ONCE",12:"DOCE",13:"TRECE",14:"CATORCE",15:"QUINCE",
                  16:"DIECISÉIS",17:"DIECISIETE",18:"DIECIOCHO",19:"DIECINUEVE",
                  20:"VEINTE",21:"VEINTIÚN",22:"VEINTIDÓS",23:"VEINTITRÉS",24:"VEINTICUATRO",
                  25:"VEINTICINCO",26:"VEINTISÉIS",27:"VEINTISIETE",28:"VEINTIOCHO",29:"VEINTINUEVE"}
    decenas = ("", "", "", "TREINTA", "CUARENTA", "CINCUENTA", "SESENTA", "SETENTA", "OCHENTA", "NOVENTA")
    centenas = ("", "CIENTO", "DOSCIENTOS", "TRESCIENTOS", "CUATROCIENTOS",
                "QUINIENTOS", "SEISCIENTOS", "SETECIENTOS", "OCHOCIENTOS", "NOVECIENTOS")
    def menor_mil(n):
        if n == 0: return ""
        if n == 100: return "CIEN"
        c, resto = divmod(n,100)
        partes = [centenas[c]] if c else []
        if resto:
            if resto < 10: partes.append(unidades[resto])
            elif resto < 30: partes.append(especiales[resto])
            else:
                d,u=divmod(resto,10)
                partes.append(decenas[d] + ((" Y " + unidades[u]) if u else ""))
        return " ".join(partes)
    def entero_letras(n):
        if n == 0: return "CERO"
        partes=[]
        millones, resto=divmod(n,1_000_000)
        miles, unidades_n=divmod(resto,1000)
        if millones:
            partes.append("UN MILLÓN" if millones == 1 else menor_mil(millones) + " MILLONES")
        if miles:
            partes.append("MIL" if miles == 1 else menor_mil(miles) + " MIL")
        if unidades_n: partes.append(menor_mil(unidades_n))
        return " ".join(partes)
    centavos = int(round((float(monto) - int(float(monto))) * 100))
    pesos = int(float(monto))
    if centavos == 100:
        pesos += 1
        centavos = 0
    moneda = "PESO" if pesos == 1 else "PESOS"
    return f"SON: {entero_letras(pesos)} {moneda} {centavos:02d}/100 M.N."

# =====================================================================
# MODAL ELIMINAR PRODUCTO
# =====================================================================
class DialogEliminarProducto(QDialog):
    def __init__(self, prod: dict, parent=None):
        super().__init__(parent)
        self.prod = prod
        self.setWindowTitle("Eliminar producto")
        self.setFixedSize(480, 310)
        self.setModal(True)
        self.setStyleSheet(STYLESHEET)
        self.setup_ui()

    def setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 24, 24, 24)
        layout.setSpacing(12)

        lbl_t = QLabel("¿Eliminar producto?")
        lbl_t.setStyleSheet("font-size: 17px; font-weight: bold; color: #0F172A;")
        layout.addWidget(lbl_t)

        grid = QGridLayout()
        grid.setVerticalSpacing(8)
        grid.setHorizontalSpacing(16)

        datos = [
            ("SKU / Código:", self.prod["codigo"]),
            ("Nombre:", self.prod["nombre"]),
            ("Existencia actual:", f"{self.prod['existencia']} {self.prod['unidadMedida']}(s)"),
            ("Ventas registradas:", f"{self.prod.get('ventas_registradas', 0)} renglones en historial"),
        ]
        for i, (k, v) in enumerate(datos):
            lk = QLabel(k)
            lk.setStyleSheet("color: #475569; font-size: 12px; font-weight: 600;")
            lv = QLabel(str(v))
            lv.setStyleSheet("color: #0F172A; font-size: 13px; font-weight: bold;")
            grid.addWidget(lk, i, 0)
            grid.addWidget(lv, i, 1)
        layout.addLayout(grid)

        puede_eliminar = (self.prod["existencia"] == 0)

        box_aviso = QFrame()
        av_layout = QVBoxLayout(box_aviso)
        lbl_aviso = QLabel()
        lbl_aviso.setWordWrap(True)

        if puede_eliminar:
            box_aviso.setStyleSheet("background-color: #FFFBEB; border: 1px solid #FCD34D; border-radius: 6px;")
            lbl_aviso.setStyleSheet("color: #92400E; font-size: 12px; border: none;")
            lbl_aviso.setText(
                "La existencia es 0. Se realizará una eliminación lógica (Inactivo): "
                "el producto dejará de mostrarse como activo sin borrar su historial."
            )
        else:
            box_aviso.setStyleSheet("background-color: #FFF1F2; border: 1px solid #FDA4AF; border-radius: 6px;")
            lbl_aviso.setStyleSheet("color: #9F1239; font-size: 12px; font-weight: bold; border: none;")
            lbl_aviso.setText(
                f"No se puede eliminar este producto porque tiene {self.prod['existencia']} "
                "unidad(es) en existencia. La existencia debe ser 0 para inactivarlo."
            )

        av_layout.addWidget(lbl_aviso)
        layout.addWidget(box_aviso)
        layout.addStretch()

        btns = QHBoxLayout()
        btns.addStretch()
        btn_cancel = QPushButton("Cancelar")
        btn_cancel.setObjectName("btn_secondary")
        btn_cancel.clicked.connect(self.reject)

        btn_del = QPushButton("Eliminar")
        btn_del.setObjectName("btn_danger")
        btn_del.setEnabled(puede_eliminar)
        btn_del.clicked.connect(self.accept)

        btns.addWidget(btn_cancel)
        btns.addWidget(btn_del)
        layout.addLayout(btns)


# =====================================================================
# MODAL ELIMINAR PROVEEDOR (Eliminación Lógica)
# =====================================================================
class DialogEliminarProveedor(QDialog):
    def __init__(self, prov: dict, parent=None):
        super().__init__(parent)
        self.prov = prov
        self.setWindowTitle("Eliminar proveedor")
        self.setFixedSize(480, 310)
        self.setModal(True)
        self.setStyleSheet(STYLESHEET)
        self.setup_ui()

    def setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 24, 24, 24)
        layout.setSpacing(12)

        lbl_t = QLabel("¿Eliminar proveedor?")
        lbl_t.setStyleSheet("font-size: 17px; font-weight: bold; color: #0F172A;")
        layout.addWidget(lbl_t)

        grid = QGridLayout()
        grid.setVerticalSpacing(8)
        grid.setHorizontalSpacing(16)

        notas_cnt = self.prov.get("notas_registradas", 0)
        datos = [
            ("ID:", str(self.prov["idProveedor"])),
            ("RFC:", self.prov["rfc"]),
            ("Nombre de la empresa:", self.prov["empresa"]),
            ("Notas de recibo registradas:", "Ninguna" if notas_cnt == 0 else f"{notas_cnt} nota(s)"),
        ]
        for i, (k, v) in enumerate(datos):
            lk = QLabel(k)
            lk.setStyleSheet("color: #475569; font-size: 12px; font-weight: 600;")
            lv = QLabel(str(v))
            lv.setStyleSheet("color: #0F172A; font-size: 13px; font-weight: bold;")
            grid.addWidget(lk, i, 0)
            grid.addWidget(lv, i, 1)
        layout.addLayout(grid)

        box_aviso = QFrame()
        box_aviso.setStyleSheet("background-color: #FFFBEB; border: 1px solid #FCD34D; border-radius: 6px;")
        av_layout = QVBoxLayout(box_aviso)
        lbl_aviso = QLabel(
            "El proveedor dejará de estar disponible como proveedor activo (eliminación lógica). "
            "Las notas de recibo ya registradas se conservan sin conflictos."
        )
        lbl_aviso.setWordWrap(True)
        lbl_aviso.setStyleSheet("color: #92400E; font-size: 12px; border: none;")
        av_layout.addWidget(lbl_aviso)
        layout.addWidget(box_aviso)
        layout.addStretch()

        btns = QHBoxLayout()
        btns.addStretch()
        btn_cancel = QPushButton("Cancelar")
        btn_cancel.setObjectName("btn_secondary")
        btn_cancel.clicked.connect(self.reject)

        btn_del = QPushButton("Eliminar")
        btn_del.setObjectName("btn_danger")
        btn_del.clicked.connect(self.accept)

        btns.addWidget(btn_cancel)
        btns.addWidget(btn_del)
        layout.addLayout(btns)



# =====================================================================
# MODAL ELIMINAR CLIENTE (BAJA LÓGICA)
# =====================================================================
class DialogEliminarCliente(QDialog):
    def __init__(self, cliente: dict, parent=None):
        super().__init__(parent)
        self.cliente=cliente
        self.setWindowTitle("Eliminar cliente")
        self.setFixedSize(440,230)
        self.setModal(True)
        self.setStyleSheet(STYLESHEET)
        layout=QVBoxLayout(self); layout.setContentsMargins(24,20,24,20); layout.setSpacing(12)
        title=QLabel("¿Eliminar cliente?"); title.setStyleSheet("font-size:18px;font-weight:bold;")
        layout.addWidget(title)
        info=QLabel(f"ID: {cliente['idCliente']}\nNombre: {cliente['nombre']}\nTeléfono: {cliente.get('telefono','')}")
        info.setStyleSheet("font-size:13px;font-weight:600;"); layout.addWidget(info)
        aviso=QLabel("El cliente pasará a estado Inactivo mediante una baja lógica. Su historial de tickets se conservará.")
        aviso.setWordWrap(True); aviso.setStyleSheet("background:#FFFBEB;color:#92400E;padding:10px;border:1px solid #FCD34D;border-radius:6px;")
        layout.addWidget(aviso); layout.addStretch()
        botones=QHBoxLayout(); botones.addStretch()
        cancelar=QPushButton("Cancelar"); cancelar.setObjectName("btn_secondary"); cancelar.clicked.connect(self.reject)
        eliminar=QPushButton("Eliminar"); eliminar.setObjectName("btn_danger"); eliminar.clicked.connect(self.accept)
        botones.addWidget(cancelar); botones.addWidget(eliminar); layout.addLayout(botones)


# =====================================================================
# VENTANA PRINCIPAL DE SPVCI LOS EXPERTOS
# =====================================================================
class SPVCIMainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("SPVCI Los Expertos · Sistema de Punto de Venta y Control de Inventario")
        self.resize(1300, 800)
        self.setStyleSheet(STYLESHEET)

        self.repo = SPVCIRepository()
        self.repo.sembrar_datos_prueba_si_vacio()

        self.modo_formulario_prod = "AGREGAR"
        self.producto_actual = None

        self.modo_formulario_prov = "AGREGAR"
        self.proveedor_actual = None

        self.modo_formulario_nota = "AGREGAR"
        self.folio_nota_editando = None
        self.renglones_nota_actual = []

        # Mapas en memoria para los buscadores con QCompleter
        self.mapa_busqueda_prods_nota = {}   # {"75012... - Nombre": dict_producto}
        self.mapa_busqueda_provs_nota = {}   # {"Empresa (RFC) [ID: 1]": idProveedor}
        self.mapa_busqueda_provs_rep = {}    # {"Empresa (RFC) [ID: 1]": idProveedor}

        self.contador_folio_rep_rec = 19
        self.modo_formulario_cliente = "AGREGAR"
        self.cliente_actual = None
        self.modo_ticket_edicion = False
        self.folio_ticket_editando = None
        self.renglones_venta = {}
        self.ticket_comprobante_actual = None

        self.setup_ui()
        self.cargar_catalogo()

    def setup_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        main_h = QHBoxLayout(central)
        main_h.setContentsMargins(0, 0, 0, 0)
        main_h.setSpacing(0)

        # Barra lateral
        sidebar = QFrame()
        sidebar.setObjectName("sidebar")
        sidebar.setFixedWidth(225)
        sb_l = QVBoxLayout(sidebar)
        sb_l.setContentsMargins(16, 20, 16, 20)
        sb_l.setSpacing(8)

        lbl_logo = QLabel("SPVCI")
        lbl_logo.setObjectName("logo_title")
        lbl_sub = QLabel("Los Expertos · Punto de venta")
        lbl_sub.setObjectName("logo_subtitle")
        sb_l.addWidget(lbl_logo)
        sb_l.addWidget(lbl_sub)
        sb_l.addSpacing(16)
        lbl_menu = QLabel("MENÚ PRINCIPAL")
        lbl_menu.setStyleSheet("color:#94A3B8;font-size:10px;font-weight:bold;padding:4px 8px;")
        sb_l.addWidget(lbl_menu)

        self.btn_nav_productos = QPushButton("Productos")
        self.btn_nav_productos.setProperty("class", "nav_btn_active")
        self.btn_nav_productos.clicked.connect(lambda: self.cambiar_seccion(0))

        self.btn_nav_proveedores = QPushButton("Proveedores")
        self.btn_nav_proveedores.setProperty("class", "nav_btn")
        self.btn_nav_proveedores.clicked.connect(lambda: self.cambiar_seccion(4))

        self.btn_nav_notas = QPushButton("Notas de recibo")
        self.btn_nav_notas.setProperty("class", "nav_btn")
        self.btn_nav_notas.clicked.connect(lambda: self.cambiar_seccion(7))

        self.btn_nav_reportes = QPushButton("Reportes")
        self.btn_nav_reportes.setProperty("class", "nav_btn")
        self.btn_nav_reportes.clicked.connect(lambda: self.cambiar_seccion(3))

        self.btn_nav_clientes = QPushButton("Clientes")
        self.btn_nav_clientes.setProperty("class", "nav_btn")
        self.btn_nav_clientes.clicked.connect(lambda: self.cambiar_seccion(10))
        self.btn_nav_venta = QPushButton("Punto de venta")
        self.btn_nav_venta.setProperty("class", "nav_btn")
        self.btn_nav_venta.clicked.connect(lambda: self.nueva_venta())
        self.btn_nav_tickets = QPushButton("Tickets")
        self.btn_nav_tickets.setProperty("class", "nav_btn")
        self.btn_nav_tickets.clicked.connect(lambda: self.cambiar_seccion(14))

        sb_l.addWidget(self.btn_nav_productos)
        sb_l.addWidget(self.btn_nav_proveedores)
        sb_l.addWidget(self.btn_nav_notas)
        sb_l.addWidget(self.btn_nav_reportes)
        sb_l.addWidget(self.btn_nav_clientes)
        sb_l.addWidget(self.btn_nav_venta)
        sb_l.addWidget(self.btn_nav_tickets)
        sb_l.addStretch()
        lbl_pie = QLabel("LOS EXPERTOS  ·  SPVCI")
        lbl_pie.setObjectName("logo_subtitle")
        sb_l.addWidget(lbl_pie)
        main_h.addWidget(sidebar)

        self.stacked = QStackedWidget()
        main_h.addWidget(self.stacked)

        self.page_catalogo = QWidget()
        self.page_formulario = QWidget()
        self.page_consultar = QWidget()
        self.page_reportes = QWidget()

        self.page_proveedores_cat = QWidget()
        self.page_proveedores_form = QWidget()
        self.page_proveedores_ver = QWidget()

        self.page_notas_cat = QWidget()
        self.page_notas_form = QWidget()
        self.page_notas_ver = QWidget()
        self.page_clientes_cat = QWidget()
        self.page_clientes_form = QWidget()
        self.page_clientes_ver = QWidget()
        self.page_punto_venta = QWidget()
        self.page_tickets = QWidget()
        self.page_comprobante = QWidget()

        self.setup_page_catalogo()
        self.setup_page_formulario()
        self.setup_page_consultar()
        self.setup_page_reportes()

        self.setup_page_proveedores_cat()
        self.setup_page_proveedores_form()
        self.setup_page_proveedores_ver()

        self.setup_page_notas_cat()
        self.setup_page_notas_form()
        self.setup_page_notas_ver()
        self.setup_page_clientes_cat()
        self.setup_page_clientes_form()
        self.setup_page_clientes_ver()
        self.setup_page_punto_venta()
        self.setup_page_tickets()
        self.setup_page_comprobante()


        for w in (
            self.page_catalogo, self.page_formulario, self.page_consultar, self.page_reportes,
            self.page_proveedores_cat, self.page_proveedores_form, self.page_proveedores_ver,
            self.page_notas_cat, self.page_notas_form, self.page_notas_ver,
            self.page_clientes_cat, self.page_clientes_form, self.page_clientes_ver,
            self.page_punto_venta, self.page_tickets, self.page_comprobante
        ):
            self.stacked.addWidget(w)

        # Apariencia uniforme de todas las tablas.
        for tabla in self.findChildren(QTableWidget):
            tabla.setAlternatingRowColors(True)
            tabla.setShowGrid(False)
            tabla.verticalHeader().setDefaultSectionSize(38)

    def cambiar_seccion(self, idx: int):
        nav_btns = (
            self.btn_nav_productos, self.btn_nav_proveedores,
            self.btn_nav_notas, self.btn_nav_reportes, self.btn_nav_clientes,
            self.btn_nav_venta, self.btn_nav_tickets
        )
        for b in nav_btns:
            b.setProperty("class", "nav_btn")
            b.style().unpolish(b)
            b.style().polish(b)

        if idx in (0, 1, 2):
            self._activar_btn(self.btn_nav_productos)
            if idx == 0:
                self.cargar_catalogo()
                self.txt_buscar_cat.setFocus()
        elif idx in (4, 5, 6):
            self._activar_btn(self.btn_nav_proveedores)
            if idx == 4:
                self.cargar_proveedores()
        elif idx in (7, 8, 9):
            self._activar_btn(self.btn_nav_notas)
            if idx == 7:
                self.cargar_notas_recibo()
        elif idx == 3:
            self._activar_btn(self.btn_nav_reportes)
            self.actualizar_buscador_proveedores_reporte()
            self.generar_todos_los_reportes()
        elif idx in (10, 11, 12):
            self._activar_btn(self.btn_nav_clientes)
            if idx == 10: self.cargar_clientes()
        elif idx == 13:
            self._activar_btn(self.btn_nav_venta)
            self._cargar_completers_venta()
        elif idx in (14, 15):
            self._activar_btn(self.btn_nav_tickets)
            if idx == 14:
                self.cargar_tickets()

        self.stacked.setCurrentIndex(idx)

    def _activar_btn(self, btn: QPushButton):
        btn.setProperty("class", "nav_btn_active")
        btn.style().unpolish(btn)
        btn.style().polish(btn)

    # =========================================================================
    # MÓDULO 1: PRODUCTOS
    # =========================================================================
    def setup_page_catalogo(self):
        layout = QVBoxLayout(self.page_catalogo)
        layout.setContentsMargins(24, 20, 24, 20)
        layout.setSpacing(12)

        top = QHBoxLayout()
        lbl = QLabel("Catálogo de productos")
        lbl.setObjectName("page_title")
        titulos = QVBoxLayout()
        titulos.setSpacing(3)
        titulos.addWidget(lbl)
        subtitulo = QLabel("Consulta y administra las existencias de tu inventario.")
        subtitulo.setStyleSheet("color:#64748B;font-size:12px;")
        titulos.addWidget(subtitulo)
        top.addLayout(titulos)
        top.addStretch()

        btn_add = QPushButton("+ Agregar producto")
        btn_add.setObjectName("btn_success")
        btn_add.clicked.connect(self.ir_agregar_producto)
        top.addWidget(btn_add)
        layout.addLayout(top)

        # Tarjetas con indicadores calculados a partir del catálogo activo.
        indicadores = QHBoxLayout()
        self.indicadores_cat = []
        for titulo, color in (("Productos activos", "#4F46E5"),
                              ("Requieren reorden", "#B45309"),
                              ("Valor a costo", "#4F46E5")):
            tarjeta = QFrame()
            tarjeta.setObjectName("metric_card")
            tarjeta.setStyleSheet(f"QFrame#metric_card {{background:white;border:1px solid #E4E8F7;border-left:3px solid {color};border-radius:8px;}}")
            tarjeta_l = QVBoxLayout(tarjeta)
            tarjeta_l.setContentsMargins(14, 7, 14, 7)
            nombre = QLabel(titulo)
            nombre.setStyleSheet("color:#64748B;font-size:12px;")
            valor = QLabel("—")
            valor.setStyleSheet(f"color:{color};font-size:22px;font-weight:bold;")
            tarjeta_l.addWidget(nombre)
            tarjeta_l.addWidget(valor)
            self.indicadores_cat.append(valor)
            indicadores.addWidget(tarjeta, 1)
        layout.addLayout(indicadores)

        self.banner_stock_catalogo = QFrame()
        self.banner_stock_catalogo.setStyleSheet(
            "background-color: #FFFBEB; border: 1px solid #FCD34D; border-radius: 6px;"
        )
        b_l = QHBoxLayout(self.banner_stock_catalogo)
        b_l.setContentsMargins(12, 8, 12, 8)
        self.lbl_banner_stock = QLabel()
        self.lbl_banner_stock.setWordWrap(True)
        self.lbl_banner_stock.setStyleSheet("color: #B45309; font-weight: 600; font-size: 12px; border: none;")
        b_l.addWidget(self.lbl_banner_stock)
        layout.addWidget(self.banner_stock_catalogo)

        filtros = QHBoxLayout()
        self.txt_buscar_cat = QLineEdit()
        self.txt_buscar_cat.setPlaceholderText("Buscar por nombre, código o categoría...")
        self.txt_buscar_cat.setClearButtonEnabled(True)
        self.txt_buscar_cat.returnPressed.connect(self.cargar_catalogo)

        lbl_cat_filtro = QLabel("Categoría:")
        lbl_cat_filtro.setStyleSheet("color: #0F172A; font-weight: bold; font-size: 13px;")

        self.cmb_filtro_cat = QComboBox()
        self.cmb_filtro_cat.addItems(["Todas", "Micas", "Accesorios", "Refacciones", "Recargas"])
        self.cmb_filtro_cat.currentTextChanged.connect(self.cargar_catalogo)

        self.chk_solo_bajo = QCheckBox("Solo stock bajo")
        self.chk_solo_bajo.stateChanged.connect(self.cargar_catalogo)

        self.chk_ver_inactivos = QCheckBox("Mostrar inactivos")
        self.chk_ver_inactivos.stateChanged.connect(self.cargar_catalogo)

        btn_buscar = QPushButton("Buscar")
        btn_buscar.setObjectName("btn_primary")
        btn_buscar.clicked.connect(self.cargar_catalogo)

        filtros.addWidget(self.txt_buscar_cat, stretch=2)
        filtros.addWidget(lbl_cat_filtro)
        filtros.addWidget(self.cmb_filtro_cat)
        filtros.addWidget(self.chk_solo_bajo)
        filtros.addWidget(self.chk_ver_inactivos)
        filtros.addWidget(btn_buscar)
        panel_filtros = QFrame()
        panel_filtros.setProperty("class", "card")
        panel_layout = QVBoxLayout(panel_filtros)
        panel_layout.setContentsMargins(12, 10, 12, 10)
        panel_layout.addLayout(filtros)
        layout.addWidget(panel_filtros)

        self.tabla_cat = QTableWidget()
        self.tabla_cat.setColumnCount(8)
        self.tabla_cat.setHorizontalHeaderLabels([
            "SKU / Código", "Nombre", "Categoría", "Unidad",
            "Costo Unit.", "Precio Venta", "Existencia", "Estado"
        ])
        self.tabla_cat.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.tabla_cat.verticalHeader().setVisible(False)
        self.tabla_cat.setEditTriggers(QTableWidget.NoEditTriggers)
        self.tabla_cat.setSelectionBehavior(QTableWidget.SelectRows)
        self.tabla_cat.setSelectionMode(QTableWidget.SingleSelection)
        # SKU completo (sin "...") y doble clic en una fila = Consultar
        self.tabla_cat.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.tabla_cat.doubleClicked.connect(lambda _i: self.consultar_producto_seleccionado())

        # Resumen corto que se actualiza en cargar_catalogo()
        self.lbl_resumen_cat = QLabel("")
        self.lbl_resumen_cat.setStyleSheet("color: #475569; font-size: 12px;")
        layout.addWidget(self.lbl_resumen_cat)
        layout.addWidget(self.tabla_cat)
        acciones=QHBoxLayout()
        ayuda = QLabel("Selecciona un producto para realizar una acción.")
        ayuda.setStyleSheet("color:#64748B;font-size:12px;")
        acciones.addWidget(ayuda)
        acciones.addStretch()
        for txt,fn,obj in [("Consultar",self.consultar_producto_seleccionado,"btn_secondary"),("Modificar",self.modificar_producto_seleccionado,"btn_primary"),("Eliminar",self.eliminar_producto_seleccionado,"btn_danger")]:
            bt=QPushButton(txt); bt.setObjectName(obj); bt.clicked.connect(fn); acciones.addWidget(bt)
        layout.addLayout(acciones)

    def cargar_catalogo(self):
        criterio = self.txt_buscar_cat.text().strip()
        cat = self.cmb_filtro_cat.currentText()
        solo_bajo = self.chk_solo_bajo.isChecked()
        inactivos = self.chk_ver_inactivos.isChecked()

        try:
            raw_prods = self.repo.buscar_productos(criterio, cat, solo_bajo, inactivos)
            productos = [normalizar_producto(p) for p in raw_prods]
            todos_activos = [normalizar_producto(p) for p in self.repo.buscar_productos()]
        except Exception as e:
            QMessageBox.critical(self, "Error de Base de Datos", f"Error al recuperar catálogo:\n{e}")
            return

        bajos = [p["nombre"] for p in todos_activos if p["existencia"] <= p["stockMinimo"]]
        # Actualizar los tres indicadores con los datos del inventario activo.
        valor_inventario = sum(p["costo"] * p["existencia"] for p in todos_activos)
        valores = (str(len(todos_activos)), str(len(bajos)), f"${valor_inventario:,.2f}")
        for etiqueta, valor in zip(self.indicadores_cat, valores):
            etiqueta.setText(valor)
        if bajos:
            self.banner_stock_catalogo.setVisible(True)
            self.lbl_banner_stock.setText(
                f"⚠ {len(bajos)} producto(s) están en o por debajo de su punto de reorden: " + ", ".join(bajos)
            )
        else:
            self.banner_stock_catalogo.setVisible(False)

        if productos:
            self.lbl_resumen_cat.setText(
                f"{len(productos)} producto(s) en la lista"
            )
        else:
            self.lbl_resumen_cat.setText("No se encontraron productos con esa búsqueda.")

        self.tabla_cat.setRowCount(len(productos))

        for row, p in enumerate(productos):
            self.tabla_cat.setRowHeight(row, 40)
            self.tabla_cat.setItem(row, 0, crear_celda(p["codigo"]))
            self.tabla_cat.setItem(row, 1, crear_celda(p["nombre"], negrita=True))
            self.tabla_cat.setItem(row, 2, crear_celda(p["categoria"]))
            self.tabla_cat.setItem(row, 3, crear_celda(p["unidadMedida"], Qt.AlignCenter))
            self.tabla_cat.setItem(row, 4, crear_celda(f"${p['costo']:,.2f}", Qt.AlignRight | Qt.AlignVCenter))
            self.tabla_cat.setItem(row, 5, crear_celda(f"${p['precio']:,.2f}", Qt.AlignRight | Qt.AlignVCenter))
            self.tabla_cat.setItem(row, 6, crear_celda(str(p["existencia"]), Qt.AlignCenter, negrita=True))

            if not p["estado"]:
                est_txt = "Inactivo"
            elif p["existencia"] == 0:
                est_txt = "Agotado"
            elif p["existencia"] <= p["stockMinimo"]:
                est_txt = "Stock bajo"
            else:
                est_txt = "Disponible"
            celda_estado = crear_celda(est_txt, Qt.AlignCenter, negrita=True)
            # Color del texto según el estado: verde = bien, naranja = poco, rojo = agotado
            colores_estado = {"Disponible": "#0D9488", "Stock bajo": "#D97706",
                              "Agotado": "#E11D48", "Inactivo": "#64748B"}
            celda_estado.setForeground(QBrush(QColor(colores_estado[est_txt])))
            self.tabla_cat.setItem(row, 7, celda_estado)

            # Toda la fila con un fondo suave: rojo si está agotado, amarillo si hay poco
            fondo_fila = {"Agotado": "#FFE4E6", "Stock bajo": "#FFFBEB"}.get(est_txt)
            if fondo_fila:
                for col in range(8):
                    self.tabla_cat.item(row, col).setBackground(QBrush(QColor(fondo_fila)))


    def setup_page_formulario(self):
        layout = QVBoxLayout(self.page_formulario)
        layout.setContentsMargins(28, 20, 28, 20)
        layout.setSpacing(14)

        top = QHBoxLayout()
        self.lbl_form_title = QLabel("Agregar producto")
        self.lbl_form_title.setStyleSheet("font-size: 20px; font-weight: bold; color: #0F172A;")
        top.addWidget(self.lbl_form_title)
        top.addStretch()
        layout.addLayout(top)

        self.lbl_restriccion_form = QLabel()
        self.lbl_restriccion_form.setStyleSheet(
            "background-color: #EEF2FF; color: #4338CA; border: 1px solid #A5B4FC; "
            "border-radius: 6px; padding: 8px 12px; font-size: 12px; font-weight: 600;"
        )
        layout.addWidget(self.lbl_restriccion_form)

        body = QHBoxLayout()
        body.setSpacing(18)

        card_left = QFrame()
        card_left.setProperty("class", "card")
        grid = QGridLayout(card_left)
        grid.setContentsMargins(20, 18, 20, 18)
        grid.setHorizontalSpacing(16)
        grid.setVerticalSpacing(10)

        grid.addWidget(self._lbl("Código / SKU *"), 0, 0)
        grid.addWidget(self._lbl("Nombre del producto *"), 0, 1)
        self.txt_codigo = QLineEdit()
        self.txt_codigo.setPlaceholderText("Código de barras o SKU único")
        self.txt_nombre = QLineEdit()
        self.txt_nombre.setPlaceholderText("Ingresa el nombre del producto")
        grid.addWidget(self.txt_codigo, 1, 0)
        grid.addWidget(self.txt_nombre, 1, 1)

        grid.addWidget(self._lbl("Categoría *"), 2, 0)
        grid.addWidget(self._lbl("Unidad de Medida *"), 2, 1)
        self.cmb_categoria = QComboBox()
        self.cmb_categoria.addItems(["Micas", "Accesorios", "Refacciones", "Recargas"])
        self.cmb_unidad = QComboBox()
        self.cmb_unidad.addItems(["pieza", "par", "kit", "metro", "caja"])
        grid.addWidget(self.cmb_categoria, 3, 0)
        grid.addWidget(self.cmb_unidad, 3, 1)

        grid.addWidget(self._lbl("Costo unitario * ($ Costo > 0)"), 4, 0)
        grid.addWidget(self._lbl("Precio de venta * ($ Precio > 0)"), 4, 1)
        self.spn_costo = QDoubleSpinBox()
        self.spn_costo.setPrefix("$ ")
        self.spn_costo.setRange(0.00, 999999.99)
        self.spn_costo.valueChanged.connect(self.recalcular_resumen_form)

        self.spn_precio = QDoubleSpinBox()
        self.spn_precio.setPrefix("$ ")
        self.spn_precio.setRange(0.00, 999999.99)
        self.spn_precio.valueChanged.connect(self.recalcular_resumen_form)
        grid.addWidget(self.spn_costo, 5, 0)
        grid.addWidget(self.spn_precio, 5, 1)

        grid.addWidget(self._lbl("Existencia *"), 6, 0)
        grid.addWidget(self._lbl("Punto de reorden * (> 0)"), 6, 1)
        self.spn_existencia = QSpinBox()
        self.spn_existencia.setRange(0, 99999)
        self.spn_existencia.valueChanged.connect(self.recalcular_resumen_form)

        self.spn_minimo = QSpinBox()
        self.spn_minimo.setRange(0, 99999)
        self.spn_minimo.valueChanged.connect(self.recalcular_resumen_form)
        grid.addWidget(self.spn_existencia, 7, 0)
        grid.addWidget(self.spn_minimo, 7, 1)

        grid.addWidget(self._lbl("Máximo en inventario *"), 8, 0)
        self.spn_maximo = QSpinBox()
        self.spn_maximo.setRange(0, 99999)
        grid.addWidget(self.spn_maximo, 9, 0)

        btns = QHBoxLayout()
        btns.addStretch()
        btn_cancel = QPushButton("Cancelar")
        btn_cancel.setObjectName("btn_secondary")
        btn_cancel.clicked.connect(lambda: self.cambiar_seccion(0))
        self.btn_guardar_form = QPushButton("Guardar")
        self.btn_guardar_form.setObjectName("btn_success")
        self.btn_guardar_form.clicked.connect(self.guardar_producto_bd)
        btns.addWidget(btn_cancel)
        btns.addWidget(self.btn_guardar_form)
        grid.addLayout(btns, 10, 0, 1, 2)

        body.addWidget(card_left, stretch=3)

        right_col = QVBoxLayout()
        card_res = QFrame()
        card_res.setProperty("class", "card")
        rl = QVBoxLayout(card_res)
        lbl_rc = QLabel("Resumen calculado")
        lbl_rc.setProperty("class", "card_title")
        rl.addWidget(lbl_rc)
        self.lbl_res_margen = QLabel("Margen unitario: $0.00")
        self.lbl_res_pct = QLabel("Margen %: 0.0 %")
        self.lbl_res_val_costo = QLabel("Valor a costo: $0.00")
        for l in (self.lbl_res_margen, self.lbl_res_pct, self.lbl_res_val_costo):
            l.setStyleSheet("font-weight: bold; color: #0F172A; font-size: 13px;")
            rl.addWidget(l)
        right_col.addWidget(card_res)

        card_al = QFrame()
        card_al.setProperty("class", "card")
        al = QVBoxLayout(card_al)
        lbl_as = QLabel("Estado de stock")
        lbl_as.setProperty("class", "card_title")
        al.addWidget(lbl_as)
        self.lbl_alerta_form = QLabel()
        self.lbl_alerta_form.setWordWrap(True)
        al.addWidget(self.lbl_alerta_form)
        right_col.addWidget(card_al)
        right_col.addStretch()

        body.addLayout(right_col, stretch=1)
        layout.addLayout(body)

    def ir_agregar_producto(self):
        self.modo_formulario_prod = "AGREGAR"
        self.producto_actual = None
        self.lbl_form_title.setText("Agregar producto")
        self.lbl_restriccion_form.setText(
            "Nota: Todo producto nuevo se registra con Existencia = 0. "
            "El Punto de Reorden, Costo y Precio deben ser números positivos mayores a 0."
        )

        self.txt_codigo.setReadOnly(False)
        self.txt_nombre.setReadOnly(False)
        self.txt_codigo.clear()
        self.txt_nombre.clear()
        self.cmb_categoria.setCurrentIndex(0)
        self.cmb_unidad.setCurrentIndex(0)
        self.spn_costo.setValue(0.00)
        self.spn_precio.setValue(0.00)
        self.spn_existencia.setValue(0)
        self.spn_existencia.setEnabled(False)
        self.spn_minimo.setValue(0)
        self.spn_maximo.setValue(0)
        self.recalcular_resumen_form()
        self.cambiar_seccion(1)

    def ir_modificar_producto(self, codigo: str):
        raw_prod = self.repo.obtener_producto(codigo)
        if not raw_prod:
            QMessageBox.warning(self, "No encontrado", "No se encontró el producto solicitado.")
            return

        prod = normalizar_producto(raw_prod)
        self.modo_formulario_prod = "MODIFICAR"
        self.producto_actual = prod
        self.lbl_form_title.setText("Modificar producto")
        self.lbl_restriccion_form.setText(
            "Nota: Se pueden modificar todos los datos del producto, excepto el nombre y el código."
        )

        self.txt_codigo.setText(prod["codigo"])
        self.txt_codigo.setReadOnly(True)
        self.txt_nombre.setText(prod["nombre"])
        self.txt_nombre.setReadOnly(True)

        self.cmb_categoria.setCurrentText(prod["categoria"])
        self.cmb_unidad.setCurrentText(prod["unidadMedida"])
        self.spn_costo.setValue(prod["costo"])
        self.spn_precio.setValue(prod["precio"])
        self.spn_existencia.setEnabled(True)
        self.spn_existencia.setValue(prod["existencia"])
        self.spn_minimo.setValue(prod["stockMinimo"])
        self.spn_maximo.setValue(prod["stockMaximo"])

        self.recalcular_resumen_form()
        self.cambiar_seccion(1)

    def recalcular_resumen_form(self):
        costo = self.spn_costo.value()
        precio = self.spn_precio.value()
        exist = self.spn_existencia.value()
        minimo = self.spn_minimo.value()

        margen = precio - costo
        pct = (margen / precio * 100.0) if precio > 0 else 0.0
        val_costo = costo * exist

        self.lbl_res_margen.setText(f"Margen unitario: ${margen:,.2f}")
        self.lbl_res_pct.setText(f"Margen %: {pct:,.1f} %")
        self.lbl_res_val_costo.setText(f"Valor a costo: ${val_costo:,.2f}")

        if exist <= minimo:
            self.lbl_alerta_form.setText(f"⚠ Existencia ({exist}) en o por debajo del punto de reorden ({minimo})")
            self.lbl_alerta_form.setStyleSheet("background:#FFFBEB; color:#B45309; padding:8px; border-radius:6px; font-weight:bold;")
        else:
            self.lbl_alerta_form.setText(f"✓ Stock por encima del punto de reorden ({minimo})")
            self.lbl_alerta_form.setStyleSheet("background:#ECFDF5; color:#047857; padding:8px; border-radius:6px; font-weight:bold;")

    def guardar_producto_bd(self):
        codigo = self.txt_codigo.text().strip()
        nombre = self.txt_nombre.text().strip()
        costo = self.spn_costo.value()
        precio = self.spn_precio.value()
        exist = self.spn_existencia.value()
        minimo = self.spn_minimo.value()
        maximo = self.spn_maximo.value()

        if not codigo or not nombre:
            QMessageBox.warning(self, "Datos incompletos", "El código y el nombre son obligatorios.")
            return
        if minimo <= 0 or costo <= 0 or precio <= 0:
            QMessageBox.warning(
                self, "Valores inválidos",
                "El punto de reorden, el costo y el precio deben ser números positivos mayores que 0."
            )
            return
        if maximo < minimo:
            QMessageBox.warning(
                self, "Máximo inválido",
                "El stock Máximo debe ser mayor o igual que el Punto de Reorden."
            )
            return

        datos = {
            "codigo": codigo,
            "nombre": nombre,
            "categoria": self.cmb_categoria.currentText(),
            "unidadMedida": self.cmb_unidad.currentText(),
            "existencia": 0 if self.modo_formulario_prod == "AGREGAR" else exist,
            "stockMinimo": minimo,
            "stockMaximo": maximo,
            "costo": costo,
            "precio": precio,
            "estado": True
        }

        try:
            if self.modo_formulario_prod == "AGREGAR":
                self.repo.agregar_producto(datos)
                QMessageBox.information(self, "Registro exitoso", "Producto registrado correctamente con existencia = 0.")
            else:
                self.repo.modificar_producto(codigo, datos)
                QMessageBox.information(self, "Actualización exitosa", "Los datos del producto se actualizaron correctamente.")
            self.cambiar_seccion(0)
        except Exception as e:
            QMessageBox.critical(self, "Error al guardar", str(e))

    def setup_page_consultar(self):
        layout = QVBoxLayout(self.page_consultar)
        layout.setContentsMargins(28, 20, 28, 20)
        layout.setSpacing(14)

        top = QHBoxLayout()
        lbl_t = QLabel("Consultar producto")
        lbl_t.setStyleSheet("font-size: 20px; font-weight: bold; color: #0F172A;")
        top.addWidget(lbl_t)
        top.addStretch()
        layout.addLayout(top)

        body = QHBoxLayout()
        body.setSpacing(18)

        card_det = QFrame()
        card_det.setProperty("class", "card")
        dl = QVBoxLayout(card_det)
        dl.setContentsMargins(20, 18, 20, 18)
        dl.setSpacing(10)

        self.lbl_ver_nombre = QLabel()
        self.lbl_ver_nombre.setStyleSheet("font-size: 17px; font-weight: bold; color: #0F172A;")
        dl.addWidget(self.lbl_ver_nombre)

        self.lbl_ver_alerta = QLabel()
        self.lbl_ver_alerta.setWordWrap(True)
        dl.addWidget(self.lbl_ver_alerta)

        self.bar_ver_stock = QProgressBar()
        self.bar_ver_stock.setFixedHeight(12)
        self.bar_ver_stock.setTextVisible(False)
        dl.addWidget(self.bar_ver_stock)

        self.grid_ver_datos = QGridLayout()
        self.grid_ver_datos.setVerticalSpacing(8)
        self.lbls_ver_valores = {}
        campos = [
            ("codigo", "SKU / Código:"),
            ("categoria", "Categoría:"),
            ("unidadMedida", "Unidad de Medida:"),
            ("precio", "Precio de venta:"),
            ("costo", "Costo unitario:"),
            ("existencia", "Existencia actual:"),
            ("stockMinimo", "Punto de reorden:"),
            ("stockMaximo", "Máximo permitido:"),
            ("estado", "Estado:")
        ]
        for idx, (key, titulo) in enumerate(campos):
            lk = QLabel(titulo)
            lk.setStyleSheet("color: #475569; font-size: 13px; font-weight: 600;")
            lv = QLabel("-")
            lv.setStyleSheet("color: #0F172A; font-size: 13px; font-weight: bold;")
            self.grid_ver_datos.addWidget(lk, idx, 0)
            self.grid_ver_datos.addWidget(lv, idx, 1)
            self.lbls_ver_valores[key] = lv

        dl.addLayout(self.grid_ver_datos)
        dl.addStretch()

        btns = QHBoxLayout()
        btns.addStretch()
        btn_volver = QPushButton("Volver")
        btn_volver.setObjectName("btn_secondary")
        btn_volver.clicked.connect(lambda: self.cambiar_seccion(0))
        self.btn_ver_ir_mod = QPushButton("Modificar")
        self.btn_ver_ir_mod.setObjectName("btn_primary")
        btns.addWidget(btn_volver)
        btns.addWidget(self.btn_ver_ir_mod)
        dl.addLayout(btns)

        body.addWidget(card_det, stretch=1)

        card_mov = QFrame()
        card_mov.setProperty("class", "card")
        ml = QVBoxLayout(card_mov)
        ml.setContentsMargins(20, 18, 20, 18)
        lbl_m = QLabel("Últimos movimientos")
        lbl_m.setProperty("class", "card_title")
        ml.addWidget(lbl_m)

        self.tabla_movs = QTableWidget()
        self.tabla_movs.setColumnCount(4)
        self.tabla_movs.setHorizontalHeaderLabels(["Fecha", "Tipo", "Documento", "Cantidad"])
        self.tabla_movs.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.tabla_movs.verticalHeader().setVisible(False)
        self.tabla_movs.setEditTriggers(QTableWidget.NoEditTriggers)
        ml.addWidget(self.tabla_movs)

        body.addWidget(card_mov, stretch=1)
        layout.addLayout(body)

    def ir_consultar_producto(self, codigo: str):
        raw_prod = self.repo.obtener_producto(codigo)
        if not raw_prod:
            QMessageBox.warning(self, "Sin coincidencias", "No se encontró el producto.")
            return

        prod = normalizar_producto(raw_prod)
        self.lbl_ver_nombre.setText(prod["nombre"])
        exist = prod["existencia"]
        minimo = prod["stockMinimo"]
        maximo = max(prod["stockMaximo"], minimo * 2, 1)

        self.bar_ver_stock.setRange(0, maximo)
        self.bar_ver_stock.setValue(min(exist, maximo))

        if exist <= minimo:
            self.lbl_ver_alerta.setText(
                f"⚠ Stock bajo: quedan {exist} unidades y el punto de reorden es {minimo}. Se recomienda reabastecer."
            )
            self.lbl_ver_alerta.setStyleSheet("background:#FFFBEB; color:#B45309; padding:10px; border-radius:6px; font-weight:bold;")
            self.bar_ver_stock.setStyleSheet("QProgressBar::chunk { background-color: #F59E0B; border-radius:4px; }")
        else:
            self.lbl_ver_alerta.setText(f"✓ Inventario disponible: {exist} unidades.")
            self.lbl_ver_alerta.setStyleSheet("background:#ECFDF5; color:#047857; padding:10px; border-radius:6px; font-weight:bold;")
            self.bar_ver_stock.setStyleSheet("QProgressBar::chunk { background-color: #10B981; border-radius:4px; }")

        self.lbls_ver_valores["codigo"].setText(prod["codigo"])
        self.lbls_ver_valores["categoria"].setText(prod["categoria"])
        self.lbls_ver_valores["unidadMedida"].setText(prod["unidadMedida"])
        self.lbls_ver_valores["precio"].setText(f"${prod['precio']:,.2f}")
        self.lbls_ver_valores["costo"].setText(f"${prod['costo']:,.2f}")
        self.lbls_ver_valores["existencia"].setText(str(exist))
        self.lbls_ver_valores["stockMinimo"].setText(str(minimo))
        self.lbls_ver_valores["stockMaximo"].setText(str(prod["stockMaximo"]))
        self.lbls_ver_valores["estado"].setText("Activo" if prod["estado"] else "Inactivo")

        try:
            self.btn_ver_ir_mod.clicked.disconnect()
        except Exception:
            pass
        self.btn_ver_ir_mod.clicked.connect(lambda: self.ir_modificar_producto(prod["codigo"]))

        movs = self.repo.obtener_movimientos_producto(codigo)
        self.tabla_movs.setRowCount(len(movs))
        for r, m in enumerate(movs):
            self.tabla_movs.setItem(r, 0, crear_celda(m.get("fecha_fmt", "")))
            self.tabla_movs.setItem(r, 1, crear_celda(m.get("tipo", "")))
            self.tabla_movs.setItem(r, 2, crear_celda(m.get("documento", "")))
            cant = int(m.get("cantidad", 0) or 0)
            cant_str = f"+{cant}" if cant > 0 else str(cant)
            self.tabla_movs.setItem(r, 3, crear_celda(cant_str, Qt.AlignCenter, negrita=True))

        self.cambiar_seccion(2)

    def abrir_eliminar_producto(self, prod: dict):
        dlg = DialogEliminarProducto(prod, self)
        if dlg.exec() == QDialog.Accepted:
            try:
                self.repo.eliminar_producto_logico(prod["codigo"])
                QMessageBox.information(
                    self, "Producto inactivado",
                    f"El producto '{prod['nombre']}' fue dado de baja lógicamente (Inactivo)."
                )
                self.cargar_catalogo()
            except Exception as e:
                QMessageBox.critical(self, "Restricción de eliminación", str(e))

    # =========================================================================
    # MÓDULO 2: PROVEEDORES (Con Validaciones de Teléfono/Correo y Baja Lógica)
    # =========================================================================
    def setup_page_proveedores_cat(self):
        layout = QVBoxLayout(self.page_proveedores_cat)
        layout.setContentsMargins(24, 20, 24, 20)
        layout.setSpacing(12)

        top = QHBoxLayout()
        lbl = QLabel("Proveedores")
        lbl.setStyleSheet("font-size: 20px; font-weight: bold; color: #0F172A;")
        top.addWidget(lbl)
        top.addStretch()

        btn_add = QPushButton("+ Agregar proveedor")
        btn_add.setObjectName("btn_success")
        btn_add.clicked.connect(self.ir_agregar_proveedor)
        top.addWidget(btn_add)
        layout.addLayout(top)

        filtros = QHBoxLayout()
        self.txt_buscar_prov = QLineEdit()
        self.txt_buscar_prov.setPlaceholderText("Buscar por nombre de empresa, contacto, RFC o ID...")
        self.txt_buscar_prov.returnPressed.connect(self.cargar_proveedores)

        self.chk_prov_inactivos = QCheckBox("Mostrar inactivos")
        self.chk_prov_inactivos.stateChanged.connect(self.cargar_proveedores)

        btn_buscar = QPushButton("Buscar")
        btn_buscar.setObjectName("btn_primary")
        btn_buscar.clicked.connect(self.cargar_proveedores)

        filtros.addWidget(self.txt_buscar_prov, stretch=1)
        filtros.addWidget(self.chk_prov_inactivos)
        filtros.addWidget(btn_buscar)
        layout.addLayout(filtros)

        self.tabla_prov = QTableWidget()
        self.tabla_prov.setColumnCount(7)
        self.tabla_prov.setHorizontalHeaderLabels([
            "ID", "RFC", "Nombre de la empresa", "Contacto", "Teléfono", "Correo", "Estado"
        ])
        self.tabla_prov.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.tabla_prov.verticalHeader().setVisible(False)
        self.tabla_prov.setEditTriggers(QTableWidget.NoEditTriggers)
        self.tabla_prov.setSelectionBehavior(QTableWidget.SelectRows)
        self.tabla_prov.setSelectionMode(QTableWidget.SingleSelection)
        layout.addWidget(self.tabla_prov)
        acciones=QHBoxLayout()
        for txt,fn,obj in [("Consultar",self.consultar_proveedor_seleccionado,"btn_secondary"),("Modificar",self.modificar_proveedor_seleccionado,"btn_primary"),("Eliminar",self.eliminar_proveedor_seleccionado,"btn_danger")]:
            bt=QPushButton(txt); bt.setObjectName(obj); bt.clicked.connect(fn); acciones.addWidget(bt)
        acciones.addStretch(); layout.addLayout(acciones)

    def cargar_proveedores(self):
        criterio = self.txt_buscar_prov.text().strip()
        inactivos = self.chk_prov_inactivos.isChecked()
        try:
            raw_provs = self.repo.buscar_proveedores(criterio, incluir_inactivos=inactivos)
            proveedores = [normalizar_proveedor(p) for p in raw_provs]
        except Exception as e:
            QMessageBox.critical(self, "Error de BD", f"No se pudieron cargar los proveedores:\n{e}")
            return

        self.tabla_prov.setRowCount(len(proveedores))
        if not proveedores and criterio:
            QMessageBox.information(self, "Búsqueda", "No se encontraron resultados.")

        for row, pr in enumerate(proveedores):
            self.tabla_prov.setRowHeight(row, 44)
            self.tabla_prov.setItem(row, 0, crear_celda(str(pr["idProveedor"]), Qt.AlignCenter))
            self.tabla_prov.setItem(row, 1, crear_celda(pr["rfc"], negrita=True))
            self.tabla_prov.setItem(row, 2, crear_celda(pr["empresa"], negrita=True))
            self.tabla_prov.setItem(row, 3, crear_celda(pr["nombre"]))
            self.tabla_prov.setItem(row, 4, crear_celda(str(pr["telefono"])))
            self.tabla_prov.setItem(row, 5, crear_celda(pr["correoE"]))
            est_str = "Activo" if pr["estado"] else "Inactivo"
            self.tabla_prov.setItem(row, 6, crear_celda(est_str, Qt.AlignCenter, negrita=True))


    def setup_page_proveedores_form(self):
        layout = QVBoxLayout(self.page_proveedores_form)
        layout.setContentsMargins(28, 20, 28, 20)
        layout.setSpacing(14)

        self.lbl_prov_form_title = QLabel("Agregar proveedor")
        self.lbl_prov_form_title.setStyleSheet("font-size: 20px; font-weight: bold; color: #0F172A;")
        layout.addWidget(self.lbl_prov_form_title)

        card = QFrame()
        card.setProperty("class", "card")
        grid = QGridLayout(card)
        grid.setContentsMargins(24, 20, 24, 20)
        grid.setHorizontalSpacing(18)
        grid.setVerticalSpacing(10)

        grid.addWidget(self._lbl("ID del proveedor (Autoincrementable)"), 0, 0)
        grid.addWidget(self._lbl("RFC * (12 o 13 caracteres, único)"), 0, 1)
        self.txt_prov_id = QLineEdit()
        self.txt_prov_id.setReadOnly(True)
        self.txt_prov_rfc = QLineEdit()
        self.txt_prov_rfc.setPlaceholderText("Ej. GAC120315KJ8")
        grid.addWidget(self.txt_prov_id, 1, 0)
        grid.addWidget(self.txt_prov_rfc, 1, 1)

        grid.addWidget(self._lbl("Nombre de la empresa / Razón social *"), 2, 0)
        grid.addWidget(self._lbl("Nombre del contacto *"), 2, 1)
        self.txt_prov_empresa = QLineEdit()
        self.txt_prov_empresa.setPlaceholderText("Nombre comercial o razón social")
        self.txt_prov_contacto = QLineEdit()
        self.txt_prov_contacto.setPlaceholderText("Persona de contacto")
        grid.addWidget(self.txt_prov_empresa, 3, 0)
        grid.addWidget(self.txt_prov_contacto, 3, 1)

        grid.addWidget(self._lbl("Teléfono * (Solo números, máx. 15 dígitos)"), 4, 0)
        grid.addWidget(self._lbl("Correo electrónico * (Debe incluir '@')"), 4, 1)
        self.txt_prov_tel = QLineEdit()
        self.txt_prov_tel.setPlaceholderText("Solo números, ej. 4494186620")
        # Validador en vivo: impide escribir letras o caracteres no numéricos, mantendo el valor como string
        regex_solo_numeros = QRegularExpression(r"^\d{0,15}$")
        self.txt_prov_tel.setValidator(QRegularExpressionValidator(regex_solo_numeros, self.txt_prov_tel))

        self.txt_prov_correo = QLineEdit()
        self.txt_prov_correo.setPlaceholderText("compras@proveedor.mx")
        grid.addWidget(self.txt_prov_tel, 5, 0)
        grid.addWidget(self.txt_prov_correo, 5, 1)

        grid.addWidget(self._lbl("Dirección / Domicilio *"), 6, 0)
        grid.addWidget(self._lbl("Página Web (opcional)"), 6, 1)
        self.txt_prov_domicilio = QLineEdit()
        self.txt_prov_domicilio.setPlaceholderText("Calle, número, colonia y ciudad")
        self.txt_prov_web = QLineEdit()
        self.txt_prov_web.setPlaceholderText("www.proveedor.mx")
        grid.addWidget(self.txt_prov_domicilio, 7, 0)
        grid.addWidget(self.txt_prov_web, 7, 1)

        btns = QHBoxLayout()
        btns.addStretch()
        btn_cancel = QPushButton("Cancelar")
        btn_cancel.setObjectName("btn_secondary")
        btn_cancel.clicked.connect(lambda: self.cambiar_seccion(4))

        self.btn_guardar_prov = QPushButton("Guardar")
        self.btn_guardar_prov.setObjectName("btn_primary")
        self.btn_guardar_prov.clicked.connect(self.guardar_proveedor_bd)

        btns.addWidget(btn_cancel)
        btns.addWidget(self.btn_guardar_prov)
        grid.addLayout(btns, 8, 0, 1, 2)

        layout.addWidget(card)
        layout.addStretch()

    def ir_agregar_proveedor(self):
        self.modo_formulario_prov = "AGREGAR"
        self.proveedor_actual = None
        self.lbl_prov_form_title.setText("Agregar proveedor")
        self.btn_guardar_prov.setText("Guardar")
        self.txt_prov_id.setText("(nuevo)")
        self.txt_prov_rfc.clear()
        self.txt_prov_empresa.clear()
        self.txt_prov_contacto.clear()
        self.txt_prov_tel.clear()
        self.txt_prov_correo.clear()
        self.txt_prov_domicilio.clear()
        self.txt_prov_web.clear()
        self.cambiar_seccion(5)

    def ir_modificar_proveedor(self, id_proveedor: int):
        raw_pr = self.repo.obtener_proveedor(id_proveedor)
        if not raw_pr:
            QMessageBox.warning(self, "No encontrado", "No se encontró el proveedor seleccionado.")
            return
        pr = normalizar_proveedor(raw_pr)
        self.modo_formulario_prov = "MODIFICAR"
        self.proveedor_actual = pr
        self.lbl_prov_form_title.setText(f"Modificar proveedor · ID {pr['idProveedor']}")
        self.btn_guardar_prov.setText("Guardar cambios")

        self.txt_prov_id.setText(str(pr["idProveedor"]))
        self.txt_prov_rfc.setText(pr["rfc"])
        self.txt_prov_empresa.setText(pr["empresa"])
        self.txt_prov_contacto.setText(pr["nombre"])
        # Limpiar cualquier espacio previo para que el validador numérico funcione limpio
        tel_limpio = "".join(ch for ch in str(pr["telefono"]) if ch.isdigit())
        self.txt_prov_tel.setText(tel_limpio)
        self.txt_prov_correo.setText(pr["correoE"])
        self.txt_prov_domicilio.setText(pr["domicilio"])
        self.txt_prov_web.setText(pr["paginaWeb"])
        self.cambiar_seccion(5)

    def guardar_proveedor_bd(self):
        rfc = self.txt_prov_rfc.text().strip().upper()
        empresa = self.txt_prov_empresa.text().strip()
        contacto = self.txt_prov_contacto.text().strip()
        tel_str = str(self.txt_prov_tel.text().strip())  # Interpretado siempre como string
        correo_str = self.txt_prov_correo.text().strip()
        domicilio = self.txt_prov_domicilio.text().strip()

        if not rfc or not empresa or not contacto or not tel_str or not correo_str or not domicilio:
            QMessageBox.warning(
                self, "Campos obligatorios",
                "Por favor completa RFC, Nombre de la empresa, Contacto, Teléfono, Correo y Dirección."
            )
            return

        if not tel_str.isdigit() or len(tel_str) < 7:
            QMessageBox.warning(
                self, "Teléfono inválido",
                "El número de teléfono solo debe contener dígitos numéricos (mínimo 7 dígitos)."
            )
            return

        if "@" not in correo_str or correo_str.startswith("@") or correo_str.endswith("@"):
            QMessageBox.warning(
                self, "Correo electrónico inválido",
                "El correo electrónico ingresado no es válido: debe contener el símbolo '@' (ejemplo: usuario@dominio.com)."
            )
            return

        datos = {
            "rfc": rfc,
            "empresa": empresa,
            "nombre": contacto,
            "telefono": tel_str,
            "correoE": correo_str,
            "domicilio": domicilio,
            "paginaWeb": self.txt_prov_web.text().strip()
        }

        try:
            if self.modo_formulario_prov == "AGREGAR":
                self.repo.agregar_proveedor(datos)
                QMessageBox.information(self, "Proveedor registrado", "El proveedor se registró correctamente.")
            else:
                self.repo.modificar_proveedor(self.proveedor_actual["idProveedor"], datos)
                QMessageBox.information(self, "Proveedor actualizado", "Los datos del proveedor fueron actualizados.")
            self.cambiar_seccion(4)
        except Exception as e:
            QMessageBox.critical(self, "Error al guardar proveedor", str(e))

    def setup_page_proveedores_ver(self):
        layout = QVBoxLayout(self.page_proveedores_ver)
        layout.setContentsMargins(28, 20, 28, 20)
        layout.setSpacing(14)

        lbl_t = QLabel("Consultar proveedor")
        lbl_t.setStyleSheet("font-size: 20px; font-weight: bold; color: #0F172A;")
        layout.addWidget(lbl_t)

        body = QHBoxLayout()
        body.setSpacing(18)

        card_det = QFrame()
        card_det.setProperty("class", "card")
        dl = QVBoxLayout(card_det)
        dl.setContentsMargins(20, 18, 20, 18)
        dl.setSpacing(10)

        lbl_sub = QLabel("Datos del proveedor (Solo lectura)")
        lbl_sub.setProperty("class", "card_title")
        dl.addWidget(lbl_sub)

        grid = QGridLayout()
        grid.setVerticalSpacing(8)
        self.lbls_ver_prov = {}
        campos_prov = [
            ("idProveedor", "ID Proveedor:"),
            ("rfc", "RFC:"),
            ("empresa", "Nombre de la empresa:"),
            ("nombre", "Nombre del contacto:"),
            ("telefono", "Teléfono:"),
            ("correoE", "Correo electrónico:"),
            ("domicilio", "Dirección / Domicilio:"),
            ("paginaWeb", "Página Web:"),
            ("estado", "Estado:")
        ]
        for idx, (k, tit) in enumerate(campos_prov):
            lk = QLabel(tit)
            lk.setStyleSheet("color: #475569; font-size: 13px; font-weight: 600;")
            lv = QLabel("-")
            lv.setWordWrap(True)
            lv.setStyleSheet("color: #0F172A; font-size: 13px; font-weight: bold;")
            grid.addWidget(lk, idx, 0)
            grid.addWidget(lv, idx, 1)
            self.lbls_ver_prov[k] = lv

        dl.addLayout(grid)
        dl.addStretch()

        btns = QHBoxLayout()
        btns.addStretch()
        btn_volver = QPushButton("Volver")
        btn_volver.setObjectName("btn_secondary")
        btn_volver.clicked.connect(lambda: self.cambiar_seccion(4))
        self.btn_prov_ver_mod = QPushButton("Modificar")
        self.btn_prov_ver_mod.setObjectName("btn_primary")
        btns.addWidget(btn_volver)
        btns.addWidget(self.btn_prov_ver_mod)
        dl.addLayout(btns)

        body.addWidget(card_det, stretch=1)

        card_notas = QFrame()
        card_notas.setProperty("class", "card")
        nl = QVBoxLayout(card_notas)
        nl.setContentsMargins(20, 18, 20, 18)
        lbl_n = QLabel("Notas de recibo recientes del proveedor")
        lbl_n.setProperty("class", "card_title")
        nl.addWidget(lbl_n)

        self.tabla_prov_notas = QTableWidget()
        self.tabla_prov_notas.setColumnCount(5)
        self.tabla_prov_notas.setHorizontalHeaderLabels(["Folio", "Fecha", "Partidas", "Unidades", "Costo total"])
        self.tabla_prov_notas.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.tabla_prov_notas.verticalHeader().setVisible(False)
        self.tabla_prov_notas.setEditTriggers(QTableWidget.NoEditTriggers)
        nl.addWidget(self.tabla_prov_notas)

        body.addWidget(card_notas, stretch=1)
        layout.addLayout(body)

    def ir_consultar_proveedor(self, id_proveedor: int):
        raw_pr = self.repo.obtener_proveedor(id_proveedor)
        if not raw_pr:
            QMessageBox.warning(self, "No encontrado", "El proveedor no existe.")
            return
        pr = normalizar_proveedor(raw_pr)

        self.lbls_ver_prov["idProveedor"].setText(str(pr["idProveedor"]))
        self.lbls_ver_prov["rfc"].setText(pr["rfc"])
        self.lbls_ver_prov["empresa"].setText(pr["empresa"])
        self.lbls_ver_prov["nombre"].setText(pr["nombre"])
        self.lbls_ver_prov["telefono"].setText(str(pr["telefono"]))
        self.lbls_ver_prov["correoE"].setText(pr["correoE"] or "No registrado")
        self.lbls_ver_prov["domicilio"].setText(pr["domicilio"])
        self.lbls_ver_prov["paginaWeb"].setText(pr["paginaWeb"] or "No registrada")
        self.lbls_ver_prov["estado"].setText("Activo" if pr["estado"] else "Inactivo")

        try:
            self.btn_prov_ver_mod.clicked.disconnect()
        except Exception:
            pass
        self.btn_prov_ver_mod.clicked.connect(lambda: self.ir_modificar_proveedor(pr["idProveedor"]))

        notas = self.repo.obtener_notas_por_proveedor(id_proveedor)
        self.tabla_prov_notas.setRowCount(len(notas))
        for r, n in enumerate(notas):
            self.tabla_prov_notas.setItem(r, 0, crear_celda(f"NR-{int(n['folio']):04d}", negrita=True))
            self.tabla_prov_notas.setItem(r, 1, crear_celda(str(n["fecha"]), Qt.AlignCenter))
            self.tabla_prov_notas.setItem(r, 2, crear_celda(str(n["partidas"]), Qt.AlignCenter))
            self.tabla_prov_notas.setItem(r, 3, crear_celda(str(n["unidades"]), Qt.AlignCenter))
            self.tabla_prov_notas.setItem(r, 4, crear_celda(f"${float(n['costo_total']):,.2f}", Qt.AlignRight | Qt.AlignVCenter, negrita=True))

        self.cambiar_seccion(6)

    def abrir_eliminar_proveedor(self, prov: dict):
        dlg = DialogEliminarProveedor(prov, self)
        if dlg.exec() == QDialog.Accepted:
            try:
                self.repo.eliminar_proveedor_logico(prov["idProveedor"])
                QMessageBox.information(
                    self, "Proveedor dado de baja",
                    f"El proveedor '{prov['empresa']}' fue desactivado mediante eliminación lógica (Inactivo)."
                )
                self.cargar_proveedores()
            except Exception as e:
                QMessageBox.critical(self, "No se pudo eliminar", str(e))

    # =========================================================================
    # MÓDULO 3: NOTAS DE RECIBO (Con Buscadores de Proveedor y Producto)
    # =========================================================================
    def setup_page_notas_cat(self):
        layout = QVBoxLayout(self.page_notas_cat)
        layout.setContentsMargins(24, 20, 24, 20)
        layout.setSpacing(12)

        top = QHBoxLayout()
        lbl = QLabel("Notas de recibo de mercancía")
        lbl.setStyleSheet("font-size: 20px; font-weight: bold; color: #0F172A;")
        top.addWidget(lbl)
        top.addStretch()

        btn_nueva = QPushButton("+ Registrar nota de recibo")
        btn_nueva.setObjectName("btn_success")
        btn_nueva.clicked.connect(self.ir_agregar_nota_recibo)
        top.addWidget(btn_nueva)
        layout.addLayout(top)

        filtros = QHBoxLayout()
        self.txt_buscar_nota = QLineEdit()
        self.txt_buscar_nota.setPlaceholderText("Buscar por folio, proveedor o estado (Activa / Cancelada)...")
        self.txt_buscar_nota.returnPressed.connect(self.cargar_notas_recibo)

        btn_buscar = QPushButton("Buscar")
        btn_buscar.setObjectName("btn_primary")
        btn_buscar.clicked.connect(self.cargar_notas_recibo)

        filtros.addWidget(self.txt_buscar_nota, stretch=1)
        filtros.addWidget(btn_buscar)
        layout.addLayout(filtros)

        self.tabla_notas = QTableWidget()
        self.tabla_notas.setColumnCount(7)
        self.tabla_notas.setHorizontalHeaderLabels([
            "Folio", "Fecha", "Proveedor que surte", "Partidas", "Unidades", "Costo Total", "Estado"
        ])
        self.tabla_notas.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.tabla_notas.horizontalHeader().setSectionResizeMode(6, QHeaderView.Fixed)
        self.tabla_notas.setColumnWidth(6, 120)
        self.tabla_notas.verticalHeader().setVisible(False)
        self.tabla_notas.setEditTriggers(QTableWidget.NoEditTriggers)
        self.tabla_notas.setSelectionBehavior(QTableWidget.SelectRows)
        self.tabla_notas.setSelectionMode(QTableWidget.SingleSelection)
        layout.addWidget(self.tabla_notas)
        acciones=QHBoxLayout()
        for txt,fn,obj in [("Consultar",self.consultar_nota_seleccionada,"btn_secondary"),("Modificar",self.modificar_nota_seleccionada,"btn_primary"),("Cancelar / Eliminar",self.cancelar_nota_seleccionada,"btn_danger")]:
            bt=QPushButton(txt); bt.setObjectName(obj); bt.clicked.connect(fn); acciones.addWidget(bt)
        acciones.addStretch(); layout.addLayout(acciones)

    def cargar_notas_recibo(self):
        filtro = self.txt_buscar_nota.text().strip()
        try:
            notas = self.repo.listar_notas_recibo(filtro)
        except Exception as e:
            QMessageBox.critical(self, "Error de BD", f"Error al cargar notas de recibo:\n{e}")
            return

        self.tabla_notas.setRowCount(len(notas))
        for r, n in enumerate(notas):
            self.tabla_notas.setRowHeight(r, 44)
            folio = int(n["folio"])
            estado = str(n.get("estado", "Activa"))
            self.tabla_notas.setItem(r, 0, crear_celda(f"NR-{folio:04d}", negrita=True))
            self.tabla_notas.setItem(r, 1, crear_celda(str(n["fecha_fmt"]), Qt.AlignCenter))
            self.tabla_notas.setItem(r, 2, crear_celda(str(n["proveedor"])))
            self.tabla_notas.setItem(r, 3, crear_celda(str(n["partidas"]), Qt.AlignCenter))
            self.tabla_notas.setItem(r, 4, crear_celda(str(n["unidades"]), Qt.AlignCenter))
            self.tabla_notas.setItem(r, 5, crear_celda(f"${float(n['total']):,.2f}", Qt.AlignRight | Qt.AlignVCenter, negrita=True))
            self.tabla_notas.setItem(r, 6, crear_celda(estado))


    def setup_page_notas_form(self):
        layout = QVBoxLayout(self.page_notas_form)
        layout.setContentsMargins(24, 18, 24, 18)
        layout.setSpacing(12)

        self.lbl_nota_form_title = QLabel("Registrar nota de recibo")
        self.lbl_nota_form_title.setStyleSheet("font-size: 20px; font-weight: bold; color: #0F172A;")
        layout.addWidget(self.lbl_nota_form_title)

        self.lbl_nota_aviso = QLabel()
        self.lbl_nota_aviso.setStyleSheet(
            "background-color: #FFFBEB; color: #B45309; border: 1px solid #FCD34D; "
            "border-radius: 6px; padding: 8px 12px; font-size: 12px; font-weight: bold;"
        )
        layout.addWidget(self.lbl_nota_aviso)

        # Cabecera de la nota con Buscador de Proveedor
        card_cab = QFrame()
        card_cab.setProperty("class", "card")
        cab_grid = QGridLayout(card_cab)
        cab_grid.setContentsMargins(16, 12, 16, 12)
        cab_grid.setHorizontalSpacing(16)

        cab_grid.addWidget(self._lbl("Folio de la nota"), 0, 0)
        cab_grid.addWidget(self._lbl("Fecha de recepción"), 0, 1)
        cab_grid.addWidget(self._lbl("Buscar proveedor que surte * (Nombre, RFC o ID)"), 0, 2)
        cab_grid.addWidget(self._lbl("Estado de la nota"), 0, 3)

        self.spn_nota_folio = QSpinBox()
        self.spn_nota_folio.setRange(1, 999999)
        self.spn_nota_folio.setPrefix("NR-")

        self.dt_nota_fecha = QDateEdit(QDate.currentDate())
        self.dt_nota_fecha.setDisplayFormat("dd/MM/yyyy")
        configurar_date_edit_claro(self.dt_nota_fecha)

        self.txt_nota_buscar_prov = QLineEdit()
        self.txt_nota_buscar_prov.setPlaceholderText("Escribe nombre de la empresa o RFC del proveedor...")

        self.cmb_nota_estado = QComboBox()
        self.cmb_nota_estado.addItems(["Activa", "Cancelada"])

        cab_grid.addWidget(self.spn_nota_folio, 1, 0)
        cab_grid.addWidget(self.dt_nota_fecha, 1, 1)
        cab_grid.addWidget(self.txt_nota_buscar_prov, 1, 2)
        cab_grid.addWidget(self.cmb_nota_estado, 1, 3)
        cab_grid.setColumnStretch(2, 2)
        layout.addWidget(card_cab)

        # Panel para buscar y agregar productos registrados a la nota (Tal como en el mockup)
        card_det = QFrame()
        card_det.setProperty("class", "card")
        det_l = QVBoxLayout(card_det)
        det_l.setContentsMargins(16, 14, 16, 14)
        det_l.setSpacing(10)

        lbl_pr = QLabel("Productos recibidos")
        lbl_pr.setProperty("class", "card_title")
        det_l.addWidget(lbl_pr)

        add_row = QHBoxLayout()
        self.txt_nota_buscar_prod = QLineEdit()
        self.txt_nota_buscar_prod.setPlaceholderText("🔍 Código o nombre del producto...")
        self.txt_nota_buscar_prod.textChanged.connect(self._al_cambiar_texto_buscador_prod_nota)
        self.txt_nota_buscar_prod.returnPressed.connect(self.agregar_renglon_a_nota)

        self.spn_nota_cant = QSpinBox()
        self.spn_nota_cant.setRange(1, 99999)
        self.spn_nota_cant.setValue(1)

        self.spn_nota_costo = QDoubleSpinBox()
        self.spn_nota_costo.setPrefix("$ ")
        self.spn_nota_costo.setRange(0.01, 999999.99)
        self.spn_nota_costo.setDecimals(2)

        btn_agregar_prod_nota = QPushButton("Agregar")
        btn_agregar_prod_nota.setObjectName("btn_primary")
        btn_agregar_prod_nota.clicked.connect(self.agregar_renglon_a_nota)

        add_row.addWidget(self.txt_nota_buscar_prod, stretch=3)
        add_row.addWidget(self._lbl("Cant.:"))
        add_row.addWidget(self.spn_nota_cant)
        add_row.addWidget(self._lbl("$ Costo unitario:"))
        add_row.addWidget(self.spn_nota_costo)
        add_row.addWidget(btn_agregar_prod_nota)
        det_l.addLayout(add_row)

        self.tabla_nota_det = QTableWidget()
        self.tabla_nota_det.setColumnCount(7)
        self.tabla_nota_det.setHorizontalHeaderLabels([
            "Cantidad", "Nombre del Producto", "$ Costo", "Subtotal",
            "Stock antes", "Stock resultante", "Acción"
        ])
        self.tabla_nota_det.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.tabla_nota_det.verticalHeader().setVisible(False)
        self.tabla_nota_det.itemChanged.connect(self._al_editar_celda_detalle_nota)
        det_l.addWidget(self.tabla_nota_det)

        bottom_bar = QHBoxLayout()
        self.lbl_nota_total = QLabel("Total de la nota: $0.00")
        self.lbl_nota_total.setStyleSheet("font-size: 16px; font-weight: bold; color: #0F172A;")
        bottom_bar.addWidget(self.lbl_nota_total)
        bottom_bar.addStretch()

        btn_cancel = QPushButton("Cancelar")
        btn_cancel.setObjectName("btn_secondary")
        btn_cancel.clicked.connect(lambda: self.cambiar_seccion(7))

        self.btn_guardar_nota = QPushButton("Registrar nota")
        self.btn_guardar_nota.setObjectName("btn_success")
        self.btn_guardar_nota.clicked.connect(self.guardar_nota_recibo_bd)

        bottom_bar.addWidget(btn_cancel)
        bottom_bar.addWidget(self.btn_guardar_nota)
        det_l.addLayout(bottom_bar)

        layout.addWidget(card_det)

    def _cargar_buscadores_nota(self):
        # 1. Configurar QCompleter para el buscador de proveedores en la nota
        self.mapa_busqueda_provs_nota.clear()
        provs = [normalizar_proveedor(p) for p in self.repo.buscar_proveedores(incluir_inactivos=False)]
        opciones_prov = []
        for p in provs:
            etiqueta = f"{p['empresa']} · RFC: {p['rfc']} [ID: {p['idProveedor']}]"
            self.mapa_busqueda_provs_nota[etiqueta] = p["idProveedor"]
            opciones_prov.append(etiqueta)

        comp_prov = QCompleter(opciones_prov, self)
        comp_prov.setCaseSensitivity(Qt.CaseInsensitive)
        comp_prov.setFilterMode(Qt.MatchContains)
        self.txt_nota_buscar_prov.setCompleter(comp_prov)

        # 2. Configurar QCompleter para el buscador de productos en la nota
        self.mapa_busqueda_prods_nota.clear()
        prods = [normalizar_producto(pr) for pr in self.repo.buscar_productos(incluir_inactivos=False)]
        opciones_prod = []
        for pr in prods:
            etiqueta = f"{pr['codigo']} · {pr['nombre']} (Stock actual: {pr['existencia']})"
            self.mapa_busqueda_prods_nota[etiqueta] = pr
            opciones_prod.append(etiqueta)

        comp_prod = QCompleter(opciones_prod, self)
        comp_prod.setCaseSensitivity(Qt.CaseInsensitive)
        comp_prod.setFilterMode(Qt.MatchContains)
        comp_prod.activated.connect(self._al_cambiar_texto_buscador_prod_nota)
        self.txt_nota_buscar_prod.setCompleter(comp_prod)

    def _resolver_producto_buscado(self, texto: str) -> dict | None:
        t = texto.strip()
        if not t:
            return None
        if t in self.mapa_busqueda_prods_nota:
            return self.mapa_busqueda_prods_nota[t]
        t_low = t.lower()
        for etiqueta, prod in self.mapa_busqueda_prods_nota.items():
            if prod["codigo"].lower() == t_low or prod["nombre"].lower() == t_low or t_low in etiqueta.lower():
                return prod
        return None

    def _resolver_proveedor_buscado(self, texto: str, mapa: dict) -> int | None:
        t = texto.strip()
        if not t:
            return None
        if t in mapa:
            return mapa[t]
        t_low = t.lower()
        for etiqueta, id_p in mapa.items():
            if t_low in etiqueta.lower():
                return id_p
        return None

    def _al_cambiar_texto_buscador_prod_nota(self, texto: str = ""):
        prod = self._resolver_producto_buscado(self.txt_nota_buscar_prod.text())
        if prod:
            self.spn_nota_costo.setValue(float(prod.get("costo", 1.0)))

    def ir_agregar_nota_recibo(self):
        self._cargar_buscadores_nota()
        if not self.mapa_busqueda_provs_nota:
            QMessageBox.warning(self, "Sin proveedores activos", "Primero debes tener al menos un proveedor activo registrado.")
            return

        self.modo_formulario_nota = "AGREGAR"
        self.folio_nota_editando = None
        self.renglones_nota_actual = []

        self.lbl_nota_form_title.setText("Registrar nota de recibo")
        self.lbl_nota_aviso.setText(
            "Al registrar la nota se sumará la Cantidad al stock en inventario de cada producto "
            "y el $ Costo recibido reemplazará el costo actual del producto."
        )

        sig_folio = self.repo.obtener_siguiente_folio_nota()
        self.spn_nota_folio.setValue(sig_folio)
        self.spn_nota_folio.setEnabled(True)
        self.dt_nota_fecha.setDate(QDate.currentDate())
        self.dt_nota_fecha.setEnabled(True)

        self.txt_nota_buscar_prov.clear()
        self.txt_nota_buscar_prov.setReadOnly(False)
        self.txt_nota_buscar_prod.clear()

        self.cmb_nota_estado.setCurrentText("Activa")
        self.cmb_nota_estado.setEnabled(False)

        self.btn_guardar_nota.setText("Registrar nota")
        self._refrescar_tabla_renglones_nota()
        self.cambiar_seccion(8)

    def ir_modificar_nota_recibo(self, folio: int):
        nota = self.repo.obtener_nota_completa(folio)
        if not nota:
            QMessageBox.warning(self, "No encontrada", "No se encontró la nota de recibo.")
            return
        if nota["estado"] != "Activa":
            QMessageBox.warning(self, "No editable", "Solo se pueden modificar notas con estado 'Activa'.")
            return

        self._cargar_buscadores_nota()
        self.modo_formulario_nota = "MODIFICAR"
        self.folio_nota_editando = folio

        self.lbl_nota_form_title.setText(f"Modificar nota de recibo · NR-{folio:04d}")
        self.lbl_nota_aviso.setText(
            "Restricción de modificación: Solo se pueden modificar los datos de los productos recibidos "
            "(cantidades, costos o quitar/agregar productos). El inventario y el costo del producto se actualizarán al guardar."
        )

        self.spn_nota_folio.setValue(folio)
        self.spn_nota_folio.setEnabled(False)

        partes = str(nota["fecha_raw"]).split("-")
        if len(partes) == 3:
            self.dt_nota_fecha.setDate(QDate(int(partes[0]), int(partes[1]), int(partes[2])))
        self.dt_nota_fecha.setEnabled(False)

        self.txt_nota_buscar_prov.setText(f"{nota['proveedor']} · RFC: {nota['rfc_proveedor']} [ID: {nota['idProveedor']}]")
        self.txt_nota_buscar_prov.setReadOnly(True)
        self.txt_nota_buscar_prod.clear()

        self.cmb_nota_estado.setCurrentText(nota["estado"])
        self.cmb_nota_estado.setEnabled(False)

        self.renglones_nota_actual = []
        for d in nota["detalles"]:
            stock_antes = int(d["existencia_actual"]) - int(d["cantidad"])
            self.renglones_nota_actual.append({
                "codigo": d["codigo"],
                "nombre": d["nombre"],
                "cantidad": int(d["cantidad"]),
                "costo": float(d["costo"]),
                "existencia_antes": max(0, stock_antes)
            })

        self.btn_guardar_nota.setText("Guardar cambios")
        self._refrescar_tabla_renglones_nota()
        self.cambiar_seccion(8)

    def agregar_renglon_a_nota(self):
        texto_busq = self.txt_nota_buscar_prod.text().strip()
        prod = self._resolver_producto_buscado(texto_busq)
        if not prod:
            QMessageBox.warning(
                self, "Producto no encontrado",
                "No se encontró ningún producto registrado con ese código o nombre. "
                "Recuerda que el producto debe estar previamente registrado en el catálogo."
            )
            return

        cant = self.spn_nota_cant.value()
        costo = self.spn_nota_costo.value()
        codigo = prod["codigo"]

        for item in self.renglones_nota_actual:
            if item["codigo"] == codigo:
                item["cantidad"] = cant
                item["costo"] = costo
                self.txt_nota_buscar_prod.clear()
                self._refrescar_tabla_renglones_nota()
                return

        self.renglones_nota_actual.append({
            "codigo": codigo,
            "nombre": prod["nombre"],
            "cantidad": cant,
            "costo": costo,
            "existencia_antes": int(prod["existencia"])
        })
        self.txt_nota_buscar_prod.clear()
        self._refrescar_tabla_renglones_nota()

    def quitar_renglon_de_nota(self, codigo: str):
        self.renglones_nota_actual = [r for r in self.renglones_nota_actual if r["codigo"] != codigo]
        self._refrescar_tabla_renglones_nota()

    def _refrescar_tabla_renglones_nota(self):
        self.tabla_nota_det.blockSignals(True)
        self.tabla_nota_det.setRowCount(len(self.renglones_nota_actual))
        total = 0.0

        for row, item in enumerate(self.renglones_nota_actual):
            self.tabla_nota_det.setRowHeight(row, 40)
            cant = int(item["cantidad"])
            costo = float(item["costo"])
            sub = cant * costo
            total += sub
            antes = int(item["existencia_antes"])
            resultante = antes + cant

            it_cant = crear_celda(str(cant), Qt.AlignCenter, negrita=True)
            it_cant.setFlags(it_cant.flags() | Qt.ItemIsEditable)
            self.tabla_nota_det.setItem(row, 0, it_cant)

            it_nom = crear_celda(f"{item['nombre']} ({item['codigo']})", negrita=True)
            it_nom.setFlags(it_nom.flags() & ~Qt.ItemIsEditable)
            self.tabla_nota_det.setItem(row, 1, it_nom)

            it_costo = crear_celda(f"{costo:.2f}", Qt.AlignRight | Qt.AlignVCenter, negrita=True)
            it_costo.setFlags(it_costo.flags() | Qt.ItemIsEditable)
            self.tabla_nota_det.setItem(row, 2, it_costo)

            it_sub = crear_celda(f"${sub:,.2f}", Qt.AlignRight | Qt.AlignVCenter)
            it_sub.setFlags(it_sub.flags() & ~Qt.ItemIsEditable)
            self.tabla_nota_det.setItem(row, 3, it_sub)

            it_ant = crear_celda(str(antes), Qt.AlignCenter)
            it_ant.setFlags(it_ant.flags() & ~Qt.ItemIsEditable)
            self.tabla_nota_det.setItem(row, 4, it_ant)

            it_res = crear_celda(str(resultante), Qt.AlignCenter, negrita=True)
            it_res.setFlags(it_res.flags() & ~Qt.ItemIsEditable)
            self.tabla_nota_det.setItem(row, 5, it_res)

            btn_quitar = QPushButton("Eliminar fila")
            btn_quitar.setStyleSheet("background:#E11D48; color:#FFFFFF; border:none; border-radius:4px; padding:4px 8px; font-size:11px; font-weight:bold;")
            btn_quitar.clicked.connect(lambda _, cod=item["codigo"]: self.quitar_renglon_de_nota(cod))
            self.tabla_nota_det.setCellWidget(row, 6, btn_quitar)

        self.lbl_nota_total.setText(f"Total de la nota: ${total:,.2f}")
        self.tabla_nota_det.blockSignals(False)

    def _al_editar_celda_detalle_nota(self, item: QTableWidgetItem):
        row = item.row()
        col = item.column()
        if row < 0 or row >= len(self.renglones_nota_actual):
            return

        texto = item.text().replace("$", "").replace(",", "").strip()
        try:
            if col == 0:
                nueva_cant = int(float(texto))
                if nueva_cant <= 0:
                    raise ValueError()
                self.renglones_nota_actual[row]["cantidad"] = nueva_cant
            elif col == 2:
                nuevo_costo = float(texto)
                if nuevo_costo <= 0:
                    raise ValueError()
                self.renglones_nota_actual[row]["costo"] = nuevo_costo
        except ValueError:
            QMessageBox.warning(self, "Valor inválido", "Ingresa una cantidad entera positiva o un costo mayor que 0.")

        self._refrescar_tabla_renglones_nota()

    def guardar_nota_recibo_bd(self):
        if not self.renglones_nota_actual:
            QMessageBox.warning(self, "Nota vacía", "Agrega al menos un producto recibido a la nota.")
            return

        folio = self.spn_nota_folio.value()
        fecha_str = self.dt_nota_fecha.date().toString("yyyy-MM-dd")

        try:
            if self.modo_formulario_nota == "AGREGAR":
                id_prov = self._resolver_proveedor_buscado(self.txt_nota_buscar_prov.text(), self.mapa_busqueda_provs_nota)
                if id_prov is None:
                    QMessageBox.warning(
                        self, "Proveedor no válido",
                        "Por favor busca y selecciona un proveedor registrado válido en el campo de proveedor."
                    )
                    return

                self.repo.agregar_nota_recibo(folio, fecha_str, int(id_prov), self.renglones_nota_actual)
                QMessageBox.information(
                    self, "Nota registrada",
                    f"La nota NR-{folio:04d} fue registrada. Se actualizó el stock y el costo de los productos recibidos."
                )
            else:
                self.repo.modificar_nota_recibo(self.folio_nota_editando, self.renglones_nota_actual)
                QMessageBox.information(
                    self, "Nota modificada",
                    f"La nota NR-{self.folio_nota_editando:04d} fue actualizada junto con las existencias y costos en inventario."
                )
            self.cambiar_seccion(7)
        except Exception as e:
            QMessageBox.critical(self, "Error al guardar nota de recibo", str(e))

    def cancelar_nota_recibo_ui(self, folio: int):
        resp = QMessageBox.question(
            self, "Cancelar nota de recibo",
            f"¿Deseas cambiar el estado de la nota NR-{folio:04d} a 'Cancelada'?\n\n"
            "Se descontarán del inventario las unidades que habían ingresado con esta nota.",
            QMessageBox.Yes | QMessageBox.No
        )
        if resp == QMessageBox.Yes:
            try:
                self.repo.cambiar_estado_nota_recibo(folio, "Cancelada")
                QMessageBox.information(self, "Nota cancelada", f"La nota NR-{folio:04d} pasó a estado Cancelada y se ajustó el stock.")
                self.cargar_notas_recibo()
            except Exception as e:
                QMessageBox.critical(self, "No se pudo cancelar", str(e))

    def setup_page_notas_ver(self):
        layout = QVBoxLayout(self.page_notas_ver)
        layout.setContentsMargins(28, 20, 28, 20)
        layout.setSpacing(14)

        lbl_t = QLabel("Consultar nota de recibo")
        lbl_t.setStyleSheet("font-size: 20px; font-weight: bold; color: #0F172A;")
        layout.addWidget(lbl_t)

        body = QHBoxLayout()
        body.setSpacing(18)

        card_left = QFrame()
        card_left.setProperty("class", "card")
        ll = QVBoxLayout(card_left)
        ll.setContentsMargins(20, 18, 20, 18)
        ll.setSpacing(10)

        self.lbl_ver_nota_folio = QLabel("NR-0000 (Solo lectura)")
        self.lbl_ver_nota_folio.setStyleSheet("font-size: 17px; font-weight: bold; color: #0F172A;")
        ll.addWidget(self.lbl_ver_nota_folio)

        grid = QGridLayout()
        grid.setVerticalSpacing(8)
        self.lbls_ver_nota = {}
        campos_n = [
            ("proveedor", "Proveedor:"),
            ("rfc", "RFC:"),
            ("fecha", "Fecha de recepción:"),
            ("estado", "Estado de la nota:"),
            ("total", "Importe total:")
        ]
        for i, (k, tit) in enumerate(campos_n):
            lk = QLabel(tit)
            lk.setStyleSheet("color: #475569; font-size: 13px; font-weight: 600;")
            lv = QLabel("-")
            lv.setWordWrap(True)
            lv.setStyleSheet("color: #0F172A; font-size: 13px; font-weight: bold;")
            grid.addWidget(lk, i, 0)
            grid.addWidget(lv, i, 1)
            self.lbls_ver_nota[k] = lv

        ll.addLayout(grid)
        ll.addStretch()

        btns = QHBoxLayout()
        btns.addStretch()
        btn_volver = QPushButton("Volver")
        btn_volver.setObjectName("btn_secondary")
        btn_volver.clicked.connect(lambda: self.cambiar_seccion(7))
        self.btn_ver_nota_mod = QPushButton("Modificar productos")
        self.btn_ver_nota_mod.setObjectName("btn_primary")
        btns.addWidget(btn_volver)
        btns.addWidget(self.btn_ver_nota_mod)
        ll.addLayout(btns)

        body.addWidget(card_left, stretch=1)

        card_right = QFrame()
        card_right.setProperty("class", "card")
        rl = QVBoxLayout(card_right)
        rl.setContentsMargins(20, 18, 20, 18)
        lbl_p = QLabel("Productos y cantidades recibidas")
        lbl_p.setProperty("class", "card_title")
        rl.addWidget(lbl_p)

        self.tabla_ver_nota_det = QTableWidget()
        self.tabla_ver_nota_det.setColumnCount(5)
        self.tabla_ver_nota_det.setHorizontalHeaderLabels([
            "Cantidad", "Nombre del Producto", "SKU", "$ Costo", "Subtotal"
        ])
        self.tabla_ver_nota_det.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.tabla_ver_nota_det.verticalHeader().setVisible(False)
        self.tabla_ver_nota_det.setEditTriggers(QTableWidget.NoEditTriggers)
        rl.addWidget(self.tabla_ver_nota_det)

        body.addWidget(card_right, stretch=2)
        layout.addLayout(body)

    def ir_consultar_nota_recibo(self, folio: int):
        nota = self.repo.obtener_nota_completa(folio)
        if not nota:
            QMessageBox.warning(self, "No encontrada", "No se encontró la nota de recibo.")
            return

        self.lbl_ver_nota_folio.setText(f"Nota de Recibo NR-{folio:04d} (Solo lectura)")
        self.lbls_ver_nota["proveedor"].setText(nota["proveedor"])
        self.lbls_ver_nota["rfc"].setText(nota["rfc_proveedor"])
        self.lbls_ver_nota["fecha"].setText(nota["fecha_fmt"])
        self.lbls_ver_nota["estado"].setText(nota["estado"])
        self.lbls_ver_nota["total"].setText(f"${nota['total']:,.2f}")

        try:
            self.btn_ver_nota_mod.clicked.disconnect()
        except Exception:
            pass
        self.btn_ver_nota_mod.setEnabled(nota["estado"] == "Activa")
        self.btn_ver_nota_mod.clicked.connect(lambda: self.ir_modificar_nota_recibo(folio))

        detalles = nota["detalles"]
        self.tabla_ver_nota_det.setRowCount(len(detalles) + 1)
        tot_u = 0
        for r, d in enumerate(detalles):
            self.tabla_ver_nota_det.setRowHeight(r, 38)
            tot_u += int(d["cantidad"])
            self.tabla_ver_nota_det.setItem(r, 0, crear_celda(str(d["cantidad"]), Qt.AlignCenter, negrita=True))
            self.tabla_ver_nota_det.setItem(r, 1, crear_celda(d["nombre"], negrita=True))
            self.tabla_ver_nota_det.setItem(r, 2, crear_celda(d["codigo"]))
            self.tabla_ver_nota_det.setItem(r, 3, crear_celda(f"${d['costo']:,.2f}", Qt.AlignRight | Qt.AlignVCenter))
            self.tabla_ver_nota_det.setItem(r, 4, crear_celda(f"${d['subtotal']:,.2f}", Qt.AlignRight | Qt.AlignVCenter, negrita=True))

        ult = len(detalles)
        self.tabla_ver_nota_det.setRowHeight(ult, 40)
        self.tabla_ver_nota_det.setItem(ult, 0, crear_celda(str(tot_u), Qt.AlignCenter, negrita=True))
        self.tabla_ver_nota_det.setItem(ult, 1, crear_celda("TOTAL DE LA NOTA", negrita=True))
        self.tabla_ver_nota_det.setItem(ult, 2, crear_celda("-"))
        self.tabla_ver_nota_det.setItem(ult, 3, crear_celda("-"))
        self.tabla_ver_nota_det.setItem(ult, 4, crear_celda(f"${nota['total']:,.2f}", Qt.AlignRight | Qt.AlignVCenter, negrita=True))

        self.cambiar_seccion(9)

    # =========================================================================
    # MÓDULO 4: REPORTES (Sin marco gris oscuro, alto contraste y buscador)
    # =========================================================================
    def setup_page_reportes(self):
        layout = QVBoxLayout(self.page_reportes)
        layout.setContentsMargins(20, 16, 20, 16)

        self.tabs_rep = QTabWidget()

        self.tab_rep_rec = QWidget()
        self.tab_rep_rec.setStyleSheet("background-color: #F7F8FF;")
        self._build_tab_recibos()
        self.tabs_rep.addTab(self.tab_rep_rec, "1. Reporte de Recibo de Mercancías")

        self.tab_rep_inv = QWidget()
        self.tab_rep_inv.setStyleSheet("background-color: #F7F8FF;")
        self._build_tab_inventario()
        self.tabs_rep.addTab(self.tab_rep_inv, "2. Reporte de Inventario")

        self.tab_rep_val = QWidget()
        self.tab_rep_val.setStyleSheet("background-color: #F7F8FF;")
        self._build_tab_valor_inventario()
        self.tabs_rep.addTab(self.tab_rep_val, "3. Reporte de Valor de Inventario")

        self.tab_rep_aba = QWidget()
        self.tab_rep_aba.setStyleSheet("background-color: #F7F8FF;")
        self._build_tab_abastecimiento()
        self.tabs_rep.addTab(self.tab_rep_aba, "4. Reporte de Abastecimiento")

        self.tab_rep_ven = QWidget()
        self.tab_rep_ven.setStyleSheet("background-color: #F7F8FF;")
        self._build_tab_ventas()
        self.tabs_rep.addTab(self.tab_rep_ven, "5. Reporte de Ventas")

        layout.addWidget(self.tabs_rep)

    def _crear_encabezado_institucional(self, nombre_reporte: str) -> tuple[QFrame, QLabel]:
        frame = QFrame()
        frame.setProperty("class", "card")
        hl = QHBoxLayout(frame)
        hl.setContentsMargins(16, 12, 16, 12)

        left = QVBoxLayout()
        lbl_emp = QLabel(f"EMPRESA: {NOMBRE_EMPRESA}")
        lbl_emp.setStyleSheet("color: #3730A3; font-size: 13px; font-weight: bold;")
        lbl_rep = QLabel(f"REPORTE: {nombre_reporte}")
        lbl_rep.setStyleSheet("color: #0F172A; font-size: 17px; font-weight: bold;")
        left.addWidget(lbl_emp)
        left.addWidget(lbl_rep)

        right = QVBoxLayout()
        lbl_fecha = QLabel(f"Fecha del reporte: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
        lbl_fecha.setStyleSheet("color: #0F172A; font-size: 12px; font-weight: bold;")
        right.addWidget(lbl_fecha, alignment=Qt.AlignRight)

        hl.addLayout(left)
        hl.addStretch()
        hl.addLayout(right)
        return frame, lbl_fecha

    def exportar_tabla_pdf(self, titulo_reporte: str, tabla: QTableWidget, texto_resumen: str, meta_extra: str = ""):
        archivo, _ = QFileDialog.getSaveFileName(
            self, f"Exportar {titulo_reporte} a PDF",
            f"{titulo_reporte.replace(' ', '_')}.pdf", "Archivos PDF (*.pdf)"
        )
        if not archivo:
            return

        fecha_str = datetime.now().strftime("%d/%m/%Y %H:%M")
        cols = tabla.columnCount()
        rows = tabla.rowCount()

        html = f"""
        <html>
        <head>
        <style>
            body {{ font-family: Arial, sans-serif; color: #0F172A; font-size: 10pt; }}
            .header {{ border-bottom: 2px solid #24366B; padding-bottom: 8px; margin-bottom: 12px; }}
            .empresa {{ font-size: 15pt; font-weight: bold; color: #3730A3; }}
            .titulo {{ font-size: 13pt; font-weight: bold; color: #0F172A; margin-top: 4px; }}
            .fecha {{ font-size: 9.5pt; color: #0F172A; margin-top: 4px; font-weight: bold; }}
            table {{ width: 100%; border-collapse: collapse; margin-top: 10px; }}
            th {{ background-color: #24366B; color: #FFFFFF; padding: 6px; font-size: 9pt; border: 1px solid #24366B; text-align: left; }}
            td {{ padding: 5px; font-size: 9pt; border: 1px solid #C7D2E8; color: #0F172A; }}
            .resumen {{ margin-top: 14px; padding: 8px; background-color: #EEF2FF; border: 1px solid #C7D2E8; font-weight: bold; font-size: 9.5pt; }}
        </style>
        </head>
        <body>
            <div class="header">
                <div class="empresa">Empresa: {NOMBRE_EMPRESA}</div>
                <div class="titulo">Reporte: {titulo_reporte}</div>
                <div class="fecha">Fecha de realización del reporte: {fecha_str}</div>
                <div class="fecha">{meta_extra}</div>
            </div>
            <table>
                <thead>
                    <tr>
        """
        for c in range(cols):
            enc_item = tabla.horizontalHeaderItem(c)
            html += f"<th>{enc_item.text() if enc_item else ''}</th>"
        html += "</tr></thead><tbody>"

        for r in range(rows):
            html += "<tr>"
            for c in range(cols):
                it = tabla.item(r, c)
                txt = it.text() if it else ""
                html += f"<td>{txt}</td>"
            html += "</tr>"

        html += f"""
                </tbody>
            </table>
            <div class="resumen">{texto_resumen}</div>
        </body>
        </html>
        """

        printer = QPrinter(QPrinter.HighResolution)
        printer.setOutputFormat(QPrinter.PdfFormat)
        printer.setOutputFileName(archivo)
        printer.setPageSize(QPageSize(QPageSize.Letter))
        printer.setPageOrientation(QPageLayout.Landscape)
        printer.setPageMargins(QMarginsF(12, 12, 12, 12), QPageLayout.Millimeter)

        doc = QTextDocument()
        doc.setHtml(html)
        doc.print_(printer)

        QMessageBox.information(self, "PDF generado", f"El reporte se exportó correctamente a:\n{archivo}")

    def exportar_tabla_excel(self, titulo_reporte: str, tabla: QTableWidget, texto_resumen: str, meta_extra: str = ""):
        if not OPENPYXL_DISPONIBLE:
            QMessageBox.warning(
                self, "Librería faltante",
                "Para exportar a Excel (.xlsx), instala openpyxl ejecutando:\n\npip install openpyxl"
            )
            return

        archivo, _ = QFileDialog.getSaveFileName(
            self, f"Exportar {titulo_reporte} a Excel",
            f"{titulo_reporte.replace(' ', '_')}.xlsx", "Archivos Excel (*.xlsx)"
        )
        if not archivo:
            return

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = titulo_reporte[:31]

        ws.append([f"EMPRESA: {NOMBRE_EMPRESA}"])
        ws.append([f"REPORTE: {titulo_reporte}"])
        ws.append([f"Fecha del reporte: {datetime.now().strftime('%d/%m/%Y %H:%M')}"])
        if meta_extra:
            ws.append([meta_extra])
        ws.append([])

        fila_enc = ws.max_row + 1
        cols = tabla.columnCount()
        headers = [tabla.horizontalHeaderItem(c).text() if tabla.horizontalHeaderItem(c) else "" for c in range(cols)]
        ws.append(headers)

        header_fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF")
        for col_idx in range(1, cols + 1):
            cell = ws.cell(row=fila_enc, column=col_idx)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center")

        for r in range(tabla.rowCount()):
            fila_vals = [tabla.item(r, c).text() if tabla.item(r, c) else "" for c in range(cols)]
            ws.append(fila_vals)

        ws.append([])
        ws.append([texto_resumen])
        ws.cell(row=ws.max_row, column=1).font = Font(bold=True)

        for col in ws.columns:
            max_len = max(len(str(cell.value or "")) for cell in col)
            col_letter = col[0].column_letter
            ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 45)

        wb.save(archivo)
        QMessageBox.information(self, "Excel generado", f"El reporte se exportó correctamente a:\n{archivo}")

    # --- REPORTE 1: RECIBO DE MERCANCÍAS (Con Buscador de Proveedor y Calendario Claro) ---
    def _build_tab_recibos(self):
        l = QVBoxLayout(self.tab_rep_rec)
        l.setSpacing(10)

        card_head = QFrame()
        card_head.setProperty("class", "card")
        hl = QHBoxLayout(card_head)
        hl.setContentsMargins(16, 12, 16, 12)

        left = QVBoxLayout()
        self.lbl_rec_empresa = QLabel(f"Empresa: {NOMBRE_EMPRESA}")
        self.lbl_rec_empresa.setStyleSheet("color: #3730A3; font-size: 14px; font-weight: bold;")
        self.lbl_rec_nombre_rep = QLabel("Nombre del reporte: Reporte de recibo de mercancías")
        self.lbl_rec_nombre_rep.setStyleSheet("color: #0F172A; font-size: 17px; font-weight: bold;")
        self.lbl_rec_prov_sel = QLabel("Proveedor: Todos los proveedores")
        self.lbl_rec_prov_sel.setStyleSheet("color: #0F172A; font-size: 13px; font-weight: bold;")
        left.addWidget(self.lbl_rec_empresa)
        left.addWidget(self.lbl_rec_nombre_rep)
        left.addWidget(self.lbl_rec_prov_sel)

        right = QVBoxLayout()
        self.lbl_rec_folio_rep = QLabel("Folio del reporte: RPT-REC-0019")
        self.lbl_rec_folio_rep.setStyleSheet("color: #0F172A; font-size: 13px; font-weight: bold;")
        self.lbl_rec_fecha_emision = QLabel(f"Fecha de realización: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
        self.lbl_rec_fecha_emision.setStyleSheet("color: #0F172A; font-size: 12px; font-weight: bold;")
        self.lbl_rec_rango = QLabel("Periodo consultado: 01/10/2026 al 30/11/2026")
        self.lbl_rec_rango.setStyleSheet("color: #3730A3; font-size: 12px; font-weight: bold;")
        right.addWidget(self.lbl_rec_folio_rep, alignment=Qt.AlignRight)
        right.addWidget(self.lbl_rec_fecha_emision, alignment=Qt.AlignRight)
        right.addWidget(self.lbl_rec_rango, alignment=Qt.AlignRight)

        hl.addLayout(left)
        hl.addStretch()
        hl.addLayout(right)
        l.addWidget(card_head)

        # Tarjeta blanca para los filtros (evita que el texto negro se pierda en el fondo)
        card_filtros = QFrame()
        card_filtros.setProperty("class", "card")
        filtros = QHBoxLayout(card_filtros)
        filtros.setContentsMargins(14, 10, 14, 10)

        self.dt_rec_ini = QDateEdit(QDate(2026, 10, 1))
        self.dt_rec_ini.setDisplayFormat("dd/MM/yyyy")
        configurar_date_edit_claro(self.dt_rec_ini)

        self.dt_rec_fin = QDateEdit(QDate(2026, 11, 30))
        self.dt_rec_fin.setDisplayFormat("dd/MM/yyyy")
        configurar_date_edit_claro(self.dt_rec_fin)

        self.txt_rep_rec_buscar_prov = QLineEdit()
        self.txt_rep_rec_buscar_prov.setPlaceholderText("🔍 Buscar proveedor (o dejar vacío para Todos)...")
        self.txt_rep_rec_buscar_prov.returnPressed.connect(lambda: self.generar_reporte_recibos(mostrar_alerta_fechas=True))

        btn_gen = QPushButton("Generar reporte")
        btn_gen.setObjectName("btn_report_generate")
        btn_gen.setStyleSheet("QPushButton { background-color: #4F46E5; color: #FFFFFF; border: none; border-radius: 6px; padding: 8px 16px; font-weight: bold; font-size: 13px; } QPushButton:hover { background-color: #3730A3; } QPushButton:pressed { background-color: #4338CA; }")
        btn_gen.clicked.connect(lambda: self.generar_reporte_recibos(mostrar_alerta_fechas=True))

        btn_pdf = QPushButton("Exportar PDF")
        btn_pdf.setObjectName("btn_secondary")
        btn_pdf.clicked.connect(self._exportar_pdf_recibo_mercancias)

        btn_xls = QPushButton("Exportar Excel")
        btn_xls.setObjectName("btn_secondary")
        btn_xls.clicked.connect(self._exportar_excel_recibo_mercancias)

        filtros.addWidget(self._lbl("Fecha inicio:"))
        filtros.addWidget(self.dt_rec_ini)
        filtros.addWidget(self._lbl("Fecha fin:"))
        filtros.addWidget(self.dt_rec_fin)
        filtros.addWidget(self._lbl("Proveedor:"))
        filtros.addWidget(self.txt_rep_rec_buscar_prov, stretch=1)
        filtros.addWidget(btn_gen)
        filtros.addWidget(btn_pdf)
        filtros.addWidget(btn_xls)
        l.addWidget(card_filtros)

        self.tabla_rep_rec = QTableWidget()
        self.tabla_rep_rec.setColumnCount(6)
        self.tabla_rep_rec.setHorizontalHeaderLabels([
            "Cantidad", "Nombre del producto", "$ Costo del recibo",
            "$ Costo unitario", "Folio Nota / Fecha", "Proveedor"
        ])
        self.tabla_rep_rec.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.tabla_rep_rec.horizontalHeader().setSectionResizeMode(5, QHeaderView.Stretch)
        self.tabla_rep_rec.verticalHeader().setVisible(False)
        self.tabla_rep_rec.setEditTriggers(QTableWidget.NoEditTriggers)
        l.addWidget(self.tabla_rep_rec)

        self.lbl_totales_rec = QLabel()
        self.lbl_totales_rec.setStyleSheet(
            "background-color: #24366B; color: #FFFFFF; padding: 10px 14px; "
            "border-radius: 6px; font-weight: bold; font-size: 13px;"
        )
        l.addWidget(self.lbl_totales_rec)

    def actualizar_buscador_proveedores_reporte(self):
        self.mapa_busqueda_provs_rep.clear()
        opciones = ["Todos los proveedores"]
        for pr in self.repo.buscar_proveedores(incluir_inactivos=True):
            p = normalizar_proveedor(pr)
            etiqueta = f"{p['empresa']} · RFC: {p['rfc']} [ID: {p['idProveedor']}]"
            self.mapa_busqueda_provs_rep[etiqueta] = p["idProveedor"]
            opciones.append(etiqueta)

        comp = QCompleter(opciones, self)
        comp.setCaseSensitivity(Qt.CaseInsensitive)
        comp.setFilterMode(Qt.MatchContains)
        self.txt_rep_rec_buscar_prov.setCompleter(comp)

    def generar_reporte_recibos(self, mostrar_alerta_fechas: bool = False):
        if self.dt_rec_ini.date() >= self.dt_rec_fin.date():
            if mostrar_alerta_fechas:
                QMessageBox.warning(
                    self, "Rango de fechas inválido",
                    "La fecha de inicio debe ser menor que la fecha de fin. Por favor corrige el rango seleccionado."
                )
            return

        f_ini_sql = self.dt_rec_ini.date().toString("yyyy-MM-dd")
        f_fin_sql = self.dt_rec_fin.date().toString("yyyy-MM-dd")
        f_ini_fmt = self.dt_rec_ini.date().toString("dd/MM/yyyy")
        f_fin_fmt = self.dt_rec_fin.date().toString("dd/MM/yyyy")

        texto_prov = self.txt_rep_rec_buscar_prov.text().strip()
        if not texto_prov or texto_prov.lower().startswith("todos"):
            id_prov = None
            nombre_prov = "Todos los proveedores"
        else:
            id_prov = self._resolver_proveedor_buscado(texto_prov, self.mapa_busqueda_provs_rep)
            if id_prov is None:
                if mostrar_alerta_fechas:
                    QMessageBox.warning(
                        self, "Proveedor no encontrado",
                        "No se encontró un proveedor que coincida con la búsqueda. "
                        "Selecciona uno de la lista o deja el campo vacío para consultar Todos."
                    )
                return
            nombre_prov = texto_prov

        folio_rep_str = f"RPT-REC-{self.contador_folio_rep_rec:04d}"
        if mostrar_alerta_fechas:
            self.contador_folio_rep_rec += 1

        self.lbl_rec_folio_rep.setText(f"Folio del reporte: {folio_rep_str}")
        self.lbl_rec_fecha_emision.setText(f"Fecha de realización: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
        self.lbl_rec_rango.setText(f"Periodo: {f_ini_fmt} al {f_fin_fmt}")
        self.lbl_rec_prov_sel.setText(f"Proveedor: {nombre_prov}")

        renglones = self.repo.reporte_recibo_mercancias_detalle(f_ini_sql, f_fin_sql, id_prov)
        if not renglones:
            self.tabla_rep_rec.setRowCount(0)
            self.lbl_totales_rec.setText("No se encontraron recepciones de mercancía en el rango de fechas y proveedor seleccionados.")
            return

        self.tabla_rep_rec.setRowCount(len(renglones) + 1)
        tot_cantidad = 0
        tot_costo_recibo = 0.0

        for r, item in enumerate(renglones):
            self.tabla_rep_rec.setRowHeight(r, 38)
            cant = int(item["cantidad"])
            costo_rec = float(item["costo_recibo"])
            costo_u = float(item["costo_unitario"])
            tot_cantidad += cant
            tot_costo_recibo += costo_rec

            self.tabla_rep_rec.setItem(r, 0, crear_celda(str(cant), Qt.AlignCenter, negrita=True))
            self.tabla_rep_rec.setItem(r, 1, crear_celda(str(item["nombre_producto"]), negrita=True))
            self.tabla_rep_rec.setItem(r, 2, crear_celda(f"${costo_rec:,.2f}", Qt.AlignRight | Qt.AlignVCenter, negrita=True))
            self.tabla_rep_rec.setItem(r, 3, crear_celda(f"${costo_u:,.2f}", Qt.AlignRight | Qt.AlignVCenter))
            self.tabla_rep_rec.setItem(r, 4, crear_celda(f"NR-{int(item['folio_nota']):04d} ({item['fecha_nota']})", Qt.AlignCenter))
            self.tabla_rep_rec.setItem(r, 5, crear_celda(str(item["proveedor"])))

        ult = len(renglones)
        self.tabla_rep_rec.setRowHeight(ult, 42)
        self.tabla_rep_rec.setItem(ult, 0, crear_celda(str(tot_cantidad), Qt.AlignCenter, negrita=True))
        self.tabla_rep_rec.setItem(ult, 1, crear_celda("TOTAL RECIBIDO EN EL PERIODO", negrita=True))
        self.tabla_rep_rec.setItem(ult, 2, crear_celda(f"${tot_costo_recibo:,.2f}", Qt.AlignRight | Qt.AlignVCenter, negrita=True))
        self.tabla_rep_rec.setItem(ult, 3, crear_celda("-", Qt.AlignCenter))
        self.tabla_rep_rec.setItem(ult, 4, crear_celda("-", Qt.AlignCenter))
        self.tabla_rep_rec.setItem(ult, 5, crear_celda(nombre_prov, negrita=True))

        self.lbl_totales_rec.setText(
            f"{self.lbl_rec_folio_rep.text()}   |   {self.lbl_rec_rango.text()}   |   "
            f"Total Unidades Recibidas: {tot_cantidad}   |   Costo Total del Recibo: ${tot_costo_recibo:,.2f}"
        )

    def _exportar_pdf_recibo_mercancias(self):
        meta = f"{self.lbl_rec_folio_rep.text()} | {self.lbl_rec_rango.text()} | {self.lbl_rec_prov_sel.text()}"
        self.exportar_tabla_pdf("Reporte de recibo de mercancías", self.tabla_rep_rec, self.lbl_totales_rec.text(), meta)

    def _exportar_excel_recibo_mercancias(self):
        meta = f"{self.lbl_rec_folio_rep.text()} | {self.lbl_rec_rango.text()} | {self.lbl_rec_prov_sel.text()}"
        self.exportar_tabla_excel("Reporte de recibo de mercancías", self.tabla_rep_rec, self.lbl_totales_rec.text(), meta)

    # --- REPORTE 2: INVENTARIO ---
    def _build_tab_inventario(self):
        l = QVBoxLayout(self.tab_rep_inv)
        enc, self.lbl_fecha_inv = self._crear_encabezado_institucional("Reporte de Inventario")
        l.addWidget(enc)

        card_f = QFrame()
        card_f.setProperty("class", "card")
        filtros = QHBoxLayout(card_f)
        filtros.setContentsMargins(14, 10, 14, 10)

        self.cmb_rep_inv_cat = QComboBox()
        self.cmb_rep_inv_cat.addItems(["Todas", "Micas", "Accesorios", "Refacciones", "Recargas"])
        self.cmb_rep_inv_cat.currentTextChanged.connect(self.generar_reporte_inventario)

        btn_gen = QPushButton("Generar reporte")
        btn_gen.setObjectName("btn_report_generate")
        btn_gen.setStyleSheet("QPushButton { background-color: #4F46E5; color: #FFFFFF; border: none; border-radius: 6px; padding: 8px 16px; font-weight: bold; font-size: 13px; } QPushButton:hover { background-color: #3730A3; } QPushButton:pressed { background-color: #4338CA; }")
        btn_gen.clicked.connect(self.generar_reporte_inventario)

        btn_pdf = QPushButton("Exportar PDF")
        btn_pdf.setObjectName("btn_secondary")
        btn_pdf.clicked.connect(lambda: self.exportar_tabla_pdf("Reporte de Inventario", self.tabla_rep_inv, self.lbl_resumen_rep_inv.text()))

        btn_xls = QPushButton("Exportar Excel")
        btn_xls.setObjectName("btn_secondary")
        btn_xls.clicked.connect(lambda: self.exportar_tabla_excel("Reporte de Inventario", self.tabla_rep_inv, self.lbl_resumen_rep_inv.text()))

        filtros.addWidget(self._lbl("Categoría:"))
        filtros.addWidget(self.cmb_rep_inv_cat)
        filtros.addWidget(btn_gen)
        filtros.addStretch()
        filtros.addWidget(btn_pdf)
        filtros.addWidget(btn_xls)
        l.addWidget(card_f)

        self.tabla_rep_inv = QTableWidget()
        self.tabla_rep_inv.setColumnCount(8)
        self.tabla_rep_inv.setHorizontalHeaderLabels([
            "Nombre del producto", "Existencia", "Unidad de Medida",
            "Punto de reorden", "Máximo", "$ Costo", "$ Precio", "Estado"
        ])
        self.tabla_rep_inv.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.tabla_rep_inv.verticalHeader().setVisible(False)
        self.tabla_rep_inv.setEditTriggers(QTableWidget.NoEditTriggers)
        l.addWidget(self.tabla_rep_inv)

        self.lbl_resumen_rep_inv = QLabel()
        self.lbl_resumen_rep_inv.setStyleSheet(
            "background-color: #24366B; color: #FFFFFF; padding: 10px 14px; "
            "border-radius: 6px; font-weight: bold; font-size: 13px;"
        )
        l.addWidget(self.lbl_resumen_rep_inv)

    def generar_reporte_inventario(self):
        self.lbl_fecha_inv.setText(f"Fecha del reporte: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
        cat = self.cmb_rep_inv_cat.currentText()
        raw_prods = self.repo.buscar_productos(categoria=cat, incluir_inactivos=True)
        productos = [normalizar_producto(p) for p in raw_prods]

        self.tabla_rep_inv.setRowCount(len(productos))
        if not productos:
            self.lbl_resumen_rep_inv.setText("No hay productos registrados para la categoría seleccionada.")
            return

        total_exist = 0
        bajos = 0
        agotados = 0
        for r, p in enumerate(productos):
            self.tabla_rep_inv.setRowHeight(r, 38)
            self.tabla_rep_inv.setItem(r, 0, crear_celda(p["nombre"], negrita=True))
            self.tabla_rep_inv.setItem(r, 1, crear_celda(str(p["existencia"]), Qt.AlignCenter))
            self.tabla_rep_inv.setItem(r, 2, crear_celda(p["unidadMedida"], Qt.AlignCenter))
            self.tabla_rep_inv.setItem(r, 3, crear_celda(str(p["stockMinimo"]), Qt.AlignCenter))
            self.tabla_rep_inv.setItem(r, 4, crear_celda(str(p["stockMaximo"]), Qt.AlignCenter))
            self.tabla_rep_inv.setItem(r, 5, crear_celda(f"${p['costo']:,.2f}", Qt.AlignRight | Qt.AlignVCenter))
            self.tabla_rep_inv.setItem(r, 6, crear_celda(f"${p['precio']:,.2f}", Qt.AlignRight | Qt.AlignVCenter))
            estado_str = "Activo" if p["estado"] else "Inactivo"
            self.tabla_rep_inv.setItem(r, 7, crear_celda(estado_str, Qt.AlignCenter, negrita=True))

            if p["estado"]:
                total_exist += p["existencia"]
                if p["existencia"] == 0:
                    agotados += 1
                elif p["existencia"] <= p["stockMinimo"]:
                    bajos += 1

        self.lbl_resumen_rep_inv.setText(
            f"Total productos listados: {len(productos)}   |   Unidades activas en existencia: {total_exist}   |   "
            f"Con stock bajo: {bajos}   |   Agotados: {agotados}"
        )

    # --- REPORTE 3: VALOR DE INVENTARIO ---
    def _build_tab_valor_inventario(self):
        l = QVBoxLayout(self.tab_rep_val)
        enc, self.lbl_fecha_val = self._crear_encabezado_institucional("Reporte de Valor de Inventario")
        l.addWidget(enc)

        card_f = QFrame()
        card_f.setProperty("class", "card")
        btn_row = QHBoxLayout(card_f)
        btn_row.setContentsMargins(14, 10, 14, 10)

        btn_gen = QPushButton("Generar reporte")
        btn_gen.setObjectName("btn_report_generate")
        btn_gen.setStyleSheet("QPushButton { background-color: #4F46E5; color: #FFFFFF; border: none; border-radius: 6px; padding: 8px 16px; font-weight: bold; font-size: 13px; } QPushButton:hover { background-color: #3730A3; } QPushButton:pressed { background-color: #4338CA; }")
        btn_gen.clicked.connect(self.generar_reporte_valor)

        btn_pdf = QPushButton("Exportar PDF")
        btn_pdf.setObjectName("btn_secondary")
        btn_pdf.clicked.connect(lambda: self.exportar_tabla_pdf("Reporte de Valor de Inventario", self.tabla_rep_val, self.lbl_totales_val.text()))

        btn_xls = QPushButton("Exportar Excel")
        btn_xls.setObjectName("btn_secondary")
        btn_xls.clicked.connect(lambda: self.exportar_tabla_excel("Reporte de Valor de Inventario", self.tabla_rep_val, self.lbl_totales_val.text()))

        btn_row.addWidget(btn_gen)
        btn_row.addStretch()
        btn_row.addWidget(btn_pdf)
        btn_row.addWidget(btn_xls)
        l.addWidget(card_f)

        self.tabla_rep_val = QTableWidget()
        self.tabla_rep_val.setColumnCount(9)
        self.tabla_rep_val.setHorizontalHeaderLabels([
            "Nombre del producto", "Existencia", "Unidad de medida",
            "$ Costo unitario", "$ Precio Unitario", "$ Valor en costo",
            "$ Valor en precio", "$ Margen de utilidad", "Estado"
        ])
        self.tabla_rep_val.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.tabla_rep_val.verticalHeader().setVisible(False)
        self.tabla_rep_val.setEditTriggers(QTableWidget.NoEditTriggers)
        l.addWidget(self.tabla_rep_val)

        self.lbl_totales_val = QLabel()
        self.lbl_totales_val.setStyleSheet(
            "background-color: #24366B; color: #FFFFFF; padding: 10px 14px; "
            "border-radius: 6px; font-weight: bold; font-size: 13px;"
        )
        l.addWidget(self.lbl_totales_val)

    def generar_reporte_valor(self):
        self.lbl_fecha_val.setText(f"Fecha del reporte: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
        raw_prods = self.repo.buscar_productos(incluir_inactivos=True)
        productos = [normalizar_producto(p) for p in raw_prods]
        self.tabla_rep_val.setRowCount(len(productos) + 1)

        suma_costo_unit = 0.0
        suma_precio_unit = 0.0
        suma_val_costo = 0.0
        suma_val_precio = 0.0
        suma_margen = 0.0
        suma_exist_activos = 0

        for r, p in enumerate(productos):
            self.tabla_rep_val.setRowHeight(r, 38)
            val_costo = p["existencia"] * p["costo"]
            val_precio = p["existencia"] * p["precio"]
            margen_utilidad = val_precio - val_costo

            self.tabla_rep_val.setItem(r, 0, crear_celda(p["nombre"], negrita=True))
            self.tabla_rep_val.setItem(r, 1, crear_celda(str(p["existencia"]), Qt.AlignCenter))
            self.tabla_rep_val.setItem(r, 2, crear_celda(p["unidadMedida"], Qt.AlignCenter))
            self.tabla_rep_val.setItem(r, 3, crear_celda(f"${p['costo']:,.2f}", Qt.AlignRight | Qt.AlignVCenter))
            self.tabla_rep_val.setItem(r, 4, crear_celda(f"${p['precio']:,.2f}", Qt.AlignRight | Qt.AlignVCenter))
            self.tabla_rep_val.setItem(r, 5, crear_celda(f"${val_costo:,.2f}", Qt.AlignRight | Qt.AlignVCenter))
            self.tabla_rep_val.setItem(r, 6, crear_celda(f"${val_precio:,.2f}", Qt.AlignRight | Qt.AlignVCenter))
            self.tabla_rep_val.setItem(r, 7, crear_celda(f"${margen_utilidad:,.2f}", Qt.AlignRight | Qt.AlignVCenter, negrita=True))
            self.tabla_rep_val.setItem(r, 8, crear_celda("Activo" if p["estado"] else "Inactivo", Qt.AlignCenter))

            if p["estado"]:
                suma_exist_activos += p["existencia"]
                suma_costo_unit += p["costo"]
                suma_precio_unit += p["precio"]
                suma_val_costo += val_costo
                suma_val_precio += val_precio
                suma_margen += margen_utilidad

        ult = len(productos)
        self.tabla_rep_val.setRowHeight(ult, 42)
        self.tabla_rep_val.setItem(ult, 0, crear_celda("TOTALES (SOLO PRODUCTOS ACTIVOS)", negrita=True))
        self.tabla_rep_val.setItem(ult, 1, crear_celda(str(suma_exist_activos), Qt.AlignCenter, negrita=True))
        self.tabla_rep_val.setItem(ult, 2, crear_celda("-", Qt.AlignCenter))
        self.tabla_rep_val.setItem(ult, 3, crear_celda(f"${suma_costo_unit:,.2f}", Qt.AlignRight | Qt.AlignVCenter, negrita=True))
        self.tabla_rep_val.setItem(ult, 4, crear_celda(f"${suma_precio_unit:,.2f}", Qt.AlignRight | Qt.AlignVCenter, negrita=True))
        self.tabla_rep_val.setItem(ult, 5, crear_celda(f"${suma_val_costo:,.2f}", Qt.AlignRight | Qt.AlignVCenter, negrita=True))
        self.tabla_rep_val.setItem(ult, 6, crear_celda(f"${suma_val_precio:,.2f}", Qt.AlignRight | Qt.AlignVCenter, negrita=True))
        self.tabla_rep_val.setItem(ult, 7, crear_celda(f"${suma_margen:,.2f}", Qt.AlignRight | Qt.AlignVCenter, negrita=True))
        self.tabla_rep_val.setItem(ult, 8, crear_celda("Activos", Qt.AlignCenter, negrita=True))

        self.lbl_totales_val.setText(
            f"SUMA DE PRODUCTOS ACTIVOS ->   "
            f"Costos Unit.: ${suma_costo_unit:,.2f}   |   "
            f"Precios Unit.: ${suma_precio_unit:,.2f}   |   "
            f"Valor en Costo: ${suma_val_costo:,.2f}   |   "
            f"Valor en Precio: ${suma_val_precio:,.2f}   |   "
            f"Margen de Utilidad: ${suma_margen:,.2f}"
        )

    # --- REPORTE 4: ABASTECIMIENTO ---
    def _build_tab_abastecimiento(self):
        l = QVBoxLayout(self.tab_rep_aba)
        enc, self.lbl_fecha_aba = self._crear_encabezado_institucional("Reporte de Abastecimiento")
        l.addWidget(enc)

        card_f = QFrame()
        card_f.setProperty("class", "card")
        top = QHBoxLayout(card_f)
        top.setContentsMargins(14, 10, 14, 10)

        lbl_crit = QLabel("Productos cuya existencia es menor o igual al punto de reorden (Existencia <= P.R.)")
        lbl_crit.setStyleSheet("color: #B45309; font-weight: bold; font-size: 13px;")
        btn_gen = QPushButton("Generar reporte")
        btn_gen.setObjectName("btn_report_generate")
        btn_gen.setStyleSheet("QPushButton { background-color: #4F46E5; color: #FFFFFF; border: none; border-radius: 6px; padding: 8px 16px; font-weight: bold; font-size: 13px; } QPushButton:hover { background-color: #3730A3; } QPushButton:pressed { background-color: #4338CA; }")
        btn_gen.clicked.connect(self.generar_reporte_abastecimiento)

        btn_pdf = QPushButton("Exportar PDF")
        btn_pdf.setObjectName("btn_secondary")
        btn_pdf.clicked.connect(lambda: self.exportar_tabla_pdf("Reporte de Abastecimiento", self.tabla_rep_aba, self.lbl_totales_aba.text()))

        btn_xls = QPushButton("Exportar Excel")
        btn_xls.setObjectName("btn_secondary")
        btn_xls.clicked.connect(lambda: self.exportar_tabla_excel("Reporte de Abastecimiento", self.tabla_rep_aba, self.lbl_totales_aba.text()))

        top.addWidget(lbl_crit)
        top.addStretch()
        top.addWidget(btn_gen)
        top.addWidget(btn_pdf)
        top.addWidget(btn_xls)
        l.addWidget(card_f)

        self.tabla_rep_aba = QTableWidget()
        self.tabla_rep_aba.setColumnCount(7)
        self.tabla_rep_aba.setHorizontalHeaderLabels([
            "Nombre del producto", "Existencia", "Unidad de medida",
            "Punto de reorden", "Máximo", "Cantidad a abastecer", "Costo del abastecimiento"
        ])
        self.tabla_rep_aba.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.tabla_rep_aba.verticalHeader().setVisible(False)
        self.tabla_rep_aba.setEditTriggers(QTableWidget.NoEditTriggers)
        l.addWidget(self.tabla_rep_aba)

        self.lbl_totales_aba = QLabel()
        self.lbl_totales_aba.setStyleSheet(
            "background-color: #24366B; color: #FFFFFF; padding: 10px 14px; "
            "border-radius: 6px; font-weight: bold; font-size: 13px;"
        )
        l.addWidget(self.lbl_totales_aba)

    def generar_reporte_abastecimiento(self):
        self.lbl_fecha_aba.setText(f"Fecha del reporte: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
        raw_prods = self.repo.buscar_productos(solo_stock_bajo=True, incluir_inactivos=False)
        productos = [normalizar_producto(p) for p in raw_prods]

        if not productos:
            self.tabla_rep_aba.setRowCount(0)
            self.lbl_totales_aba.setText("No se requiere abastecimiento: todos los productos superan su punto de reorden.")
            return

        self.tabla_rep_aba.setRowCount(len(productos) + 1)
        total_unidades_abastecer = 0
        total_costo_abastecimiento = 0.0

        for r, p in enumerate(productos):
            self.tabla_rep_aba.setRowHeight(r, 38)
            exist = p["existencia"]
            pr = p["stockMinimo"]
            maximo = p["stockMaximo"]
            cant_abastecer = max(maximo - exist, pr - exist)
            costo_abast = cant_abastecer * p["costo"]

            total_unidades_abastecer += cant_abastecer
            total_costo_abastecimiento += costo_abast

            self.tabla_rep_aba.setItem(r, 0, crear_celda(p["nombre"], negrita=True))
            self.tabla_rep_aba.setItem(r, 1, crear_celda(str(exist), Qt.AlignCenter))
            self.tabla_rep_aba.setItem(r, 2, crear_celda(p["unidadMedida"], Qt.AlignCenter))
            self.tabla_rep_aba.setItem(r, 3, crear_celda(str(pr), Qt.AlignCenter))
            self.tabla_rep_aba.setItem(r, 4, crear_celda(str(maximo), Qt.AlignCenter))
            self.tabla_rep_aba.setItem(r, 5, crear_celda(str(cant_abastecer), Qt.AlignCenter, negrita=True))
            self.tabla_rep_aba.setItem(r, 6, crear_celda(f"${costo_abast:,.2f}", Qt.AlignRight | Qt.AlignVCenter, negrita=True))

        ult = len(productos)
        self.tabla_rep_aba.setRowHeight(ult, 42)
        self.tabla_rep_aba.setItem(ult, 0, crear_celda("TOTAL ABASTECIMIENTO", negrita=True))
        for c_vacia in (1, 2, 3, 4):
            self.tabla_rep_aba.setItem(ult, c_vacia, crear_celda("-", Qt.AlignCenter))
        self.tabla_rep_aba.setItem(ult, 5, crear_celda(str(total_unidades_abastecer), Qt.AlignCenter, negrita=True))
        self.tabla_rep_aba.setItem(ult, 6, crear_celda(f"${total_costo_abastecimiento:,.2f}", Qt.AlignRight | Qt.AlignVCenter, negrita=True))

        self.lbl_totales_aba.setText(
            f"Productos por reabastecer: {len(productos)}   |   "
            f"Total Cantidad a Abastecer: {total_unidades_abastecer} unidades   |   "
            f"Costo Total del Abastecimiento: ${total_costo_abastecimiento:,.2f}"
        )

    def generar_todos_los_reportes(self):
        self.generar_reporte_recibos(mostrar_alerta_fechas=False)
        self.generar_reporte_inventario()
        self.generar_reporte_valor()
        self.generar_reporte_abastecimiento()


    # =========================================================================
    # MÓDULO 4: CLIENTES
    # =========================================================================
    def setup_page_clientes_cat(self):
        l=QVBoxLayout(self.page_clientes_cat); l.setContentsMargins(24,20,24,20); l.setSpacing(12)
        h=QHBoxLayout(); title=QLabel("Catálogo de clientes"); title.setStyleSheet("font-size:22px;font-weight:bold;")
        h.addWidget(title); h.addStretch()
        self.btn_cliente_nuevo=QPushButton("+ Agregar cliente"); self.btn_cliente_nuevo.setObjectName("btn_success")
        self.btn_cliente_nuevo.clicked.connect(self.ir_agregar_cliente); h.addWidget(self.btn_cliente_nuevo); l.addLayout(h)
        card=QFrame(); card.setProperty("class","card"); cl=QHBoxLayout(card)
        self.txt_buscar_cliente=QLineEdit(); self.txt_buscar_cliente.setPlaceholderText("Buscar por nombre, ID, teléfono o RFC…")
        self.txt_buscar_cliente.textChanged.connect(self.cargar_clientes)
        self.chk_clientes_inactivos=QCheckBox("Mostrar inactivos"); self.chk_clientes_inactivos.stateChanged.connect(self.cargar_clientes)
        btn_buscar_cliente=QPushButton("Buscar"); btn_buscar_cliente.setObjectName("btn_primary"); btn_buscar_cliente.clicked.connect(self.cargar_clientes)
        self.txt_buscar_cliente.returnPressed.connect(self.cargar_clientes)
        cl.addWidget(self.txt_buscar_cliente,1); cl.addWidget(self.chk_clientes_inactivos); cl.addWidget(btn_buscar_cliente); l.addWidget(card)
        self.tabla_clientes=QTableWidget(0,7); self.tabla_clientes.setHorizontalHeaderLabels(["ID","Nombre","Teléfono","Correo","RFC","Ciudad","Estado"])
        self.tabla_clientes.horizontalHeader().setSectionResizeMode(1,QHeaderView.Stretch)
        self.tabla_clientes.verticalHeader().setVisible(False); self.tabla_clientes.setEditTriggers(QTableWidget.NoEditTriggers)
        self.tabla_clientes.setSelectionBehavior(QTableWidget.SelectRows); self.tabla_clientes.setSelectionMode(QTableWidget.SingleSelection)
        l.addWidget(self.tabla_clientes,1)
        b=QHBoxLayout()
        for txt,fn,obj in [("Consultar",self.consultar_cliente_ui,"btn_secondary"),("Modificar",self.ir_modificar_cliente,"btn_primary"),("Eliminar",self.eliminar_cliente_ui,"btn_danger")]:
            bt=QPushButton(txt); bt.setObjectName(obj); bt.clicked.connect(fn); b.addWidget(bt)
        b.addStretch(); l.addLayout(b)

    def setup_page_clientes_form(self):
        l=QVBoxLayout(self.page_clientes_form); l.setContentsMargins(24,20,24,20)
        self.lbl_titulo_cliente=QLabel("Agregar cliente"); self.lbl_titulo_cliente.setStyleSheet("font-size:22px;font-weight:bold;")
        l.addWidget(self.lbl_titulo_cliente)
        card=QFrame(); card.setProperty("class","card"); g=QGridLayout(card); g.setHorizontalSpacing(18); g.setVerticalSpacing(10)
        campos=[("ID del cliente","idCliente"),("Nombre *","nombre"),("Domicilio / Calle *","domicilio"),("Número de calle *","numeroCalle"),
                ("Colonia *","colonia"),("Ciudad *","ciudad"),("Código Postal *","codigoPostal"),("Teléfono *","telefono"),
                ("Correo electrónico *","correoE"),("RFC *","rfc")]
        self.edits_cliente={}
        for i,(label,key) in enumerate(campos):
            r,c=divmod(i,2); g.addWidget(self._lbl(label),r*2,c)
            e=QLineEdit(); e.setPlaceholderText("(nuevo)" if key=="idCliente" else label.replace(" *",""))
            if key=="idCliente": e.setReadOnly(True)
            if key=="codigoPostal": e.setMaxLength(5); e.setValidator(QRegularExpressionValidator(QRegularExpression(r"[0-9]{0,5}"),e))
            if key=="telefono": e.setMaxLength(10); e.setValidator(QRegularExpressionValidator(QRegularExpression(r"[0-9]{0,10}"),e))
            if key=="rfc": e.setMaxLength(13); e.setValidator(QRegularExpressionValidator(QRegularExpression(r"[A-ZÑ&]{0,4}[0-9]{0,6}[A-Z0-9]{0,3}"),e))
            if key=="correoE": e.setMaxLength(100)
            self.edits_cliente[key]=e; g.addWidget(e,r*2+1,c)
        # Restaurar el tamaño y comportamiento original del formulario:
        # la tarjeta vuelve a ocupar el espacio disponible, sin ancho máximo.
        l.addWidget(card)
        b=QHBoxLayout(); b.addStretch()
        cancel=QPushButton("Cancelar"); cancel.setObjectName("btn_secondary"); cancel.clicked.connect(lambda:self.cambiar_seccion(10))
        save=QPushButton("Guardar cliente"); save.setObjectName("btn_success"); save.clicked.connect(self.guardar_cliente_bd)
        b.addWidget(cancel); b.addWidget(save); l.addLayout(b)

    def setup_page_clientes_ver(self):
        l=QVBoxLayout(self.page_clientes_ver)
        l.setContentsMargins(24,20,24,20)
        l.setSpacing(12)

        title=QLabel("Consultar cliente")
        title.setStyleSheet("font-size:22px;font-weight:bold;")
        l.addWidget(title)

        # Distribución siguiendo la pantalla "Consultar producto":
        # tarjeta de datos a la izquierda e historial a la derecha,
        # ambas ocupando la altura disponible.
        body=QHBoxLayout()
        body.setSpacing(18)

        card_datos=QFrame()
        card_datos.setProperty("class","card")
        datos_layout=QVBoxLayout(card_datos)
        datos_layout.setContentsMargins(20,18,20,18)
        datos_layout.setSpacing(12)

        self.lbl_nombre_cliente_consulta=QLabel("Datos del cliente")
        self.lbl_nombre_cliente_consulta.setStyleSheet("font-size:18px;font-weight:bold;")
        datos_layout.addWidget(self.lbl_nombre_cliente_consulta)

        form=QGridLayout()
        form.setContentsMargins(0,0,0,0)
        form.setHorizontalSpacing(18)
        form.setVerticalSpacing(10)
        self.lbl_datos_cliente={}
        campos=[
            ("ID del cliente","idCliente"),
            ("Nombre","nombre"),
            ("Domicilio / Calle","domicilio"),
            ("Número de calle","numeroCalle"),
            ("Colonia","colonia"),
            ("Ciudad","ciudad"),
            ("Código Postal","codigoPostal"),
            ("Teléfono","telefono"),
            ("Correo electrónico","correoE"),
            ("RFC","rfc"),
            ("Estado","estado")
        ]
        for fila,(label,key) in enumerate(campos):
            etiqueta=self._lbl(label + ":")
            valor=QLabel("—")
            valor.setWordWrap(True)
            valor.setStyleSheet("font-weight:600;")
            self.lbl_datos_cliente[key]=valor
            form.addWidget(etiqueta,fila,0,Qt.AlignVCenter)
            form.addWidget(valor,fila,1,Qt.AlignVCenter)
        form.setColumnStretch(0,1)
        form.setColumnStretch(1,1)
        datos_layout.addLayout(form)
        datos_layout.addStretch(1)
        body.addWidget(card_datos,1)

        card_historial=QFrame()
        card_historial.setProperty("class","card")
        historial_layout=QVBoxLayout(card_historial)
        historial_layout.setContentsMargins(20,18,20,18)
        historial_layout.setSpacing(10)
        subtitulo=QLabel("Ventas recientes")
        subtitulo.setStyleSheet("font-size:17px;font-weight:bold;")
        historial_layout.addWidget(subtitulo)

        self.tabla_historial_cliente=QTableWidget(0,5)
        self.tabla_historial_cliente.setHorizontalHeaderLabels(
            ["Folio","Fecha","Método de pago","Total","Estado"]
        )
        self.tabla_historial_cliente.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.tabla_historial_cliente.verticalHeader().setVisible(False)
        self.tabla_historial_cliente.setEditTriggers(QTableWidget.NoEditTriggers)
        self.tabla_historial_cliente.setSelectionBehavior(QTableWidget.SelectRows)
        self.tabla_historial_cliente.setSelectionMode(QTableWidget.SingleSelection)
        historial_layout.addWidget(self.tabla_historial_cliente,1)
        body.addWidget(card_historial,1)

        l.addLayout(body,1)

        acciones=QHBoxLayout()
        acciones.addStretch(1)
        back=QPushButton("Volver")
        back.setObjectName("btn_secondary")
        back.clicked.connect(lambda:self.cambiar_seccion(10))
        self.btn_modificar_cliente_consulta=QPushButton("Modificar")
        self.btn_modificar_cliente_consulta.setObjectName("btn_primary")
        self.btn_modificar_cliente_consulta.clicked.connect(self.modificar_cliente_consultado)
        acciones.addWidget(back)
        acciones.addWidget(self.btn_modificar_cliente_consulta)
        l.addLayout(acciones)

    def _fila_seleccionada(self, tabla, titulo):
        fila=tabla.currentRow()
        if fila < 0:
            QMessageBox.warning(self,titulo,"Seleccione un registro de la tabla.")
            return None
        return fila

    def consultar_producto_seleccionado(self):
        r=self._fila_seleccionada(self.tabla_cat,"Productos")
        if r is not None: self.ir_consultar_producto(self.tabla_cat.item(r,0).text())

    def modificar_producto_seleccionado(self):
        r=self._fila_seleccionada(self.tabla_cat,"Productos")
        if r is not None: self.ir_modificar_producto(self.tabla_cat.item(r,0).text())

    def eliminar_producto_seleccionado(self):
        r=self._fila_seleccionada(self.tabla_cat,"Productos")
        if r is not None:
            codigo=self.tabla_cat.item(r,0).text()
            prod=next((normalizar_producto(x) for x in self.repo.buscar_productos(codigo) if str(x.get("codigo",x.get("CODIGO","")))==codigo),None)
            if prod: self.abrir_eliminar_producto(prod)

    def consultar_proveedor_seleccionado(self):
        r=self._fila_seleccionada(self.tabla_prov,"Proveedores")
        if r is not None: self.ir_consultar_proveedor(int(self.tabla_prov.item(r,0).text()))

    def modificar_proveedor_seleccionado(self):
        r=self._fila_seleccionada(self.tabla_prov,"Proveedores")
        if r is not None: self.ir_modificar_proveedor(int(self.tabla_prov.item(r,0).text()))

    def eliminar_proveedor_seleccionado(self):
        r=self._fila_seleccionada(self.tabla_prov,"Proveedores")
        if r is not None:
            ident=int(self.tabla_prov.item(r,0).text())
            prov=next((normalizar_proveedor(x) for x in self.repo.buscar_proveedores(str(ident),incluir_inactivos=True) if int(x.get("idProveedor",x.get("idproveedor",0)) or 0)==ident),None)
            if prov: self.abrir_eliminar_proveedor(prov)

    def consultar_nota_seleccionada(self):
        r=self._fila_seleccionada(self.tabla_notas,"Notas de recibo")
        if r is not None: self.ir_consultar_nota_recibo(int(self.tabla_notas.item(r,0).text().replace("NR-","")))

    def modificar_nota_seleccionada(self):
        r=self._fila_seleccionada(self.tabla_notas,"Notas de recibo")
        if r is not None: self.ir_modificar_nota_recibo(int(self.tabla_notas.item(r,0).text().replace("NR-","")))

    def cancelar_nota_seleccionada(self):
        r=self._fila_seleccionada(self.tabla_notas,"Notas de recibo")
        if r is not None:
            folio=int(self.tabla_notas.item(r,0).text().replace("NR-",""))
            estado=self.tabla_notas.item(r,6).text() if self.tabla_notas.item(r,6) else ""
            if estado and "Cancelada" in estado:
                QMessageBox.information(self,"Notas de recibo","La nota seleccionada ya está cancelada.")
                return
            self.cancelar_nota_recibo_ui(folio)

    def cargar_clientes(self,*_):
        rows=self.repo.listar_clientes(self.txt_buscar_cliente.text(),self.chk_clientes_inactivos.isChecked())
        self.tabla_clientes.setRowCount(len(rows))
        for i,r in enumerate(rows):
            vals=[r["idCliente"],r["nombre"],r["telefono"],r.get("correoE") or "",r.get("rfc") or "",r.get("ciudad") or "", "Activo" if r["estado"] else "Inactivo"]
            for j,v in enumerate(vals): self.tabla_clientes.setItem(i,j,crear_celda(v,Qt.AlignCenter if j in (0,2,6) else Qt.AlignLeft|Qt.AlignVCenter))
            self.tabla_clientes.setRowHeight(i,34)

    def _cliente_seleccionado(self):
        row=self.tabla_clientes.currentRow()
        if row<0: QMessageBox.warning(self,"Clientes","Seleccione un cliente."); return None
        return self.repo.obtener_cliente(int(self.tabla_clientes.item(row,0).text()))

    def ir_agregar_cliente(self):
        self.modo_formulario_cliente="AGREGAR"; self.cliente_actual=None
        for k,e in self.edits_cliente.items(): e.clear()
        self.edits_cliente["idCliente"].setText("(nuevo)")
        self.lbl_titulo_cliente.setText("Agregar cliente"); self.cambiar_seccion(11)

    def ir_modificar_cliente(self, cliente=None):
        c = cliente or self._cliente_seleccionado()
        if not c:return
        if not c["estado"]: QMessageBox.warning(self,"Cliente","No se puede modificar un cliente inactivo."); return
        self.modo_formulario_cliente="MODIFICAR"; self.cliente_actual=c
        for k,e in self.edits_cliente.items(): e.setText("(nuevo)" if k=="idCliente" else str(c.get(k) or ""))
        self.edits_cliente["idCliente"].setText(str(c["idCliente"]))
        self.lbl_titulo_cliente.setText(f"Modificar cliente · ID {c['idCliente']}"); self.cambiar_seccion(11)

    def guardar_cliente_bd(self):
        d={k:e.text().strip() for k,e in self.edits_cliente.items() if k!="idCliente"}
        try:
            if not d["nombre"]: raise ValueError("El nombre es obligatorio.")
            if not all(d[k] for k in ("domicilio","numeroCalle","colonia","ciudad","codigoPostal","telefono","correoE","rfc")):
                raise ValueError("Todos los campos marcados con * son obligatorios.")
            if "$" in d["correoE"]: raise ValueError("El correo electrónico no puede contener el símbolo $.")
            if not __import__("re").fullmatch(r"[^\s@]+@[^\s@]+\.[^\s@]+",d["correoE"]):
                raise ValueError("Correo inválido. Use usuario@dominio.ext.")
            if not __import__("re").fullmatch(r"[0-9]{5}",d["codigoPostal"]): raise ValueError("El Código Postal debe contener 5 dígitos.")
            if not __import__("re").fullmatch(r"[0-9]{10}",d["telefono"]): raise ValueError("El teléfono debe contener exactamente 10 dígitos.")
            if not __import__("re").fullmatch(r"[A-ZÑ&]{4}\d{6}[A-Z0-9]{3}",d["rfc"]): raise ValueError("El RFC debe tener 13 caracteres y respetar el formato indicado.")
            yy,mm,dd=int(d["rfc"][4:6]),int(d["rfc"][6:8]),int(d["rfc"][8:10])
            from datetime import date
            anio=2000+yy if yy<=datetime.now().year%100 else 1900+yy
            try: date(anio,mm,dd)
            except ValueError: raise ValueError("La fecha de nacimiento en el RFC no es válida (YYMMDD).")
            if self.modo_formulario_cliente=="AGREGAR": nuevo=self.repo.agregar_cliente(d); msg=f"Cliente registrado con ID {nuevo}."
            else: self.repo.modificar_cliente(self.cliente_actual["idCliente"],d); msg="Cliente actualizado correctamente."
            QMessageBox.information(self,"Clientes",msg); self.cambiar_seccion(10)
        except Exception as e: QMessageBox.critical(self,"No se pudo guardar",str(e))

    def consultar_cliente_ui(self):
        c=self._cliente_seleccionado()
        if not c:return
        self.cliente_actual = c
        self.lbl_nombre_cliente_consulta.setText(str(c.get("nombre") or "Datos del cliente"))
        for k,w in self.lbl_datos_cliente.items():
            v=("Activo" if c["estado"] else "Inactivo") if k=="estado" else c.get(k,"")
            w.setText(str(v or "—"))
        rows=self.repo.historial_compras_cliente(c["idCliente"]); self.tabla_historial_cliente.setRowCount(len(rows))
        for i,r in enumerate(rows):
            vals=[r["folio"],r["fecha"],r["metodoPago"],f"${r['total']:,.2f}",r["estado"]]
            for j,v in enumerate(vals): self.tabla_historial_cliente.setItem(i,j,crear_celda(v))
        self.cambiar_seccion(12)

    def modificar_cliente_consultado(self):
        if not self.cliente_actual:
            QMessageBox.warning(self, "Clientes", "No hay un cliente seleccionado para modificar.")
            return
        self.ir_modificar_cliente(self.cliente_actual)

    def eliminar_cliente_ui(self):
        c=self._cliente_seleccionado()
        if not c:return
        dlg=DialogEliminarCliente(c,self)
        if dlg.exec()==QDialog.Accepted:
            try:self.repo.eliminar_cliente_logico(c["idCliente"]); self.cargar_clientes()
            except Exception as e:QMessageBox.critical(self,"Clientes",str(e))

    # =========================================================================
    # MÓDULO 5: PUNTO DE VENTA
    # =========================================================================
    def setup_page_punto_venta(self):
        l=QVBoxLayout(self.page_punto_venta); l.setContentsMargins(24,20,24,20)
        self.lbl_titulo_venta=QLabel("Punto de venta"); self.lbl_titulo_venta.setStyleSheet("font-size:22px;font-weight:bold;"); l.addWidget(self.lbl_titulo_venta)
        row=QHBoxLayout()
        left=QVBoxLayout(); cardp=QFrame(); cardp.setProperty("class","card"); cp=QVBoxLayout(cardp)
        cp.addWidget(self._lbl("Buscar producto por nombre o SKU"))
        self.txt_producto_venta=QLineEdit(); self.txt_producto_venta.setPlaceholderText("Escriba SKU o nombre del producto…")
        self.txt_producto_venta.returnPressed.connect(self.agregar_producto_venta); self.txt_producto_venta.textChanged.connect(self._filtrar_coincidencias_venta); cp.addWidget(self.txt_producto_venta)
        self.cmb_producto_venta=QComboBox(); self.cmb_producto_venta.setEditable(False); self.cmb_producto_venta.setPlaceholderText("Coincidencias")
        cp.addWidget(self.cmb_producto_venta)
        self.spin_cantidad_venta=QSpinBox(); self.spin_cantidad_venta.setMinimum(1); self.spin_cantidad_venta.setMaximum(999999)
        cp.addWidget(self._lbl("Cantidad")); cp.addWidget(self.spin_cantidad_venta)
        add=QPushButton("Agregar producto"); add.setObjectName("btn_primary"); add.clicked.connect(self.agregar_producto_venta); cp.addWidget(add)
        left.addWidget(cardp)
        cardc=QFrame(); cardc.setProperty("class","card"); cc=QVBoxLayout(cardc); cc.addWidget(self._lbl("Cliente (opcional)"))
        self.txt_cliente_venta=QLineEdit(); self.txt_cliente_venta.setPlaceholderText("Público en general o buscar cliente activo…"); cc.addWidget(self.txt_cliente_venta)
        self.lbl_cliente_venta=QLabel("Cliente: Público en general"); cc.addWidget(self.lbl_cliente_venta); left.addWidget(cardc); row.addLayout(left,1)
        right=QVBoxLayout(); self.tabla_venta=QTableWidget(0,6); self.tabla_venta.setHorizontalHeaderLabels(["SKU","Producto","Cantidad","Precio unitario","Subtotal","Acción"])
        self.tabla_venta.horizontalHeader().setSectionResizeMode(1,QHeaderView.Stretch); self.tabla_venta.setEditTriggers(QTableWidget.NoEditTriggers); right.addWidget(self.tabla_venta,1)
        cardpay=QFrame(); cardpay.setProperty("class","card"); pay=QGridLayout(cardpay)
        pay.addWidget(self._lbl("Método de pago"),0,0); self.cmb_metodo_pago=QComboBox(); self.cmb_metodo_pago.addItems(["Efectivo","Tarjeta"]); self.cmb_metodo_pago.currentTextChanged.connect(self._recalcular_cobro); pay.addWidget(self.cmb_metodo_pago,0,1)
        pay.addWidget(self._lbl("Efectivo recibido"),1,0); self.spn_efectivo=QDoubleSpinBox(); self.spn_efectivo.setMaximum(999999999); self.spn_efectivo.setDecimals(2); self.spn_efectivo.valueChanged.connect(self._recalcular_cobro); pay.addWidget(self.spn_efectivo,1,1)
        self.lbl_total_venta=QLabel("$0.00"); self.lbl_total_venta.setStyleSheet("font-size:22px;font-weight:bold;color:#3730A3;"); pay.addWidget(QLabel("TOTAL A PAGAR"),2,0); pay.addWidget(self.lbl_total_venta,2,1)
        self.lbl_cambio_venta=QLabel("Cambio a entregar: $0.00"); self.lbl_cambio_venta.setStyleSheet("font-size:22px;font-weight:bold;color:#0D9488;"); pay.addWidget(self.lbl_cambio_venta,3,0,1,2)
        self.lbl_total_letras_venta=QLabel("SON: CERO PESOS 00/100 M.N."); self.lbl_total_letras_venta.setWordWrap(True); pay.addWidget(self.lbl_total_letras_venta,4,0,1,2)
        right.addWidget(cardpay); row.addLayout(right,3); l.addLayout(row,1)
        b=QHBoxLayout(); self.btn_cobrar=QPushButton("Cobrar y generar ticket"); self.btn_cobrar.setObjectName("btn_success"); self.btn_cobrar.clicked.connect(self.cobrar_venta)
        back=QPushButton("Volver"); back.setObjectName("btn_secondary"); back.clicked.connect(lambda:self.cambiar_seccion(0))
        b.addWidget(back); b.addStretch(); b.addWidget(self.btn_cobrar); l.addLayout(b)

    def _cargar_completers_venta(self):
        self.mapa_clientes_venta={}
        clientes=self.repo.listar_clientes()
        from PySide6.QtCore import QStringListModel
        comp=QCompleter([f"{c['idCliente']} - {c['nombre']} ({c['telefono']})" for c in clientes],self.txt_cliente_venta)
        comp.setCaseSensitivity(Qt.CaseInsensitive); comp.setFilterMode(Qt.MatchContains); self.txt_cliente_venta.setCompleter(comp)
        self.mapa_clientes_venta={f"{c['idCliente']} - {c['nombre']} ({c['telefono']})":c["idCliente"] for c in clientes}
        prods=self.repo.productos_para_venta()
        self.mapa_productos_venta={f"{p['codigo']} - {p['nombre']} (Stock: {p['existencia']})":p for p in prods}
        self.mapa_productos_por_codigo={p["codigo"]:p for p in prods}
        self._filtrar_coincidencias_venta(self.txt_producto_venta.text())
        comp_prod=QCompleter(list(self.mapa_productos_venta.keys()),self.txt_producto_venta)
        comp_prod.setCaseSensitivity(Qt.CaseInsensitive)
        comp_prod.setFilterMode(Qt.MatchContains)  # busca en cualquier parte: nombre o código
        self.txt_producto_venta.setCompleter(comp_prod)
        if not self.modo_ticket_edicion:self.txt_cliente_venta.setEnabled(True)

    def _coincidencias_venta(self, texto: str) -> list:
        """Productos cuyo texto contiene TODAS las palabras escritas (sin acentos ni mayúsculas)."""
        palabras = normalizar_texto(texto).split()
        return [k for k in getattr(self, "mapa_productos_venta", {})
                if all(w in normalizar_texto(k) for w in palabras)]

    def _filtrar_coincidencias_venta(self, texto: str):
        if not hasattr(self, "mapa_productos_venta"):
            return
        self.cmb_producto_venta.blockSignals(True)
        self.cmb_producto_venta.clear()
        for k in self._coincidencias_venta(texto):
            self.cmb_producto_venta.addItem(k)
        self.cmb_producto_venta.setCurrentIndex(-1)
        self.cmb_producto_venta.blockSignals(False)

    def nueva_venta(self):
        if self.renglones_venta and self.stacked.currentIndex() != 15:
            respuesta = QMessageBox.question(self, "Nueva venta", "¿Descartar los productos de la venta actual?", QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if respuesta != QMessageBox.Yes:
                return
        self.modo_ticket_edicion=False; self.folio_ticket_editando=None; self.renglones_venta={}
        self.lbl_titulo_venta.setText("Punto de venta · Nueva venta")
        self.txt_cliente_venta.clear(); self.txt_cliente_venta.setEnabled(True)
        self.cmb_metodo_pago.setEnabled(True); self.spn_efectivo.setValue(0)
        self.cambiar_seccion(13); self._refrescar_venta()
        self.txt_producto_venta.setFocus()  # listo para escribir o escanear

    def agregar_producto_venta(self):
        texto=self.txt_producto_venta.text().strip()
        clave=texto or self.cmb_producto_venta.currentText()
        producto=self.mapa_productos_venta.get(clave)
        if producto is None:
            cod=clave.split(" - ",1)[0].strip()
            producto=self.mapa_productos_por_codigo.get(cod)
        if producto is None and texto:
            # Búsqueda por nombre: si hay una sola coincidencia se agrega directo
            coincidencias=self._coincidencias_venta(texto)
            if len(coincidencias)==1:
                producto=self.mapa_productos_venta[coincidencias[0]]
            elif len(coincidencias)>1:
                QMessageBox.information(self,"Punto de venta",
                    f"Hay {len(coincidencias)} productos que coinciden con «{texto}».\n"
                    "Elija uno de la lista de sugerencias o escriba más palabras.")
                return
        if producto is None:
            producto=next((p for p in self.repo.productos_para_venta(texto) if p["codigo"]==texto),None)
        if not producto:
            QMessageBox.warning(self,"Punto de venta","Seleccione un producto activo del catálogo."); return
        cantidad=self.spin_cantidad_venta.value()
        existente=self.renglones_venta.get(producto["codigo"],{})
        actual=existente.get("cantidad",0)
        stock_permitido=existente.get("stock",producto["existencia"])
        if actual+cantidad>stock_permitido:
            QMessageBox.warning(self,"Stock insuficiente",f"Disponible para esta operación: {stock_permitido}. Ya agregado: {actual}."); return
        self.renglones_venta[producto["codigo"]]={"codigo":producto["codigo"],"nombre":producto["nombre"],"cantidad":actual+cantidad,"precio":existente.get("precio",producto["precio"]),"stock":stock_permitido}
        self.txt_producto_venta.clear(); self.spin_cantidad_venta.setValue(1); self._refrescar_venta()

    def _refrescar_venta(self):
        self.tabla_venta.setRowCount(len(self.renglones_venta))
        for i,(cod,d) in enumerate(self.renglones_venta.items()):
            sub=d["cantidad"]*d["precio"]; vals=[cod,d["nombre"],d["cantidad"],f"${d['precio']:,.2f}",f"${sub:,.2f}"]
            for j,v in enumerate(vals):self.tabla_venta.setItem(i,j,crear_celda(v,Qt.AlignCenter if j in (0,2) else Qt.AlignLeft|Qt.AlignVCenter))
            acciones=QWidget(); al=QHBoxLayout(acciones); al.setContentsMargins(2,2,2,2); al.setSpacing(3)
            menos=QPushButton("−"); menos.setMaximumWidth(30); menos.clicked.connect(lambda _,c=cod:self.ajustar_cantidad_venta(c,-1))
            mas=QPushButton("+"); mas.setMaximumWidth(30); mas.clicked.connect(lambda _,c=cod:self.ajustar_cantidad_venta(c,1))
            borrar=QPushButton("×"); borrar.setMaximumWidth(30); borrar.setObjectName("btn_danger"); borrar.clicked.connect(lambda _,c=cod:self.quitar_producto_venta(c))
            al.addWidget(menos); al.addWidget(mas); al.addWidget(borrar); self.tabla_venta.setCellWidget(i,5,acciones)
        self._recalcular_cobro()

    def ajustar_cantidad_venta(self,codigo,delta):
        d=self.renglones_venta.get(codigo)
        if not d:return
        nueva=d["cantidad"]+delta
        if nueva<=0:self.renglones_venta.pop(codigo,None)
        elif nueva>d["stock"]:
            QMessageBox.warning(self,"Stock insuficiente",f"Existencia disponible: {d['stock']}.")
            return
        else:d["cantidad"]=nueva
        self._refrescar_venta()

    def quitar_producto_venta(self,codigo):
        self.renglones_venta.pop(codigo,None); self._refrescar_venta()

    def _recalcular_cobro(self,*_):
        total=sum(d["cantidad"]*d["precio"] for d in self.renglones_venta.values()) if hasattr(self,"renglones_venta") else 0
        self.lbl_total_venta.setText(f"${total:,.2f}"); self.lbl_total_letras_venta.setText(numero_a_letras_mx(total))
        if self.cmb_metodo_pago.currentText()=="Tarjeta":
            self.spn_efectivo.setEnabled(False); self.lbl_cambio_venta.setText("Cambio a entregar: $0.00")
        else:
            self.spn_efectivo.setEnabled(True); self.lbl_cambio_venta.setText(f"Cambio a entregar: ${max(self.spn_efectivo.value()-total,0):,.2f}")

    def cobrar_venta(self):
        if not self.renglones_venta: QMessageBox.warning(self,"Punto de venta","Agregue productos antes de cobrar."); return
        if self.modo_ticket_edicion:
            try:
                self.repo.modificar_ticket_detalle(self.folio_ticket_editando,list(self.renglones_venta.values()))
                QMessageBox.information(self,"Ticket","Ticket actualizado y existencias ajustadas.")
                self.modo_ticket_edicion=False
                self.mostrar_comprobante(self.folio_ticket_editando)
            except Exception as e:QMessageBox.critical(self,"No se pudo modificar",str(e))
            return
        texto=self.txt_cliente_venta.text().strip()
        id_cliente=self.mapa_clientes_venta.get(texto)
        if texto and id_cliente is None:
            QMessageBox.warning(self,"Cliente","Seleccione un cliente de la lista o deje el campo vacío para Público en general."); return
        try:
            folio=self.repo.registrar_ticket(id_cliente,self.cmb_metodo_pago.currentText(),self.spn_efectivo.value(),list(self.renglones_venta.values()))
            self.mostrar_comprobante(folio)
        except Exception as e:QMessageBox.critical(self,"No se pudo cobrar",str(e))

    # =========================================================================
    # TICKETS: CONSULTA, MODIFICACIÓN Y CANCELACIÓN
    # =========================================================================
    def setup_page_tickets(self):
        l=QVBoxLayout(self.page_tickets); l.setContentsMargins(24,20,24,20)
        title=QLabel("Gestión y consulta de tickets"); title.setStyleSheet("font-size:22px;font-weight:bold;"); l.addWidget(title)
        f=QFrame(); f.setProperty("class","card"); g=QGridLayout(f)
        g.addWidget(self._lbl("Folio"),0,0); self.txt_filtro_folio=QLineEdit(); self.txt_filtro_folio.setPlaceholderText("155 o V-000155"); g.addWidget(self.txt_filtro_folio,0,1)
        g.addWidget(self._lbl("Fecha desde"),0,2); self.dt_ticket_desde=QDateEdit(QDate.currentDate().addMonths(-1)); configurar_date_edit_claro(self.dt_ticket_desde); g.addWidget(self.dt_ticket_desde,0,3)
        g.addWidget(self._lbl("Fecha hasta"),0,4); self.dt_ticket_hasta=QDateEdit(QDate.currentDate()); configurar_date_edit_claro(self.dt_ticket_hasta); g.addWidget(self.dt_ticket_hasta,0,5)
        g.addWidget(self._lbl("Método de pago"),1,0); self.cmb_filtro_metodo=QComboBox(); self.cmb_filtro_metodo.addItems(["Todos","Efectivo","Tarjeta"]); g.addWidget(self.cmb_filtro_metodo,1,1)
        b=QPushButton("Buscar"); b.setObjectName("btn_primary"); b.clicked.connect(self.cargar_tickets); g.addWidget(b,1,2)
        l.addWidget(f)
        self.tabla_tickets=QTableWidget(0,6); self.tabla_tickets.setHorizontalHeaderLabels(["Folio","Fecha","Cliente","Método de pago","Total","Estado"])
        self.tabla_tickets.horizontalHeader().setSectionResizeMode(2,QHeaderView.Stretch); self.tabla_tickets.setSelectionBehavior(QTableWidget.SelectRows); self.tabla_tickets.setEditTriggers(QTableWidget.NoEditTriggers)
        self.tabla_tickets.itemSelectionChanged.connect(self._mostrar_detalle_ticket_seleccionado); l.addWidget(self.tabla_tickets,1)
        self.lbl_detalle_ticket=QLabel("Seleccione un ticket para consultar su detalle."); self.lbl_detalle_ticket.setWordWrap(True); l.addWidget(self.lbl_detalle_ticket)
        self.tabla_detalle_ticket=QTableWidget(0,5); self.tabla_detalle_ticket.setHorizontalHeaderLabels(["SKU","Nombre","Cantidad","Precio unitario","Subtotal"]); self.tabla_detalle_ticket.horizontalHeader().setSectionResizeMode(1,QHeaderView.Stretch); l.addWidget(self.tabla_detalle_ticket,1)
        self.lbl_total_ticket_consulta=QLabel(""); self.lbl_total_ticket_consulta.setStyleSheet("font-size:16px;font-weight:bold;"); l.addWidget(self.lbl_total_ticket_consulta)
        b=QHBoxLayout()
        bt=QPushButton("Ver comprobante / Reimprimir"); bt.setObjectName("btn_secondary"); bt.clicked.connect(self.ver_comprobante_ticket); b.addWidget(bt)
        self.btn_modificar_ticket=QPushButton("Modificar ticket"); self.btn_modificar_ticket.setObjectName("btn_primary"); self.btn_modificar_ticket.clicked.connect(self.modificar_ticket_ui); b.addWidget(self.btn_modificar_ticket)
        self.btn_cancelar_ticket=QPushButton("Cancelar ticket"); self.btn_cancelar_ticket.setObjectName("btn_danger"); self.btn_cancelar_ticket.clicked.connect(self.cancelar_ticket_ui); b.addWidget(self.btn_cancelar_ticket)
        self.btn_modificar_ticket.setEnabled(False); self.btn_cancelar_ticket.setEnabled(False)
        b.addStretch(); l.addLayout(b)

    def cargar_tickets(self,*_):
        desde=self.dt_ticket_desde.date().toString("yyyy-MM-dd"); hasta=self.dt_ticket_hasta.date().toString("yyyy-MM-dd")
        if desde>hasta: QMessageBox.warning(self,"Fechas","La fecha inicial no puede ser mayor que la fecha final."); return
        self.tickets_actuales=self.repo.listar_tickets(self.txt_filtro_folio.text(),desde,hasta,self.cmb_filtro_metodo.currentText(),True)
        self.tabla_tickets.setRowCount(len(self.tickets_actuales))
        for i,r in enumerate(self.tickets_actuales):
            vals=[r["folio"],r["fecha"],r["cliente"],r["metodoPago"],f"${r['total']:,.2f}","Activo" if r["estado"] else "Cancelado"]
            for j,v in enumerate(vals):self.tabla_tickets.setItem(i,j,crear_celda(v))
        self.tabla_detalle_ticket.setRowCount(0); self.lbl_total_ticket_consulta.setText("")

    def _folio_ticket_seleccionado(self):
        row=self.tabla_tickets.currentRow()
        return int(self.tabla_tickets.item(row,0).text()) if row>=0 else None

    def _mostrar_detalle_ticket_seleccionado(self):
        folio=self._folio_ticket_seleccionado()
        if not folio:
            self.btn_modificar_ticket.setEnabled(False); self.btn_cancelar_ticket.setEnabled(False)
            return
        t=self.repo.obtener_ticket_completo(folio)
        if not t:return
        self.btn_modificar_ticket.setEnabled(bool(t["estado"])); self.btn_cancelar_ticket.setEnabled(bool(t["estado"]))
        self.ticket_seleccionado=t
        self.lbl_detalle_ticket.setText(f"Folio V-{folio:06d} · {t['fecha']} · Cliente: {t['cliente']} · Método: {t['metodoPago']} · Estado: {'Activo/Vigente' if t['estado'] else 'Cancelado'}")
        self.tabla_detalle_ticket.setRowCount(len(t["detalles"]))
        for i,d in enumerate(t["detalles"]):
            for j,v in enumerate([d["codigo"],d["nombre"],d["cantidad"],f"${d['precio']:,.2f}",f"${d['subtotal']:,.2f}"]):
                self.tabla_detalle_ticket.setItem(i,j,crear_celda(v))
        self.lbl_total_ticket_consulta.setText(f"TOTAL: ${t['total']:,.2f} · {numero_a_letras_mx(t['total'])}")

    def ver_comprobante_ticket(self):
        folio=self._folio_ticket_seleccionado()
        if folio:self.mostrar_comprobante(folio)

    def modificar_ticket_ui(self):
        folio=self._folio_ticket_seleccionado()
        if not folio:return
        t=self.repo.obtener_ticket_completo(folio)
        if not t or not t["estado"]: QMessageBox.warning(self,"Ticket","Solo se pueden modificar tickets activos."); return
        self.modo_ticket_edicion=True; self.folio_ticket_editando=folio; self.renglones_venta={}
        for d in t["detalles"]:
            p=next((x for x in self.repo.productos_para_venta(d["codigo"])),None)
            self.renglones_venta[d["codigo"]]={"codigo":d["codigo"],"nombre":d["nombre"],"cantidad":d["cantidad"],"precio":d["precio"],"stock":(p["existencia"] if p else 0)+d["cantidad"]}
        self.lbl_titulo_venta.setText(f"Modificar ticket V-{folio:06d} · Folio, fecha y cliente bloqueados")
        self.txt_cliente_venta.setText(t["cliente"]); self.txt_cliente_venta.setEnabled(False); self.cmb_metodo_pago.setCurrentText(t["metodoPago"]); self.cmb_metodo_pago.setEnabled(False)
        self._cargar_completers_venta(); self._refrescar_venta(); self.cambiar_seccion(13)

    def cancelar_ticket_ui(self):
        folio=self._folio_ticket_seleccionado()
        if not folio:return
        t=self.repo.obtener_ticket_completo(folio)
        if not t or not t["estado"]: QMessageBox.warning(self,"Ticket","El ticket ya está cancelado."); return
        motivo,ok=QInputDialog.getText(self,"Cancelar ticket","Motivo de cancelación (obligatorio):")
        if not ok:return
        if not motivo.strip(): QMessageBox.warning(self,"Cancelación","Debe indicar el motivo de cancelación."); return
        confirmar=QMessageBox.question(self,"Confirmar cancelación",f"¿Cancelar el ticket V-{folio:06d} y reintegrar su stock?",QMessageBox.Yes|QMessageBox.No,QMessageBox.No)
        if confirmar!=QMessageBox.Yes:return
        try:self.repo.cancelar_ticket(folio,motivo); QMessageBox.information(self,"Ticket","Ticket cancelado y stock reintegrado."); self.cargar_tickets()
        except Exception as e:QMessageBox.critical(self,"Cancelación",str(e))

    # =========================================================================
    # COMPROBANTE DE VENTA
    # =========================================================================
    def setup_page_comprobante(self):
        l=QVBoxLayout(self.page_comprobante); l.setContentsMargins(24,20,24,20)
        self.lbl_banner_comprobante=QLabel("✓ Venta registrada correctamente"); self.lbl_banner_comprobante.setStyleSheet("background:#CCFBF1;color:#115E59;padding:14px;font-size:17px;font-weight:bold;border-radius:8px;"); l.addWidget(self.lbl_banner_comprobante)
        body=QHBoxLayout()
        self.lbl_ticket_impreso=QLabel(); self.lbl_ticket_impreso.setTextFormat(Qt.RichText); self.lbl_ticket_impreso.setWordWrap(True); self.lbl_ticket_impreso.setAlignment(Qt.AlignTop)
        self.lbl_ticket_impreso.setStyleSheet("background:white;border:1px solid #C7D2E8;border-radius:8px;padding:24px;")
        body.addWidget(self.lbl_ticket_impreso,2)
        card=QFrame(); card.setProperty("class","card"); cl=QVBoxLayout(card); cl.addWidget(QLabel("Ticket generado"))
        self.lbl_resumen_comprobante=QLabel(); self.lbl_resumen_comprobante.setWordWrap(True); cl.addWidget(self.lbl_resumen_comprobante)
        for txt,fn in [("Guardar PDF",self.guardar_pdf_ticket),("Imprimir",self.imprimir_ticket),("Nueva venta",self.nueva_venta)]:
            bt=QPushButton(txt); bt.setObjectName("btn_primary" if txt!="Nueva venta" else "btn_success"); bt.clicked.connect(fn); cl.addWidget(bt)
        cl.addStretch(); body.addWidget(card,1); l.addLayout(body,1)
        back=QPushButton("Ir a tickets"); back.setObjectName("btn_secondary"); back.clicked.connect(lambda:self.cambiar_seccion(14)); l.addWidget(back)

    def mostrar_comprobante(self,folio):
        t=self.repo.obtener_ticket_completo(folio)
        if not t: QMessageBox.warning(self,"Ticket","No se encontró el ticket."); return
        self.ticket_comprobante_actual=t
        filas="".join(f"<tr><td>{d['nombre']}<br><small>{d['codigo']} · {d['cantidad']} x ${d['precio']:,.2f}</small></td><td align='right'>${d['subtotal']:,.2f}</td></tr>" for d in t["detalles"])
        html=f"""<div style='font-family:Arial;color:#0F172A;text-align:center;'><h2>LOS EXPERTOS</h2><b>COMPROBANTE DE VENTA</b><hr>
        <p>Folio: V-{folio:06d}<br>Fecha: {t['fecha']}<br>Cliente: {t['cliente']}</p>
        <table width='100%' cellspacing='6'>{filas}</table><hr><h2>TOTAL: ${t['total']:,.2f}</h2>
        <b>{numero_a_letras_mx(t['total'])}</b><p>Método de pago: {t['metodoPago']}<br>Efectivo recibido: ${t['total']+t['cambio']:,.2f}<br>Cambio: ${t['cambio']:,.2f}</p><hr><small>¡Gracias por su compra!</small></div>"""
        self.lbl_ticket_impreso.setText(html)
        self.lbl_resumen_comprobante.setText(f"Folio V-{folio:06d}\nFecha: {t['fecha']}\nTotal: ${t['total']:,.2f}\n{numero_a_letras_mx(t['total'])}")
        self.cambiar_seccion(15)

    def guardar_pdf_ticket(self):
        if not self.ticket_comprobante_actual:return
        archivo,_=QFileDialog.getSaveFileName(self,"Guardar comprobante PDF",f"Ticket_V-{self.ticket_comprobante_actual['folio']:06d}.pdf","PDF (*.pdf)")
        if not archivo:return
        printer=QPrinter(QPrinter.HighResolution); printer.setOutputFormat(QPrinter.PdfFormat); printer.setOutputFileName(archivo)
        doc=QTextDocument(); doc.setHtml(self.lbl_ticket_impreso.text()); doc.print_(printer)
        QMessageBox.information(self,"PDF","Comprobante guardado correctamente.")

    def imprimir_ticket(self):
        if not self.ticket_comprobante_actual:return
        printer=QPrinter(QPrinter.HighResolution); doc=QTextDocument(); doc.setHtml(self.lbl_ticket_impreso.text()); doc.print_(printer)

    # =========================================================================
    # QUINTO REPORTE: VENTAS VIGENTES
    # =========================================================================
    def _build_tab_ventas(self):
        l=QVBoxLayout(self.tab_rep_ven); l.setSpacing(10)
        card=QFrame(); card.setProperty("class","card"); g=QGridLayout(card)
        self.lbl_fecha_ven=QLabel("Periodo consultado: —")
        self.lbl_realizacion_ven=QLabel("Fecha de realización: —")
        g.addWidget(QLabel("Reporte de Ventas · Los Expertos"),0,0,1,4)
        g.addWidget(self.lbl_fecha_ven,1,0,1,2); g.addWidget(self.lbl_realizacion_ven,1,2,1,2)
        g.addWidget(self._lbl("Fecha inicial"),2,0); self.dt_ven_ini=QDateEdit(QDate.currentDate().addMonths(-1)); configurar_date_edit_claro(self.dt_ven_ini); g.addWidget(self.dt_ven_ini,2,1)
        g.addWidget(self._lbl("Fecha final"),2,2); self.dt_ven_fin=QDateEdit(QDate.currentDate()); configurar_date_edit_claro(self.dt_ven_fin); g.addWidget(self.dt_ven_fin,2,3)
        g.addWidget(self._lbl("Método de pago"),3,0); self.cmb_ven_metodo=QComboBox(); self.cmb_ven_metodo.addItems(["Todos","Efectivo","Tarjeta"]); g.addWidget(self.cmb_ven_metodo,3,1)
        btn=QPushButton("Generar reporte"); btn.setObjectName("btn_report_generate"); btn.setStyleSheet("QPushButton { background-color: #4F46E5; color: #FFFFFF; border: none; border-radius: 6px; padding: 8px 16px; font-weight: bold; font-size: 13px; } QPushButton:hover { background-color: #3730A3; } QPushButton:pressed { background-color: #4338CA; }"); btn.clicked.connect(self.generar_reporte_ventas); g.addWidget(btn,3,2)
        pdf=QPushButton("Exportar PDF"); pdf.setObjectName("btn_secondary"); pdf.clicked.connect(self._exportar_pdf_ventas); g.addWidget(pdf,3,3)
        excel=QPushButton("Exportar Excel"); excel.setObjectName("btn_secondary"); excel.clicked.connect(self._exportar_excel_ventas); g.addWidget(excel,4,3)
        l.addWidget(card)
        self.tabla_rep_ven=QTableWidget(0,5); self.tabla_rep_ven.setHorizontalHeaderLabels(["Folio del ticket","Fecha del ticket","Cliente","Método de pago","Total a pagar"])
        self.tabla_rep_ven.horizontalHeader().setSectionResizeMode(2,QHeaderView.Stretch); self.tabla_rep_ven.setEditTriggers(QTableWidget.NoEditTriggers); l.addWidget(self.tabla_rep_ven,1)
        self.lbl_totales_ven=QLabel(); self.lbl_totales_ven.setStyleSheet("background:#24366B;color:white;padding:10px 14px;border-radius:6px;font-weight:bold;"); l.addWidget(self.lbl_totales_ven)

    def generar_reporte_ventas(self):
        ini=self.dt_ven_ini.date(); fin=self.dt_ven_fin.date()
        if ini>fin: QMessageBox.warning(self,"Fechas","La fecha inicial no puede ser mayor que la fecha final."); return
        rows=self.repo.reporte_ventas(ini.toString("yyyy-MM-dd"),fin.toString("yyyy-MM-dd"),self.cmb_ven_metodo.currentText())
        self.lbl_fecha_ven.setText(f"Periodo consultado: {ini.toString('dd/MM/yyyy')} al {fin.toString('dd/MM/yyyy')}")
        self.lbl_realizacion_ven.setText(f"Fecha de realización: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
        self.tabla_rep_ven.setRowCount(len(rows)+1); total=0.0
        for i,r in enumerate(rows):
            total+=r["total"]
            vals=[f"V-{r['folio']:06d}",r["fecha"],r["cliente"],r["metodoPago"],f"${r['total']:,.2f}"]
            for j,v in enumerate(vals):self.tabla_rep_ven.setItem(i,j,crear_celda(v))
        ult=len(rows)
        for j,v in enumerate(["TOTAL","","","",f"${total:,.2f}"]):self.tabla_rep_ven.setItem(ult,j,crear_celda(v,Qt.AlignRight|Qt.AlignVCenter,True))
        self.lbl_totales_ven.setText(f"Tickets vigentes: {len(rows)}   |   Total vendido: ${total:,.2f}   |   {numero_a_letras_mx(total)}")

    def _exportar_pdf_ventas(self):
        self.exportar_tabla_pdf("Reporte de Ventas",self.tabla_rep_ven,self.lbl_totales_ven.text(),self.lbl_fecha_ven.text())

    def _exportar_excel_ventas(self):
        self.exportar_tabla_excel("Reporte de Ventas",self.tabla_rep_ven,self.lbl_totales_ven.text(),self.lbl_fecha_ven.text())

    def _lbl(self, txt: str) -> QLabel:
        l = QLabel(txt)
        l.setProperty("class", "field_label")
        return l


def aplicar_paleta_clara_global(app: QApplication):
    """Sobrescribe la paleta oscura de Windows 11 para que ningún subcontrol herede fondos oscuros."""
    pal = QPalette()
    pal.setColor(QPalette.Window, QColor("#EEF2FF"))
    pal.setColor(QPalette.WindowText, QColor("#0F172A"))
    pal.setColor(QPalette.Base, QColor("#FFFFFF"))
    pal.setColor(QPalette.AlternateBase, QColor("#F7F8FF"))
    pal.setColor(QPalette.ToolTipBase, QColor("#FFFFFF"))
    pal.setColor(QPalette.ToolTipText, QColor("#0F172A"))
    pal.setColor(QPalette.Text, QColor("#0F172A"))
    pal.setColor(QPalette.Button, QColor("#FFFFFF"))
    pal.setColor(QPalette.ButtonText, QColor("#0F172A"))
    pal.setColor(QPalette.BrightText, QColor("#E11D48"))
    pal.setColor(QPalette.Highlight, QColor("#4F46E5"))
    pal.setColor(QPalette.HighlightedText, QColor("#FFFFFF"))
    app.setPalette(pal)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    aplicar_paleta_clara_global(app)
    win = SPVCIMainWindow()
    win.show()
    sys.exit(app.exec())
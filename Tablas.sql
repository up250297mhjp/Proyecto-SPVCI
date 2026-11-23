CREATE DATABASE IF NOT EXISTS spvci
CHARACTER SET utf8mb4
COLLATE utf8mb4_spanish_ci;
USE spvci;

CREATE TABLE Proveedor (
    idProveedor INT UNSIGNED NOT NULL AUTO_INCREMENT COMMENT 'Clave primaria',
    nombre VARCHAR(100) NOT NULL COMMENT 'Nombre del contacto',
    domicilio VARCHAR(150) NOT NULL COMMENT 'Domicilio del proveedor',
    telefono VARCHAR(15) NOT NULL COMMENT 'Telefono de contacto',
    correoE VARCHAR(100) NULL COMMENT 'Correo electronico',
    paginaWeb VARCHAR(150) NULL COMMENT 'Pagina web',
    empresa VARCHAR(100) NOT NULL COMMENT 'Razon social o nombre comercial',
    rfc VARCHAR(13) NOT NULL COMMENT 'RFC del proveedor',
    CONSTRAINT pk_proveedor PRIMARY KEY (idProveedor),
    CONSTRAINT uq_proveedor_rfc UNIQUE (rfc)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
COMMENT='Proveedores que surten mercancia';

CREATE TABLE Cliente (
    idCliente INT UNSIGNED NOT NULL AUTO_INCREMENT COMMENT 'Clave primaria',
    nombre VARCHAR(100) NOT NULL COMMENT 'Nombre completo del cliente',
    domicilio VARCHAR(150) NULL COMMENT 'Domicilio del cliente',
    telefono VARCHAR(15) NULL COMMENT 'Telefono de contacto',
    correoE VARCHAR(100) NULL COMMENT 'Correo electronico',
    paginaWeb VARCHAR(150) NULL COMMENT 'Pagina web',
    rfc VARCHAR(13) NULL COMMENT 'RFC (opcional)',
    CONSTRAINT pk_cliente PRIMARY KEY (idCliente),
    CONSTRAINT uq_cliente_rfc UNIQUE (rfc)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
COMMENT='Clientes registrados del punto de venta';

CREATE TABLE Producto (
    codigo VARCHAR(20) NOT NULL COMMENT 'Clave primaria: codigo del producto',
    nombre VARCHAR(100) NOT NULL COMMENT 'Nombre del producto',
    categoria VARCHAR(50) NOT NULL COMMENT 'Accesorios, refacciones o micas',
    unidadMedida VARCHAR(20) NOT NULL DEFAULT 'pieza' COMMENT 'Unidad de medida',
    existencia INT NOT NULL DEFAULT 0 COMMENT 'Piezas disponibles en inventario',
    stockMinimo INT NOT NULL DEFAULT 0 COMMENT 'Existencia minima antes de reabastecer',
    costo DECIMAL(10,2) NOT NULL COMMENT 'Costo de compra',
    precio DECIMAL(10,2) NOT NULL COMMENT 'Precio de venta',
    estado BOOLEAN NOT NULL DEFAULT TRUE COMMENT 'TRUE activo, FALSE baja logica',
    CONSTRAINT pk_producto PRIMARY KEY (codigo),
    CONSTRAINT ck_producto_existencia CHECK (existencia >= 0),
    CONSTRAINT ck_producto_stockminimo CHECK (stockMinimo >= 0),
    CONSTRAINT ck_producto_costo CHECK (costo >= 0),
    CONSTRAINT ck_producto_precio CHECK (precio >= 0),
    INDEX idx_producto_nombre (nombre),
    INDEX idx_producto_categoria (categoria)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
COMMENT='Articulos del inventario';

CREATE TABLE NotaRecibo (
    folio INT UNSIGNED NOT NULL AUTO_INCREMENT COMMENT 'Clave primaria: folio de la nota',
    fecha DATE NOT NULL COMMENT 'Fecha de recepcion de la mercancia',
    total DECIMAL(12,2) NOT NULL DEFAULT 0.00 COMMENT 'Importe total de la nota',
    idProveedor INT UNSIGNED NOT NULL COMMENT 'Clave foranea: proveedor que surte',
    CONSTRAINT pk_notarecibo PRIMARY KEY (folio),
    CONSTRAINT ck_notarecibo_total CHECK (total >= 0),
    CONSTRAINT fk_notarecibo_proveedor FOREIGN KEY (idProveedor)
        REFERENCES Proveedor (idProveedor)
        ON UPDATE CASCADE ON DELETE RESTRICT,
    INDEX idx_notarecibo_proveedor (idProveedor)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
COMMENT='Notas de recibo de mercancia surtida por un proveedor';

CREATE TABLE DetalleNotaRecibo (
    folioNota INT UNSIGNED NOT NULL COMMENT 'Clave foranea: nota de recibo',
    codigoProducto VARCHAR(20) NOT NULL COMMENT 'Clave foranea: producto recibido',
    cantidad INT UNSIGNED NOT NULL COMMENT 'Piezas recibidas',
    costoUnitario DECIMAL(10,2) NOT NULL COMMENT 'Costo por pieza en esta nota',
    CONSTRAINT pk_detallenotarecibo PRIMARY KEY (folioNota, codigoProducto),
    CONSTRAINT ck_detallenota_cantidad CHECK (cantidad > 0),
    CONSTRAINT ck_detallenota_costo CHECK (costoUnitario > 0),
    CONSTRAINT fk_detallenota_nota FOREIGN KEY (folioNota)
        REFERENCES NotaRecibo (folio)
        ON UPDATE CASCADE ON DELETE CASCADE,
    CONSTRAINT fk_detallenota_producto FOREIGN KEY (codigoProducto)
        REFERENCES Producto (codigo)
        ON UPDATE CASCADE ON DELETE RESTRICT,
    INDEX idx_detallenota_producto (codigoProducto)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
COMMENT='Renglones de cada nota de recibo (NotaRecibo - Producto)';

CREATE TABLE Ticket (
    folio INT UNSIGNED NOT NULL AUTO_INCREMENT COMMENT 'Clave primaria: folio del ticket',
    fecha DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP COMMENT 'Fecha y hora de la venta',
    metodoPago VARCHAR(20) NOT NULL DEFAULT 'efectivo' COMMENT 'Metodo o forma de pago',
    descuento DECIMAL(10,2) NOT NULL DEFAULT 0.00 COMMENT 'Descuento aplicado al ticket',
    total DECIMAL(12,2) NOT NULL COMMENT 'Total de la venta',
    cambio DECIMAL(10,2) NOT NULL DEFAULT 0.00 COMMENT 'Cambio entregado al cliente',
    idCliente INT UNSIGNED NULL COMMENT 'Clave foranea: cliente (opcional, 0..1)',
    CONSTRAINT pk_ticket PRIMARY KEY (folio),
    CONSTRAINT ck_ticket_descuento CHECK (descuento >= 0),
    CONSTRAINT ck_ticket_total CHECK (total >= 0),
    CONSTRAINT ck_ticket_cambio CHECK (cambio >= 0),
    CONSTRAINT fk_ticket_cliente FOREIGN KEY (idCliente)
        REFERENCES Cliente (idCliente)
        ON UPDATE CASCADE ON DELETE SET NULL,
    INDEX idx_ticket_fecha (fecha),
    INDEX idx_ticket_cliente (idCliente)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
COMMENT='Comprobantes de venta';

CREATE TABLE DetalleTicket (
    folioTicket INT UNSIGNED NOT NULL COMMENT 'Clave foranea: ticket',
    codigoProducto VARCHAR(20) NOT NULL COMMENT 'Clave foranea: producto vendido',
    cantidad INT UNSIGNED NOT NULL COMMENT 'Piezas vendidas',
    precioUnitario DECIMAL(10,2) NOT NULL COMMENT 'Precio por pieza al momento de la venta',
    descuento DECIMAL(10,2) NOT NULL DEFAULT 0.00 COMMENT 'Descuento del renglon',
    CONSTRAINT pk_detalleticket PRIMARY KEY (folioTicket, codigoProducto),
    CONSTRAINT ck_detalleticket_cantidad CHECK (cantidad > 0),
    CONSTRAINT ck_detalleticket_precio CHECK (precioUnitario >= 0),
    CONSTRAINT ck_detalleticket_descuento CHECK (descuento >= 0),
    CONSTRAINT fk_detalleticket_ticket FOREIGN KEY (folioTicket)
        REFERENCES Ticket (folio)
        ON UPDATE CASCADE ON DELETE CASCADE,
    CONSTRAINT fk_detalleticket_producto FOREIGN KEY (codigoProducto)
        REFERENCES Producto (codigo)
        ON UPDATE CASCADE ON DELETE RESTRICT,
    INDEX idx_detalleticket_producto (codigoProducto)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
COMMENT='Renglones de cada ticket (Ticket - Producto)';


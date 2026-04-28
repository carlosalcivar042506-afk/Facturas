# Control de Facturas — Gemini AI

Sistema web para extracción automática de datos de facturas en PDF usando Gemini 2.5 Flash.

## ✅ Funcionalidades

- **Carga de PDFs** (texto o imagen) con arrastre o selección
- **Extracción automática** con Gemini 2.5 Flash (con fallback a 2.0-flash y 2.0-flash-lite)
- **Casilla de API Key editable** en la barra superior — se guarda en el navegador y se puede cambiar en cualquier momento
- **Búsqueda en todas las páginas** (factura + orden de compra + guía de remisión)
- **Tabla de cuentas → RUBRO** con 53 códigos contables
- **Edición manual** de cualquier campo antes de exportar
- **Exportación a Excel** con formato de Control de Facturas

## 📋 Campos Extraídos

| Campo | Descripción |
|-------|-------------|
| FECHA | Fecha de emisión de la factura (DD/MM/AAAA) |
| PROVEEDOR | Nombre comercial del emisor (nunca Mobilsol) |
| FACTURA | Número completo (001-001-000002702) |
| SOLPED | Código SOL + dígitos (buscado en todas las páginas) |
| ORDEN | Código OC + dígitos (buscado en todas las páginas) |
| RUBRO | Nombre del rubro según tabla de cuentas |
| CUENTA | Código contable (52-1-08-01-12) |
| PROYECTO | Solo si aparece literalmente "Proyecto:" en el PDF |
| DESCRIPCIÓN | Observaciones de la OC o descripción de ítems |
| SUB TOTAL | Monto sin IVA |
| IVA | Monto del IVA |
| TOTAL FACTURAR | Total de la factura |

## 🔑 API Key de Gemini

1. Obtener key gratis en: https://aistudio.google.com/app/apikey
2. Pegar en la casilla azul **superior** de la aplicación
3. Hacer clic en **💾 Guardar**
4. La key se guarda en el navegador y persiste entre sesiones

### Límites del Tier Gratuito
- 20 solicitudes/día por modelo
- El sistema intenta automáticamente: `gemini-2.5-flash` → `gemini-2.0-flash` → `gemini-2.0-flash-lite`
- Si se agota la cuota, el sistema mostrará un mensaje claro indicando esperar al día siguiente

## 🏃 Uso

1. Abrir la aplicación en el navegador
2. Configurar la API Key en la barra superior
3. Arrastrar PDF(s) a la zona de carga
4. Esperar la extracción (10-30 segundos por factura)
5. Revisar y editar campos si es necesario
6. Exportar a Excel con el botón "Descargar Excel"

## 📊 Tabla de Cuentas (RUBRO)

| Código | RUBRO |
|--------|-------|
| 52-1-08-01-01 | BAÑOS |
| 52-1-08-01-02 | RIEGO |
| 52-1-08-01-05 | EQUIPOS CONTRA INCENDIO |
| 52-1-08-01-06 | INSTALACIONES ELECTRICAS |
| 52-1-08-01-12 | AIRE ACONDICIONADO |
| 52-1-08-01-18 | GENERACION |
| 52-1-08-01-27 | IMPREVISTOS |
| 52-1-24-02-01 | FOCOS Y LAMPARAS |
| 52-1-28-02-01 | TECNICO |
| (ver app para lista completa de 53 códigos) | |

## 🛠️ Arquitectura Técnica

- **Backend**: Python/Flask + Google Gemini API
- **Procesamiento PDF**: PyMuPDF (fitz) — convierte páginas a imágenes JPEG
- **IA**: Gemini 2.5 Flash (visión multimodal) — analiza todas las páginas
- **Frontend**: HTML + Tailwind CSS (vanilla JS, sin frameworks)
- **Proceso**: PM2 (daemon)
- **Puerto**: 3000

## 📡 Endpoints API

| Método | Ruta | Descripción |
|--------|------|-------------|
| GET | `/` | Interfaz web |
| GET | `/api/health` | Estado del sistema |
| POST | `/api/set-key` | Guardar API Key |
| POST | `/api/config` | Actualizar configuración |
| GET | `/api/cuentas` | Lista de cuentas/rubros |
| POST | `/api/extract` | Extraer datos de PDF |
| POST | `/api/export` | Exportar a Excel |

## 🔧 Replicar/Ajustar el Sistema

### Agregar nuevos rubros
Editar `server.py` → diccionario `TABLA_CUENTAS`:
```python
TABLA_CUENTAS = {
    "52-1-08-01-01": "BAÑOS",
    # Agregar aquí: "codigo": "NOMBRE_RUBRO",
}
```

### Modificar el prompt de extracción
Editar la función `build_prompt()` en `server.py` para ajustar qué campos extraer o cómo interpretarlos.

### Cambiar el modelo de IA
Editar en `server.py`:
```python
CONFIG = {
    "models": ["gemini-2.5-flash", "gemini-2.0-flash", "gemini-2.0-flash-lite"],
}
```

### Reiniciar el servidor
```bash
cd /home/user/factura-ocr
pm2 restart factura-ocr
```

## 📅 Última actualización
Abril 2026

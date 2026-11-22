# Puerto A&M | Business Intelligence

Dashboard corporativo conectado a un archivo Excel real. Los indicadores, gráficos, filtros y reportes se calculan en el servidor a partir de `data/Puerto_AM_Base_Datos_DEMO.xlsx` con `openpyxl`. **No hay datos ficticios incrustados en JavaScript.** Incluye además un asistente de IA («A&M Copiloto») que responde **únicamente** con lo calculado desde ese Excel.

> ⚠️ Simulación académica para Top Learning. Todos los clientes, pedidos, productos, precios y metas son inventados. Ninguna cifra representa operaciones reales de Puerto A&M. Importes en MXN.

## Publicación

- **Repositorio:** https://github.com/psiquipitirija/puerto-am-bi
- **Demo pública temporal:** https://patients-conclusions-potentially-homeland.trycloudflare.com
- El enlace HTTPS funciona mientras sigan activos el servidor local y el túnel de Cloudflare. Para una URL permanente, sigue `DEPLOY.md` y crea el Web Service en Render (el repo ya trae `render.yaml`).
- **GitHub Pages no puede hostear este dashboard**: solo sirve archivos estáticos y aquí el backend Python/FastAPI calcula los datos del Excel y alimenta el chat.

## Estado del workspace

Se inspeccionó la carpeta `../stilis-landing`: es un landing de **Stilís** (React + Vite) sin relación con Puerto A&M ni con un chatbot. No se modificó. Por ello este dashboard se construyó como **aplicación independiente**, lista para integrarse después como sección administrativa de un portal Puerto A&M.

## Estructura

```
puerto-am-bi/
├── app/
│   ├── __init__.py
│   ├── main.py                 # FastAPI: endpoints, chat, exportación, archivos estáticos
│   ├── chat.py                 # Asistente fundamentado: intents, regla SIN_DATO y opción LLM
│   └── excel_reader.py         # Lectura del Excel, filtros y agregaciones
├── data/
│   └── Puerto_AM_Base_Datos_DEMO.xlsx
├── scripts/
│   ├── generar_excel.py        # Genera la base de 360 registros
│   ├── verificar_conexion.py   # Prueba real de lectura → edición → actualización
│   └── probar_chat.py          # Prueba 20 preguntas del asistente (--local sin servidor)
├── static/
│   ├── index.html
│   ├── styles.css
│   ├── app.js
│   └── vendor/chart.umd.min.js # Chart.js 4.5 local (sin CDN en runtime)
├── requirements.txt
├── .env.example               # Variables opcionales del modelo externo
├── ejecutar.bat
└── README.md
```

## Instalación y ejecución

```powershell
cd "C:\Users\Administrador\OneDrive\Documentos\Curso IA\OpenCode-Carpeta\puerto-am-bi"
python -m pip install -r requirements.txt
python scripts/generar_excel.py        # solo si quieres recrear la base original
python -m uvicorn app.main:app --host 127.0.0.1 --port 8124
```

Abre **http://127.0.0.1:8124/**. En Windows también puedes hacer doble clic en `ejecutar.bat`.

No necesitas Node, npm ni conexión a internet para usar la app (Chart.js está incluido en `static/vendor/`).

## El Excel

`Puerto_AM_Base_Datos_DEMO.xlsx` contiene cuatro hojas con formato corporativo, tablas, filtros, validaciones, barras de datos, congelación de paneles y moneda MXN:

| Hoja | Contenido |
| --- | --- |
| **Ventas** | 360 pedidos ficticios (120 por mes) de enero a marzo de 2026. Columnas: Fecha, Pedido, Producto, Categoría, Cantidad, Precio unitario, Costo unitario, Ingreso, Costo total, Utilidad, Canal, Método de pago. Las tres columnas calculadas son fórmulas de Excel. |
| **Productos** | 18 SKU con categoría, precio, costo y margen. |
| **Metas** | Metas mensuales de ingresos, pedidos y utilidad para los tres meses. |
| **Instrucciones** | Aviso de simulación, reglas de edición, pasos para refrescar y fórmulas usadas. |

### Editar y actualizar

1. Abre el Excel y cambia, por ejemplo, la **Cantidad** o el **Precio unitario** de cualquier pedido en `Ventas`.
2. Guarda con Ctrl+S, manteniendo el mismo nombre y la misma carpeta `data/`.
3. En el dashboard pulsa **Actualizar datos** (barra superior).

El backend vuelve a leer el archivo (`POST /api/refresh`) y recalcula todo. También detecta cambios de `mtime`/tamaño automáticamente en la siguiente petición. El indicador «Excel conectado · fecha/hora» muestra cuándo se leyó por última vez.

**Reglas:** no renombrar hojas ni encabezados de la fila 1; fechas reales de Excel; números sin símbolos de moneda en cantidad/precio/costo; no dejar filas vacías intermedias. Se pueden agregar pedidos nuevos.

## Cómo se calcula

Todas las operaciones están en `app/excel_reader.py`:

```
ingreso por fila = Cantidad × Precio unitario
costo por fila   = Cantidad × Costo unitario
utilidad bruta   = Σ ingresos − Σ costos
ticket promedio  = Σ ingresos ÷ pedidos únicos
margen bruto     = utilidad bruta ÷ Σ ingresos
crecimiento      = (ingresos del último mes / ingresos del mes anterior) − 1
cumplimiento     = Σ ingresos mensuales ÷ Σ metas de ingresos de esos meses
```

Los filtros de fecha, producto, categoría y canal se aplican **antes** de cualquier agregación. Las metas son globales por mes (no se prorratean por producto o categoría); si filtras por producto, el cumplimiento se compara con la meta mensual completa y así se indica en la interfaz. Se usa `Decimal` para importes y se redondea solo al entregar JSON.

## API

| Método | Endpoint | Descripción |
| --- | --- | --- |
| GET | `/api/health` | Estado del servicio y existencia del Excel |
| GET | `/api/meta` | Archivo, modificación, filas, caché |
| POST | `/api/refresh` | Fuerza nueva lectura del Excel |
| GET | `/api/dashboard` | KPIs, gráficos, opciones de filtros, tabla y metadatos |
| POST | `/api/chat` | Respuesta fundamentada en el Excel (acepta los mismos filtros) |
| GET | `/api/export?formato=csv\|xlsx` | Reporte de las filas filtradas |

Los parámetros de filtro son `fecha_desde`, `fecha_hasta` (aaaa-mm-dd), `producto`, `categoria` y `canal`. Ejemplo:

```
http://127.0.0.1:8124/api/dashboard?fecha_desde=2026-02-01&fecha_hasta=2026-02-28&categoria=Oficina
```

Si el Excel falta, está corrupto o cambiaron los encabezados, la API responde `502` con un mensaje legible y el dashboard muestra un banner de error con botón de reintento.

## Dashboard

- Sidebar corporativa, topbar con estado de conexión, temas claro/oscuro y diseño responsive.
- 8 KPIs + banda de cumplimiento de metas.
- Seis gráficos Chart.js: evolución (línea + barras de pedidos), ingresos vs. costos, categorías (dona), productos destacados, metas vs. reales y canales.
- Filtros por rango de fechas, producto, categoría y canal, con rangos rápidos por mes.
- Tabla de 360 registros con búsqueda, orden por columna y paginación.
- Exportación CSV (BOM UTF-8 para Excel) y XLSX con formato.
- Sección «Fuente de datos» con metadatos reales del archivo e instrucciones plegables.
- Sección «Asistente IA» con interfaz de chat, sugerencias, citas de fuente y acceso rápido flotante.

## Asistente IA · «A&M Copiloto»

Es un analista de datos con personalidad ejecutiva, cálida y breve. Responde en español de México, cita la hoja de origen de cada cifra y **nunca afirma algo que no esté en el Excel**.

### Regla estricta

1. El backend construye el payload con `build_dashboard(filtros)` y se lo pasa al motor de chat.
2. El motor **solo** puede leer ese payload: KPIs, tendencias, categorías, productos, metas, canales, métodos de pago y filas filtradas.
3. Si el dato no existe (clientes, empleados, sueldos, impuestos, proyecciones, meses sin registros, etc.), responde exactamente:

   > No tengo ese dato en el Excel cargado. Solo respondo con información de las hojas Ventas, Productos y Metas. Puedo ayudarte con ventas, pedidos, ticket promedio, unidades, costos, utilidad, margen, crecimiento, metas, categorías, productos, canales y métodos de pago.

4. El frontend envía `{mensaje, ...state.filters}` a `POST /api/chat`, así que el chat responde sobre **el mismo subconjunto** que ves en los KPIs y gráficos. Al cambiar los filtros, la cabecera del chat muestra el periodo/producto/categoría/canal activo.
5. Al pulsar **Actualizar datos**, la conversación se reinicia para que ninguna respuesta antigua quede con cifras desactualizadas.

Ejemplos que sí responde con datos: «¿Cuánto vendimos en total?», «¿Cuánto vendimos en marzo?», «¿Cómo fue febrero comparado con enero?», «¿Cuál es el margen bruto?», «¿Cuál es el cumplimiento de metas?», «¿Cuánto vendió la categoría Oficina?», «¿Cuáles son los métodos de pago?».

### Modo de ejecución

- **Por defecto (sin configuración):** motor local determinista en `app/chat.py`. No llama a ningún servicio externo, no necesita API key y siempre cita cifras verificables.
- **Opcional:** si defines `PUERTOAM_LLM_API_KEY`, se usa un modelo externo con `temperature=0`, un system prompt con la regla estricta y un contexto cerrado con solo los datos del Excel (resúmenes + máximo 40 filas de detalle). Si falla la llamada, cae automáticamente al motor local. Copia `.env.example` como referencia y exporta las variables antes de iniciar el servidor.

```powershell
# Ejemplo (solo si quieres el modelo externo)
$env:PUERTOAM_LLM_API_KEY = "tu-clave"
$env:PUERTOAM_LLM_MODEL = "gpt-4o-mini"
python -m uvicorn app.main:app --host 127.0.0.1 --port 8124
```

Prueba rápida del asistente:

```powershell
python scripts/probar_chat.py            # contra el servidor en :8124
python scripts/probar_chat.py --local    # sin servidor, llama al motor directo
```

## Verificación realizada

- ✅ El Excel tiene **360 registros** (mínimo solicitado: 300), 18 productos, 3 metas y 4 hojas.
- ✅ Sumas directas con `openpyxl` coinciden exactamente con la API: ingresos **$566,546**, costos **$317,971**, utilidad **$248,575**, 360 pedidos.
- ✅ **Conexión real probada**: se cambió `Ventas!E2` de 4 a 14 unidades, se guardó el Excel y tras `Actualizar datos` los ingresos pasaron de $566,546 a $567,226 (+10 × $68.00). Luego se regeneró la base original. Reproducible con `python scripts/verificar_conexion.py`.
- ✅ Filtros verificados: enero = 120 pedidos / $177,387; categoría+ producto reducen KPIs y tabla de forma coherente.
- ✅ Exportación CSV/XLSX con `Content-Disposition` correcto.
- ✅ Auditoría CDP en 1600, 1440, 1280, 1024, 900, 768, 600 y 390 px: sin overflow horizontal, sin recortes y 6 gráficos activos. Sin errores de consola.
- ✅ Menú móvil, tema oscuro, refresco y exportación probados en navegador a 390 px.
- ✅ Asistente probado con 20 preguntas: métricas globales, meses, comparaciones, metas, categorías, productos, canales, métodos de pago, filtros activos y tres preguntas sin dato que devuelven `SIN_DATO` sin inventar. Reproducible con `python scripts/probar_chat.py`.
- ✅ Chat verificado en navegador de escritorio y a 390 px: sugerencias, envío con Enter, citas de fuente, botón flotante, reinicio y respuesta con el filtro de enero ($177,387). Sin errores de consola.
- ✅ Auditoría CDP repetida con la sección del chat en 1600…390 px: sin overflow horizontal ni recortes.

## Límites

Los importes no incluyen impuestos ni inflación. Las metas son supuestos ficticios. La demo no implementa autenticación: es un panel administrativo local para práctica; si se integra en un portal público, debe colocarse detrás de login y separarse de la vista pública.

El asistente no consulta internet ni bases de datos externas. El modo LLM opcional sí envía el contexto calculado al proveedor configurado; no lo actives si los datos fueran confidenciales. Las sugerencias del chat son preguntas de ejemplo, no datos.

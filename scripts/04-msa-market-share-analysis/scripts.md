# 4. SCRIPTS MSA (MARKET SHARE ANALYSIS)

**Nota:** *Los siguientes scripts están diseñados para el análisis de datos orientado a la estimación de market share, la generación de visualizaciones analíticas y la identificación de insights estratégicos a partir de datos históricos obtenidos de TradeMap. Se sugiere implementar este flujo de trabajo en entornos de ejecución Python tales como:*

- [Google Colab](https://colab.research.google.com/)
- [Python-Fiddle](https://python-fiddle.com/)
- [Kaggle Notebooks](https://www.kaggle.com/code)
- [Deepnote](https://deepnote.com/)

## 4.1. Importar e instalar librerías

```
# Instalar librerías si es necesario
!pip install openpyxl

import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
from google.colab import files
import warnings
warnings.filterwarnings('ignore')

# Configurar estilo de gráficos
plt.style.use('seaborn-v0_8')
sns.set_palette("husl")
```

4.2. Cargar archivos fuente (Trademap)
Nota: El script reconoce automáticamente las columnas de países y de valores por año, con cabeceras en español o en inglés. Si la detección no es concluyente, pedirá elegir las columnas por su número. Si el botón de subida no funciona, suba el archivo desde el panel de carpetas de Colab y escriba su nombre en NOMBRE_ARCHIVO.
```
# Detección flexible de columnas (español / inglés) + asignación manual

# Modo de asignación de columnas:
#   "Automático" -> el script detecta las columnas y solo pregunta si no está seguro.
#   "Manual"     -> el script siempre le pide elegir las columnas.
MODO_ASIGNACION = "Automático"  # @param ["Automático", "Manual"]

# Si el botón "Elegir archivos" no funciona, suba el Excel desde el panel de carpetas
# (ícono 📁 a la izquierda de Colab) y escriba aquí su nombre. Si lo deja vacío y hay
# un solo Excel en la carpeta, se usará ese.
NOMBRE_ARCHIVO = ""  # @param {type:"string"}

import re
import io
import unicodedata

# --- Vocabulario reconocido en las cabeceras (sin tildes, en minúsculas) ---
PALABRAS_PAIS = [
    'exportadores', 'importadores', 'exporters', 'importers', 'exportador', 'importador',
    'exporter', 'importer', 'pais', 'paises', 'country', 'countries', 'socio', 'socios',
    'partner', 'partners', 'mercado', 'mercados', 'market', 'markets', 'proveedor',
    'proveedores', 'supplier', 'suppliers', 'origen', 'origin', 'destino', 'destination',
]
PALABRAS_VALOR = [
    'valor', 'value', 'importado', 'importada', 'importados', 'exportado', 'exportada',
    'exportados', 'imported', 'exported', 'importaciones', 'exportaciones', 'imports',
    'exports', 'usd', 'dolares', 'dollars', 'miles', 'thousand',
]
# Columnas con año que NO son valores comerciales (cantidades, tasas, participación, etc.)
PALABRAS_EXCLUIR = [
    'cantidad', 'quantity', 'unidad', 'unit', 'tasa', 'rate', 'crecimiento', 'growth',
    'participacion', 'share', 'saldo', 'balance', 'unitario', 'unitaria', 'precio',
    'price', 'distancia', 'distance', 'concentracion', 'concentration', 'arancel', 'tariff',
    'posicion', 'ranking', '%',
]
# Filas agregadas que deben excluirse (equivalentes de "Mundo")
FILAS_TOTAL = {'mundo', 'world', 'total', 'monde', 'total mundial', 'world total', 'todos', 'all'}

PATRON_ANIO = re.compile(r'(?<!\d)((?:19|20)\d{2})(?!\d)')


def _normalizar(texto):
    """Minúsculas, sin tildes y con espacios simples (solo para comparar)."""
    texto = '' if texto is None else str(texto)
    texto = unicodedata.normalize('NFKD', texto)
    texto = ''.join(c for c in texto if not unicodedata.combining(c))
    return re.sub(r'\s+', ' ', texto).strip().lower()


def _contiene(texto_normalizado, palabras):
    tokens = set(re.findall(r'[a-z%]+', texto_normalizado))
    return any(p in tokens if p.isalpha() else p in texto_normalizado for p in palabras)


def _anio_de(cabecera):
    encontrados = PATRON_ANIO.findall(str(cabecera))
    return int(encontrados[-1]) if len(set(encontrados)) == 1 else None


def _a_numero(serie):
    """Convierte a número; tolera separadores de miles en archivos exportados como texto."""
    if serie.dtype.kind in 'if':
        return serie.astype(float)
    texto = serie.astype(str).str.replace('\u00a0', '', regex=False).str.replace(' ', '', regex=False).str.strip()
    texto = texto.where(~texto.str.fullmatch(r'-?\d{1,3}(,\d{3})+'), texto.str.replace(',', '', regex=False))
    texto = texto.where(~texto.str.fullmatch(r'-?\d{1,3}(\.\d{3})+'), texto.str.replace('.', '', regex=False))
    return pd.to_numeric(texto, errors='coerce')


def _proporcion_numerica(serie):
    no_vacios = serie.dropna()
    no_vacios = no_vacios[no_vacios.astype(str).str.strip() != '']
    if len(no_vacios) == 0:
        return 0.0
    return _a_numero(no_vacios).notna().mean()


def leer_archivo_trademap(nombre_archivo):
    """Lee .xlsx, .xls (incluye los .xls de Trade Map que en realidad son HTML o texto), .csv y .txt."""
    extension = nombre_archivo.lower().rsplit('.', 1)[-1]
    with open(nombre_archivo, 'rb') as f:
        contenido = f.read()
    errores = []

    if extension in ('xlsx', 'xlsm', 'xls'):
        try:
            return pd.read_excel(io.BytesIO(contenido), sheet_name=0, header=None)
        except Exception as e:
            errores.append(f'Excel: {e}')
    # Trade Map suele descargar ".xls" que por dentro son una tabla HTML
    try:
        tablas = pd.read_html(io.BytesIO(contenido))
        if tablas:
            t = max(tablas, key=lambda t: t.shape[0] * t.shape[1])
            if not isinstance(t.columns, pd.RangeIndex):      # la cabecera vuelve a ser una fila
                cab = [c[-1] if isinstance(c, tuple) else c for c in t.columns]
                t = pd.DataFrame([cab] + t.values.tolist())
            return t
    except Exception as e:
        errores.append(f'HTML: {e}')
    # ...o texto separado por tabulaciones / punto y coma / comas
    for codificacion in ('utf-8-sig', 'utf-16', 'latin-1'):
        try:
            texto = contenido.decode(codificacion)
        except Exception:
            continue
        for separador in ('\t', ';', ','):
            try:
                tabla = pd.read_csv(io.StringIO(texto), sep=separador, header=None, dtype=str,
                                    engine='python', skip_blank_lines=True)
                if tabla.shape[1] >= 2:
                    return tabla
            except Exception as e:
                errores.append(f'Texto ({codificacion}, {separador!r}): {e}')
    raise ValueError('❌ No fue posible leer el archivo. Verifique que sea la exportación de Trade Map '
                     '(.xlsx, .xls, .csv o .txt).\nDetalle: ' + ' | '.join(errores[:3]))


def detectar_fila_encabezado(tabla, max_filas=25):
    """Trade Map a veces incluye filas de título antes de la cabecera: se busca la fila más probable."""
    mejor_fila, mejor_puntaje = 0, -1
    for i in range(min(max_filas, len(tabla))):
        celdas = [_normalizar(c) for c in tabla.iloc[i].tolist() if pd.notna(c) and str(c).strip() != '']
        if len(celdas) < 2:
            continue
        con_anio = sum(1 for c in celdas if PATRON_ANIO.search(c))
        con_pais = sum(1 for c in celdas if _contiene(c, PALABRAS_PAIS))
        puntaje = con_anio * 2 + con_pais * 3 + (1 if con_anio and con_pais else 0)
        if puntaje > mejor_puntaje:
            mejor_fila, mejor_puntaje = i, puntaje
    return mejor_fila


def preparar_tabla(tabla):
    """Asigna la fila de cabecera detectada y limpia los nombres de columna."""
    fila = detectar_fila_encabezado(tabla)
    cabeceras, vistos = [], {}
    for j, c in enumerate(tabla.iloc[fila].tolist()):
        if isinstance(c, float) and c.is_integer():   # Excel guarda 2020 como 2020.0
            c = int(c)
        nombre = re.sub(r'\s+', ' ', str(c)).strip() if pd.notna(c) and str(c).strip() != '' else f'Columna_{j}'
        if nombre in vistos:                 # evita nombres repetidos
            vistos[nombre] += 1
            nombre = f'{nombre} ({vistos[nombre]})'
        else:
            vistos[nombre] = 0
        cabeceras.append(nombre)
    datos = tabla.iloc[fila + 1:].copy()
    datos.columns = cabeceras
    datos = datos.dropna(how='all').reset_index(drop=True)
    return datos, fila


def detectar_columnas(datos):
    """Devuelve (columna_paises, {columna_valor: año}, lista_de_avisos)."""
    avisos = []

    # 1) Columna de países / socios comerciales
    candidatos_pais = [c for c in datos.columns
                       if _contiene(_normalizar(c), PALABRAS_PAIS) and not PATRON_ANIO.search(str(c))
                       and _proporcion_numerica(datos[c]) < 0.5]
    if not candidatos_pais:
        # Sin palabra clave: primera columna mayoritariamente de texto
        candidatos_pais = [c for c in datos.columns
                           if datos[c].notna().mean() > 0.5 and _proporcion_numerica(datos[c]) < 0.2][:1]
        if candidatos_pais:
            avisos.append(f'La columna de países se dedujo por su contenido: "{candidatos_pais[0]}".')
    if len(candidatos_pais) > 1:
        avisos.append(f'Varias columnas parecen contener países: {candidatos_pais}.')
        columna_pais = None
    else:
        columna_pais = candidatos_pais[0] if candidatos_pais else None

    # 2) Columnas de valores comerciales por año
    con_anio = {c: _anio_de(c) for c in datos.columns if c != columna_pais and _anio_de(c) is not None}
    con_anio = {c: a for c, a in con_anio.items() if not _contiene(_normalizar(c), PALABRAS_EXCLUIR)}
    con_valor = {c: a for c, a in con_anio.items() if _contiene(_normalizar(c), PALABRAS_VALOR)}
    columnas_valor = con_valor if con_valor else con_anio   # p. ej. cabeceras que son solo "2020", "2021"...
    columnas_valor = {c: a for c, a in columnas_valor.items() if _proporcion_numerica(datos[c]) >= 0.5}

    # Si hay dos tipos de valor (p. ej. importado y exportado) el resultado no es concluyente
    familias = {PATRON_ANIO.sub('', _normalizar(c)).strip(' -_:') for c in columnas_valor}
    anios = list(columnas_valor.values())
    if len(familias) > 1:
        avisos.append('Se encontraron varios tipos de valor por año: ' + ', '.join(sorted(familias)) + '.')
    if len(anios) != len(set(anios)):
        avisos.append('Hay años repetidos entre las columnas de valores.')

    concluyente = (columna_pais is not None and len(columnas_valor) > 0
                   and len(familias) == 1 and len(anios) == len(set(anios)))
    return columna_pais, (columnas_valor if concluyente else {}), avisos


def mostrar_columnas(datos):
    print('\n📋 Columnas disponibles en el archivo:')
    for i, c in enumerate(datos.columns):
        ejemplo = datos[c].dropna().astype(str).head(2).tolist()
        print(f'  [{i:>2}] {c}   → ej.: {ejemplo}')


def _pedir_indices(mensaje, total, multiple):
    while True:
        respuesta = input(mensaje).strip()
        try:
            partes = [p for p in re.split(r'[,\s;]+', respuesta) if p]
            indices = []
            for p in partes:
                if '-' in p:                       # admite rangos: 2-6
                    a, b = map(int, p.split('-'))
                    indices.extend(range(min(a, b), max(a, b) + 1))
                else:
                    indices.append(int(p))
            if not indices or (not multiple and len(indices) != 1):
                raise ValueError
            if any(i < 0 or i >= total for i in indices):
                print(f'   ⚠️ Use números entre 0 y {total - 1}.')
                continue
            return list(dict.fromkeys(indices))
        except ValueError:
            print('   ⚠️ Entrada no válida. ' + ('Ejemplo: 2,3,4 o 2-6' if multiple else 'Escriba un solo número.'))


def asignacion_manual(datos, sugerencia_pais=None, sugerencia_valores=None, max_intentos=3):
    """Permite elegir las columnas por su número (índice) y valida la elección."""
    mostrar_columnas(datos)
    columnas = list(datos.columns)
    for intento in range(1, max_intentos + 1):
        # 1) Columna de países
        if sugerencia_pais is not None:
            print(f'\n💡 Sugerencia para países: [{columnas.index(sugerencia_pais)}] {sugerencia_pais}')
        columna_pais = columnas[_pedir_indices('👉 Número de la columna de PAÍSES / socios comerciales: ',
                                               len(columnas), False)[0]]
        if _proporcion_numerica(datos[columna_pais]) > 0.5:
            print(f'   ⚠️ "{columna_pais}" contiene principalmente números, no nombres de países.')
            continue

        # 2) Columnas de valores
        if sugerencia_valores:
            print('💡 Sugerencia para valores: ' + ','.join(str(columnas.index(c)) for c in sugerencia_valores))
        indices = _pedir_indices('👉 Números de las columnas de VALOR COMERCIAL por año '
                                 '(separados por coma o rango, ej. 2,3,4 o 2-6): ', len(columnas), True)
        elegidas = [columnas[i] for i in indices if columnas[i] != columna_pais]
        no_numericas = [c for c in elegidas if _proporcion_numerica(datos[c]) < 0.5]
        if not elegidas:
            print('   ⚠️ Elija columnas de valores distintas a la columna de países.')
            continue
        if no_numericas:
            print('   ⚠️ Estas columnas no contienen valores numéricos: ' + ', '.join(no_numericas))
            continue

        # 3) Año de cada columna (se pregunta solo si la cabecera no lo indica)
        columnas_valor = {}
        for c in elegidas:
            anio = _anio_de(c)
            while anio is None:
                r = input(f'   La columna "{c}" no indica el año. Escriba el año (ej. 2023): ').strip()
                anio = int(r) if re.fullmatch(r'(19|20)\d{2}', r) else None
                if anio is None:
                    print('   ⚠️ Escriba un año de cuatro cifras, por ejemplo 2023.')
            columnas_valor[c] = anio

        errores = validar_correspondencias(datos, columna_pais, columnas_valor)
        if not errores:
            return columna_pais, columnas_valor
        print('\n❌ La selección tiene inconsistencias:')
        for e in errores:
            print('   -', e)
        print(f'   Intente de nuevo ({intento}/{max_intentos}).\n')
    raise ValueError('❌ No fue posible identificar las columnas requeridas (países y valores comerciales '
                     'por año). Verifique que el archivo sea una exportación de Trade Map con datos.')


def validar_correspondencias(datos, columna_pais, columnas_valor):
    errores = []
    if columna_pais in columnas_valor:
        errores.append('La columna de países no puede ser también una columna de valores.')
    if _proporcion_numerica(datos[columna_pais]) > 0.5:
        errores.append(f'La columna "{columna_pais}" contiene principalmente números, no nombres de países.')
    for c, a in columnas_valor.items():
        if _proporcion_numerica(datos[c]) < 0.5:
            errores.append(f'La columna "{c}" no contiene valores numéricos suficientes.')
    anios = list(columnas_valor.values())
    if len(anios) != len(set(anios)):
        errores.append('Hay dos o más columnas asignadas al mismo año.')
    if not columnas_valor:
        errores.append('Debe seleccionar al menos una columna de valores.')
    return errores


# ---------------- EJECUCIÓN ----------------
import os, glob

def obtener_archivo(nombre=''):
    """1) usa el nombre indicado; 2) si hay un solo Excel ya subido, lo usa; 3) abre el botón de subida."""
    extensiones = ('.xlsx', '.xls', '.xlsm', '.csv', '.txt')
    if nombre:
        for ruta in (nombre, os.path.join('/content', nombre)):
            if os.path.exists(ruta):
                return ruta
        raise FileNotFoundError(f'❌ No se encontró "{nombre}". Revise que el nombre sea exacto '
                                '(con la extensión, p. ej. archivo.xlsx) y que esté en el panel de carpetas 📁.')
    existentes = sorted([f for f in glob.glob('/content/*') if f.lower().endswith(extensiones)],
                        key=os.path.getmtime, reverse=True)
    if len(existentes) == 1:
        print(f'📂 Se usará el archivo ya subido: {os.path.basename(existentes[0])}')
        return existentes[0]
    try:
        subido = files.upload()
        if subido:
            return list(subido.keys())[0]
    except Exception as error:
        print(f'⚠️ El botón de subida no funcionó ({type(error).__name__}).')
    if existentes:
        print(f'📂 Se usará el archivo más reciente de la carpeta: {os.path.basename(existentes[0])}')
        return existentes[0]
    raise FileNotFoundError('❌ No se pudo cargar el archivo. Súbalo arrastrándolo al panel de carpetas 📁 '
                            '(a la izquierda de Colab), escriba su nombre en NOMBRE_ARCHIVO y ejecute de nuevo esta celda.')

# Subir archivo Excel directamente desde TradeMap
filename = obtener_archivo(NOMBRE_ARCHIVO.strip())
uploaded = {filename: None}

# Leer datos (detecta formato y fila de cabecera)
tabla_original = leer_archivo_trademap(filename)
df_raw, fila_cabecera = preparar_tabla(tabla_original)

# Verificar estructura original
print(f"Fila de cabecera detectada: {fila_cabecera + 1}")
print("Columnas disponibles:", df_raw.columns.tolist())
print("\nPrimeras 5 filas:")
print(df_raw.head())

# Identificar columnas (automático con respaldo manual)
columna_pais, columnas_valor, avisos = detectar_columnas(df_raw)
modo_manual = _normalizar(MODO_ASIGNACION).startswith('manual')
if modo_manual or columna_pais is None or not columnas_valor:
    if not modo_manual:
        print('\n⚠️ La detección automática no fue concluyente:')
        for a in avisos or ['No se encontraron las columnas requeridas.']:
            print('   -', a)
        print('   Seleccione las columnas manualmente.')
    sugerencia_valores = [c for c in df_raw.columns if _anio_de(c) is not None
                          and not _contiene(_normalizar(c), PALABRAS_EXCLUIR)]
    columna_pais, columnas_valor = asignacion_manual(df_raw, columna_pais, sugerencia_valores)
else:
    for a in avisos:
        print('ℹ️', a)
    errores = validar_correspondencias(df_raw, columna_pais, columnas_valor)
    if errores:
        print('\n❌ Inconsistencias en la detección automática:')
        for e in errores:
            print('   -', e)
        columna_pais, columnas_valor = asignacion_manual(df_raw, columna_pais, list(columnas_valor))

year_columns = list(columnas_valor.keys())
print(f"\n✅ Columna de países: {columna_pais}")
print(f"✅ Columnas de años encontradas: {year_columns}")

# Transformar de formato ancho a formato largo (melt)
df_melted = df_raw.melt(
    id_vars=[columna_pais],
    value_vars=year_columns,
    var_name='year_column',
    value_name='trade_value_usd'
)

# Asignar año a cada columna
df_melted['year'] = df_melted['year_column'].map(columnas_valor).astype(int)
df_melted = df_melted.rename(columns={columna_pais: 'partner_country'})

# Limpiar datos
df_melted['partner_country'] = df_melted['partner_country'].astype(str).str.strip()
df_melted['trade_value_usd'] = _a_numero(df_melted['trade_value_usd'])
df = df_melted[['year', 'partner_country', 'trade_value_usd']].copy()

# Eliminar filas con datos faltantes y fila "Mundo" (o equivalentes: World, Total...)
df = df.dropna(subset=['trade_value_usd'])
df = df[~df['partner_country'].map(_normalizar).isin(FILAS_TOTAL | {'nan', ''})].reset_index(drop=True)

if df.empty:
    raise ValueError('❌ No quedaron datos después de la limpieza. Revise las columnas seleccionadas.')

print(f"\nDatos procesados: {len(df)} filas")
print(f"Países únicos: {df['partner_country'].nunique()}")
print(f"Años disponibles: {sorted(df['year'].unique())}")
print("\nPrimeros datos procesados:")
print(df.head(10))
```


## 4.3. Calcular Market Share, Ranking y aplicar filtros

```
# Calcular market share por año y país
market_evolution = df.groupby(['year', 'partner_country'])['trade_value_usd'].sum().reset_index()

# Calcular market share total por año
yearly_totals = market_evolution.groupby('year')['trade_value_usd'].sum()
market_evolution['market_share'] = market_evolution.apply(
    lambda x: (x['trade_value_usd'] / yearly_totals[x['year']]) * 100, axis=1
)

# Identificar top países
top_countries = df.groupby('partner_country')['trade_value_usd'].sum().nlargest(8).index

# Filtrar solo top países
market_evolution_top = market_evolution[market_evolution['partner_country'].isin(top_countries)]
```

## 4.4.  Generar Tablas Pivot para análisis

```
# TABLA 1: Market Share por año y país
print("📊 TABLA 1: MARKET SHARE POR AÑO (%)")
print("="*60)
pivot_share = market_evolution_top.pivot(index='year', columns='partner_country', values='market_share')
pivot_share_rounded = pivot_share.round(1)
print(pivot_share_rounded.to_string())

# TABLA 2: Valores de comercio (Millones USD)
print("\n\n💰 TABLA 2: VALORES DE COMERCIO (Millones USD)")
print("="*60)
pivot_value = market_evolution_top.pivot(index='year', columns='partner_country', values='trade_value_usd')
pivot_value_millions = (pivot_value / 1000).round(1)  # Convertir a millones
print(pivot_value_millions.to_string())

# TABLA 3: Crecimiento anual por país
print("\n\n📈 TABLA 3: CRECIMIENTO ANUAL COMPUESTO (CAGR %)")
print("="*60)
growth_data = market_evolution_top.groupby('partner_country').apply(
    lambda x: ((x['trade_value_usd'].iloc[-1] / x['trade_value_usd'].iloc[0]) ** (1/(len(x)-1)) - 1) * 100
).reset_index(name='cagr')
growth_summary = growth_data.set_index('partner_country')['cagr'].round(1).sort_values(ascending=False)
print(growth_summary.to_string())

# TABLA 4: Ranking por año
print("\n\n🏆 TABLA 4: RANKING DE PAÍSES POR AÑO (Top 5)")
print("="*60)
for year in sorted(df['year'].unique()):
    year_data = market_evolution_top[market_evolution_top['year'] == year].nlargest(5, 'market_share')
    print(f"\n{year}:")
    for i, (_, row) in enumerate(year_data.iterrows(), 1):
        print(f"  {i}. {row['partner_country']}: {row['market_share']:.1f}% ({row['trade_value_usd']/1000:.0f}M USD)")
```

## 4.5. Generar Gráficos analíticos

```
# Requiere las variables creadas en 4.1-4.4:
#   df, market_evolution_top, pivot_share, pivot_value, growth_data
# No modifica ningún cálculo: solo controla los colores de los gráficos.
import re

import matplotlib.colors as mcolors
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

# ------------------------------------------------------------
# 1) PALETAS PREDETERMINADAS (único lugar donde se definen colores)
# ------------------------------------------------------------
PALETAS_PREDETERMINADAS = {
    'Ejecutivo Clásico': {
        'descripcion': 'Azules sobrios, grises y acentos dorados. Informes ejecutivos y corporativos.',
        'colores': ['#1B365D', '#4A7BA7', '#8FB3D9', '#4A5568', '#A0AEC0', '#B8860B', '#D98E32', '#2C5282'],
    },
    'Comercio Global': {
        'descripcion': 'Azules y verdes con acentos vivos. Negocios internacionales y competitividad.',
        'colores': ['#0A3D62', '#1E88C8', '#00A896', '#2E9E5B', '#8BC34A', '#F4A300', '#E4572E', '#5C7AEA'],
    },
    'Corporativo Contemporáneo': {
        'descripcion': 'Neutros, tierras y acentos cálidos. Presentaciones analíticas modernas.',
        'colores': ['#3B3A36', '#8A7968', '#A9552B', '#D1A36B', '#5F6B4A', '#A7B58A', '#D9694F', '#E8C77A'],
    },
}
OPCION_PERSONALIZADA = 'Personalizada (pegar código CSS)'

# Una letra mayor que este umbral de luminosidad se considera demasiado clara sobre fondo blanco
UMBRAL_LUZ = 0.85
MIN_COLORES = 2
_REGEX_HEX = re.compile(r'#(?:[0-9a-fA-F]{8}|[0-9a-fA-F]{6}|[0-9a-fA-F]{3})(?![0-9a-fA-F])')


# ------------------------------------------------------------
# 2) EXTRACCIÓN Y VALIDACIÓN DE COLORES (el CSS NUNCA se ejecuta)
# ------------------------------------------------------------
def _luz(hex_color):
    r, g, b = mcolors.to_rgb(hex_color)
    return 0.299 * r + 0.587 * g + 0.114 * b


def extraer_colores_css(texto_css):
    """Extrae solo códigos HEX (#RGB, #RRGGBB, #RRGGBBAA) de un texto CSS.
    Devuelve (colores_validos, mensajes). colores_validos == [] si no es utilizable."""
    mensajes = []
    if not texto_css or not texto_css.strip():
        return [], ["El campo está vacío. Pegue el código CSS de su paleta."]

    texto = re.sub(r'/\*.*?\*/', ' ', texto_css, flags=re.S)       # ignorar comentarios CSS
    encontrados = _REGEX_HEX.findall(texto)
    if not encontrados:
        return [], ["No se encontraron colores HEX (ej. #1B365D) en el texto suministrado."]

    colores = []
    for c in encontrados:
        c = c.lower()
        if len(c) == 9:                                              # #rrggbbaa (Coolors): quitar alfa
            c = c[:7]
        elif len(c) == 4:                                            # #rgb -> #rrggbb
            c = '#' + ''.join(ch * 2 for ch in c[1:])
        if c not in colores:                                         # sin duplicados, conserva el orden
            colores.append(c)

    claros = [c for c in colores if _luz(c) > UMBRAL_LUZ]
    utiles = [c for c in colores if c not in claros]
    if claros:
        mensajes.append(f"Se omitieron {len(claros)} color(es) demasiado claros para fondo blanco: {', '.join(claros)}")
    if len(utiles) < MIN_COLORES:
        mensajes.append(f"Se necesitan al menos {MIN_COLORES} colores utilizables y solo hay {len(utiles)}.")
        return [], mensajes
    return utiles, mensajes


# ------------------------------------------------------------
# 3) ESQUEMA CROMÁTICO CENTRALIZADO
# ------------------------------------------------------------
def _mezclar(c1, c2, t):
    a, b = np.array(mcolors.to_rgb(c1)), np.array(mcolors.to_rgb(c2))
    return mcolors.to_hex(a * (1 - t) + b * t)


def _expandir(colores, n):
    """Devuelve n colores; si hay menos, genera tonos intermedios de la misma paleta."""
    if n <= len(colores):
        return list(colores[:n])
    cmap = mcolors.LinearSegmentedColormap.from_list('paleta', colores)
    return [mcolors.to_hex(cmap(x)) for x in np.linspace(0, 1, n)]


def construir_esquema(colores, paises):
    """Asigna un color fijo a cada país y deriva los tonos de texto, ejes y cuadrícula."""
    paises = list(dict.fromkeys(paises))                             # únicos, en orden
    asignados = _expandir(colores, len(paises))
    oscuro = min(colores, key=_luz)
    return {
        'por_pais': dict(zip(paises, asignados)),
        'texto': _mezclar(oscuro, '#000000', 0.25),
        'eje': _mezclar(oscuro, '#ffffff', 0.40),
        'grilla': _mezclar(oscuro, '#ffffff', 0.85),
        'referencia': _mezclar(oscuro, '#ffffff', 0.30),
        'principal': colores[0],
    }


# ------------------------------------------------------------
# 4) GRÁFICOS (mismos cálculos y estructura; solo cambian los colores)
# ------------------------------------------------------------
def dibujar_graficos(esquema):
    c = esquema['por_pais']
    estilo = {
        'text.color': esquema['texto'], 'axes.labelcolor': esquema['texto'],
        'axes.titlecolor': esquema['texto'], 'axes.edgecolor': esquema['eje'],
        'xtick.color': esquema['texto'], 'ytick.color': esquema['texto'],
        'grid.color': esquema['grilla'],
    }
    with plt.rc_context(estilo):
        fig, axes = plt.subplots(2, 2, figsize=(16, 14))

        # Market Share Trends con valores en puntos clave
        pivot_share.plot(kind='line', marker='o', ax=axes[0, 0], linewidth=2, markersize=6,
                         color=[c[p] for p in pivot_share.columns])
        axes[0, 0].set_title('Market Share Evolution (%)', fontsize=14, fontweight='bold')
        axes[0, 0].set_ylabel('Market Share (%)')
        axes[0, 0].grid(True, alpha=0.6)
        axes[0, 0].legend(bbox_to_anchor=(1.05, 1), loc='upper left')

        # Añadir valores para Brasil y Colombia (principales)
        for country in ['Brasil', 'Colombia']:
            if country in pivot_share.columns:
                for year in pivot_share.index:
                    value = pivot_share.loc[year, country]
                    if not pd.isna(value):
                        axes[0, 0].annotate(f'{value:.1f}%', (year, value),
                                            textcoords="offset points", xytext=(0, 10),
                                            ha='center', fontsize=8, color=esquema['texto'])

        # Trade Value Evolution
        (pivot_value / 1000).plot(kind='line', marker='s', ax=axes[0, 1], linewidth=2, markersize=6,
                                  color=[c[p] for p in pivot_value.columns])
        axes[0, 1].set_title('Trade Value Evolution (Million USD)', fontsize=14, fontweight='bold')
        axes[0, 1].set_ylabel('Million USD')
        axes[0, 1].grid(True, alpha=0.6)

        # Market Share actual con valores
        latest_year = df['year'].max()
        latest_data = market_evolution_top[market_evolution_top['year'] == latest_year].nlargest(8, 'market_share')
        serie_ms = latest_data.set_index('partner_country')['market_share']
        bars = serie_ms.plot(kind='bar', ax=axes[1, 0], color=[c[p] for p in serie_ms.index])
        axes[1, 0].set_title(f'Market Share {latest_year} (%)', fontsize=14, fontweight='bold')
        axes[1, 0].set_ylabel('Market Share (%)')
        axes[1, 0].tick_params(axis='x', rotation=45)

        # Añadir valores encima de las barras
        for i, (country, value) in enumerate(serie_ms.items()):
            axes[1, 0].text(i, value + 0.5, f'{value:.1f}%', ha='center', va='bottom',
                            fontsize=9, color=esquema['texto'])

        # CAGR con valores
        serie_cagr = growth_data.set_index('partner_country')['cagr']
        bars2 = serie_cagr.plot(kind='bar', ax=axes[1, 1], color=[c[p] for p in serie_cagr.index])
        axes[1, 1].set_title('CAGR % (Compound Annual Growth Rate)', fontsize=14, fontweight='bold')
        axes[1, 1].set_ylabel('CAGR (%)')
        axes[1, 1].tick_params(axis='x', rotation=45)
        axes[1, 1].axhline(y=0, color=esquema['referencia'], linestyle='-', alpha=0.6)

        # Añadir valores encima de las barras
        for i, (country, value) in enumerate(serie_cagr.items()):
            axes[1, 1].text(i, value + 0.5 if value > 0 else value - 1, f'{value:.1f}%',
                            ha='center', va='bottom' if value > 0 else 'top',
                            fontsize=9, color=esquema['texto'])

        plt.tight_layout()
        plt.show()
    return fig, axes, bars, bars2, latest_year, latest_data


def _todos_los_paises():
    """Orden estable de países para que cada uno tenga el mismo color en los 4 gráficos."""
    orden = list(top_countries) if 'top_countries' in globals() else []
    for nombre in (list(pivot_share.columns) + list(pivot_value.columns)
                   + list(growth_data['partner_country'])):
        orden.append(nombre)
    return list(dict.fromkeys(orden))


def generar_con_paleta(colores):
    global fig, axes, bars, bars2, latest_year, latest_data
    esquema = construir_esquema(colores, _todos_los_paises())
    fig, axes, bars, bars2, latest_year, latest_data = dibujar_graficos(esquema)


# ------------------------------------------------------------
# 5) INTERFAZ INTERACTIVA (Google Colab / Jupyter)
# ------------------------------------------------------------
def _html_vista_previa(colores):
    cuadros = ''.join(
        f'<span style="display:inline-block;width:34px;height:20px;background:{c};'
        f'border:1px solid #888;margin-right:2px;" title="{c}"></span>' for c in colores)
    return f'<div style="margin:2px 0 8px 0">{cuadros}</div>'


def mostrar_selector_paletas():
    try:
        import ipywidgets as widgets
        from IPython.display import HTML, clear_output, display
    except ImportError:
        print("ℹ️ No es posible mostrar la interfaz interactiva en este entorno. "
              "Se usa la paleta 'Ejecutivo Clásico'.")
        generar_con_paleta(PALETAS_PREDETERMINADAS['Ejecutivo Clásico']['colores'])
        return

    opciones = list(PALETAS_PREDETERMINADAS) + [OPCION_PERSONALIZADA]
    selector = widgets.RadioButtons(options=opciones, value=opciones[0], description='Paleta:',
                                    layout=widgets.Layout(width='500px'),
                                    style={'description_width': '60px'})
    vista = widgets.HTML()
    css = widgets.Textarea(
        placeholder=("Pegue aquí el código CSS de su paleta (ej. Coolors -> Export -> CSS):\n"
                     ":root {\n  --color-1: #264653;\n  --color-2: #2a9d8f;\n  --color-3: #e76f51;\n}"),
        layout=widgets.Layout(width='620px', height='140px', display='none'))
    boton = widgets.Button(description='Generar gráficos', button_style='success', icon='bar-chart')
    salida = widgets.Output()

    def refrescar(*_):
        nombre = selector.value
        if nombre == OPCION_PERSONALIZADA:
            css.layout.display = 'block'
            vista.value = "<i>Pegue el CSS con colores HEX y pulse «Generar gráficos».</i>"
        else:
            css.layout.display = 'none'
            p = PALETAS_PREDETERMINADAS[nombre]
            vista.value = f"<b>{nombre}</b>: {p['descripcion']}" + _html_vista_previa(p['colores'])

    selector.observe(refrescar, 'value')
    refrescar()

    def al_pulsar(_):
        with salida:
            clear_output(wait=True)
            nombre = selector.value
            if nombre == OPCION_PERSONALIZADA:
                colores, avisos = extraer_colores_css(css.value)
                for a in avisos:
                    print("⚠️", a)
                if not colores:
                    print("❌ No se pudo construir una paleta utilizable. Revise el código CSS e intente de nuevo.")
                    return
                display(HTML("<b>Paleta personalizada detectada:</b>" + _html_vista_previa(colores)))
            else:
                colores = PALETAS_PREDETERMINADAS[nombre]['colores']
            try:
                generar_con_paleta(colores)
            except NameError as e:
                print(f"❌ Faltan variables de los apartados 4.3 y 4.4 ({e}). Ejecútelos primero.")

    boton.on_click(al_pulsar)
    display(widgets.VBox([widgets.HTML("<b>🎨 Seleccione la paleta de colores de los gráficos</b>"),
                          selector, vista, css, boton, salida]))


mostrar_selector_paletas()
# Tras pulsar "Generar gráficos" quedan disponibles: fig, axes, bars, bars2, latest_year, latest_data
```

## 4.7. Generación de Promts para análisis posteriores

**Nota:** *Reemplazar "market_name" (Line 7) por el mercado objetivo (país-destino) antes de proceder con la ejecución del script.*

```
# GENERAR PROMPT AUTOMÁTICO PARA ANÁLISIS
def generate_analysis_prompt(df, market_evolution_top, product_name):
    latest_year = df['year'].max()
    first_year = df['year'].min()

    # Detectar mercado objetivo (desde columnas o contexto)
    market_name = "Estados Unidos"  # Por defecto, se podría dinamizar

    # Generar estadísticas clave
    top3 = market_evolution_top[market_evolution_top['year'] == latest_year].nlargest(3, 'market_share')
    total_market = df[df['year'] == latest_year]['trade_value_usd'].sum()

    prompt = f"""
    Actúa como consultor en inteligencia de mercados y analiza estos datos de market share de {product_name.lower()} en {market_name} ({first_year}-{latest_year}). Proporciona insights estratégicos específicos:

    CONTEXTO DEL MERCADO:
    - Producto: {product_name}
    - Mercado objetivo: {market_name}
    - Período: {first_year}-{latest_year}
    - Tamaño total del mercado ({latest_year}): ${total_market/1000:.0f} millones USD

    TOP 3 LÍDERES ACTUALES ({latest_year}):
    """

    for i, (_, row) in enumerate(top3.iterrows(), 1):
        prompt += f"\n    {i}. {row['partner_country']}: {row['market_share']:.1f}% market share"

    prompt += f"""

    DATOS PARA ANÁLISIS:
    [Pegar aquí las 4 tablas generadas por el script anterior]

    PREGUNTAS ESPECÍFICAS A RESPONDER:

    1. ANÁLISIS COMPETITIVO:
       - ¿Cuáles son las tendencias de los líderes del mercado?
       - ¿Qué países están ganando/perdiendo participación?
       - ¿Hay oportunidades por concentración/fragmentación del mercado?

    2. OPORTUNIDADES DE ENTRADA:
       - ¿Qué nichos o segmentos presentan menos competencia?
       - ¿Existen países que han reducido su participación creando vacíos?
       - ¿Cuál sería el mejor timing para entrar al mercado?

    3. BENCHMARKING REGIONAL:
       - ¿Cómo se comportan países con perfil similar al de tu empresa?
       - ¿Qué países de América Latina tienen mejor performance?
       - ¿Cuáles son las mejores prácticas observables?

    4. ESTRATEGIA DE POSICIONAMIENTO:
       - ¿Competir directamente con líderes o buscar nichos?
       - ¿Qué ventajas competitivas serían más relevantes?
       - ¿Cuál sería un market share objetivo realista en 2-3 años?

    5. RIESGOS Y AMENAZAS:
       - ¿Qué competidores emergentes representan amenazas?
       - ¿Hay signos de saturación o madurez del mercado?
       - ¿Qué factores externos podrían afectar la industria?

    FORMATO DE RESPUESTA:
    - Insights concisos y accionables
    - Recomendaciones estratégicas específicas
    - Identificación de 2-3 oportunidades prioritarias
    - Métricas clave a monitorear
    """

    return prompt

# Generar y mostrar el prompt
analysis_prompt = generate_analysis_prompt(df, market_evolution_top, product_name)
print("="*80)
print("📋 PROMPT PARA ANÁLISIS ESTRATÉGICO:")
print("="*80)
print(analysis_prompt)
print("\n" + "="*80)
print("💡 INSTRUCCIONES:")
print("1. Copia este prompt")
print("2. Pega las 4 tablas de datos generadas arriba")
print("3. Úsalo con cualquier chatbot para obtener insights estratégicos")
print("="*80)
```

## 4.8. Helpers (opcionales) para depuración y personalización
### 4.8.1. Promt para análisis de oportunidad: Origen (Proveedor)-Destino

```
Con base en lo anterior, redacta un texto argumentativo, en tono formal y académico en tercera persona, máximo una cuartilla (250 palabras) donde evidencies los resultados de este estudio de mercados. Enfatiza en el nivel de oportunidad que tendría [COLOMBIA] para ingresar a este mercado desde el punto de vista comercial basándose únicamente en los datos suministrados.
```

### 4.8.2. Promt para personalizar colores en los gráficos

```
Con base en la siguiente paleta de colores:

[CSS.CODE]

Ajusta el siguiente trozo de código en consecuencia, da por hecho que las librerías ya se encuentran instaladas:

[PY.CODE.4.5]
```

### 4.8.3. Ajuste para datos faltantes o sin datos históricos (No calcula CAGR)

```
# TABLA 1: Market Share por año y país
print("📊 TABLA 1: MARKET SHARE POR AÑO (%)")
print("="*60)
pivot_share = market_evolution_top.pivot(index='year', columns='partner_country', values='market_share')
pivot_share_rounded = pivot_share.round(1)
print(pivot_share_rounded.to_string())

# TABLA 2: Valores de comercio (Millones USD)
print("\n\n💰 TABLA 2: VALORES DE COMERCIO (Millones USD)")
print("="*60)
pivot_value = market_evolution_top.pivot(index='year', columns='partner_country', values='trade_value_usd')
pivot_value_millions = (pivot_value / 1000).round(1)  # Convertir a millones
print(pivot_value_millions.to_string())

# TABLA 3: Crecimiento anual por país
# El cálculo del CAGR requiere al menos dos años de datos. Como solo se tiene un año (2024),
# esta tabla no se puede generar con los datos actuales.
# print("\n\n📈 TABLA 3: CRECIMIENTO ANUAL COMPUESTO (CAGR %)")
# print("="*60)
# growth_data = market_evolution_top.groupby('partner_country').apply(
#     lambda x: ((x['trade_value_usd']. iloc[-1] / x['trade_value_usd'].iloc[0]) ** (1/(len(x)-1)) - 1) * 100
# ).reset_index(name='cagr')
# growth_summary = growth_data.set_index('partner_country')['cagr'].round(1).sort_values(ascending=False)
# print(growth_summary.to_string())
print("\n\n📈 TABLA 3: CRECIMIENTO ANUAL COMPUESTO (CAGR %)")
print("="*60)
print("No se puede calcular el CAGR. Se requiere al menos dos años de datos para este cálculo.")

# TABLA 4: Ranking por año
print("\n\n🏆 TABLA 4: RANKING DE PAÍSES POR AÑO (Top 5)")
print("="*60)
for year in sorted(df['year'].unique()):
    year_data = market_evolution_top[market_evolution_top['year'] == year].nlargest(5, 'market_share')
    print(f"\n{year}:")
    for i, (_, row) in enumerate(year_data.iterrows(), 1):
        print(f"  {i}. {row['partner_country']}: {row['market_share']:.1f}% ({row['trade_value_usd']/1000:.0f}M USD)")
```

### 4.8.4. Ajuste para datos faltantes (Calcula CAGR omitiendo datos vacíos o nulos)

```
# TABLA 1: Market Share por año y país
print("📊 TABLA 1: MARKET SHARE POR AÑO (%)")
print("="*60)
pivot_share = market_evolution_top.pivot(index='year', columns='partner_country', values='market_share')
pivot_share_rounded = pivot_share.round(1)
print(pivot_share_rounded.to_string())

# TABLA 2: Valores de comercio (Millones USD)
print("\n\n💰 TABLA 2: VALORES DE COMERCIO (Millones USD)")
print("="*60)
pivot_value = market_evolution_top.pivot(index='year', columns='partner_country', values='trade_value_usd')
pivot_value_millions = (pivot_value / 1000).round(1)  # Convertir a millones
print(pivot_value_millions.to_string())

# TABLA 3: Crecimiento anual por país
print("\n\n📈 TABLA 3: CRECIMIENTO ANUAL COMPUESTO (CAGR %)")
print("="*60)
growth_data = market_evolution_top.groupby('partner_country').apply(
    lambda x: ((x['trade_value_usd'].iloc[-1] / x['trade_value_usd'].iloc[0]) ** (1/(len(x)-1)) - 1) * 100 if len(x) > 1 else np.nan
).reset_index(name='cagr')
growth_summary = growth_data.set_index('partner_country')['cagr'].round(1).sort_values(ascending=False)
print(growth_summary.to_string())

# TABLA 4: Ranking por año
print("\n\n🏆 TABLA 4: RANKING DE PAÍSES POR AÑO (Top 5)")
print("="*60)
for year in sorted(df['year'].unique()):
    year_data = market_evolution_top[market_evolution_top['year'] == year].nlargest(5, 'market_share')
    print(f"\n{year}:")
    for i, (_, row) in enumerate(year_data.iterrows(), 1):
        print(f"  {i}. {row['partner_country']}: {row['market_share']:.1f}% ({row['trade_value_usd']/1000:.0f}M USD)")
```

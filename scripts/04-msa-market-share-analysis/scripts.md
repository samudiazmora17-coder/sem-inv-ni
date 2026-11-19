# 4. SCRIPTS MSA (MARKET SHARE ANALYSIS)

**Nota:** *Los siguientes scripts están diseñados para el análisis de datos orientado a la estimación de market share, la generación de visualizaciones analíticas y la identificación de insights estratégicos a partir de datos históricos obtenidos de TradeMap. Se sugiere implementar este flujo de trabajo en entornos de ejecución Python tales como:*

- [Google Colab](https://colab.research.google.com/)
- [Python-Fiddle](https://python-fiddle.com/)
- [Kaggle Notebooks](https://www.kaggle.com/code)
- [Deepnote](https://deepnote.com/)

4.1. Importar e instalar librerías
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

4.2. Cargar archivos fuente (Trademap)
Nota: El script reconoce automáticamente las columnas de países y de valores por año, con cabeceras en español o en inglés. Si la detección no es concluyente, pedirá elegir las columnas por su número. Si el botón de subida no funciona, suba el archivo desde el panel de carpetas de Colab y escriba su nombre en NOMBRE_ARCHIVO.

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

4.3. Calcular Market Share, Ranking y aplicar filtros
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

4.4. Generar Tablas Pivot para análisis
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
4.5. Generar Gráficos analíticos
Nota: Al ejecutar la celda aparece un selector con tres paletas (Ejecutivo Clásico, Comercio Global y Corporativo Contemporáneo) y una opción personalizada para pegar el código CSS exportado de Coolors. Pulse "Generar gráficos" para aplicar la paleta a los cuatro gráficos. El código CSS no se ejecuta: solo se extraen los colores HEX.

# Gráficos analíticos con selector de paleta de colores
import re
import matplotlib.colors as mcolors

# ---------- 1) DEFINICIÓN CENTRALIZADA DE COLORES ----------
# Para cambiar o agregar una paleta basta con editar este diccionario.
PALETAS_PREDETERMINADAS = {
    'Ejecutivo Clásico': {
        'descripcion': 'Azules, grises y acentos sobrios. Para informes ejecutivos y presentaciones corporativas.',
        'colores': ['#1F4E79', '#D08C2E', '#2A9D8F', '#8E3B46', '#5B9BD5', '#7F6A3B', '#9B7EBD', '#4A5560'],
    },
    'Comercio Global': {
        'descripcion': 'Azules, verdes y complementarios. Evoca negocios internacionales y mercados globales.',
        'colores': ['#0E5A8A', '#E07B39', '#4FB3D9', '#2E9E5B', '#7A4EA3', '#C9474B', '#5A6C7D', '#88A93B'],
    },
    'Corporativo Contemporáneo': {
        'descripcion': 'Neutros, tierras y acentos cálidos. Para reportes analíticos y comerciales modernos.',
        'colores': ['#4A4E69', '#D9774B', '#3F7CAC', '#C9A227', '#7B5E8F', '#B5473A', '#5F8A6E', '#3B3A36'],
    },
}
OPCION_PERSONALIZADA = 'Personalizada (pegar código CSS)'
PALETA_INICIAL = 'Ejecutivo Clásico'
MINIMO_COLORES = 3

# Colores de apoyo (texto, ejes, cuadrícula): iguales para todas las paletas
COLORES_APOYO = {'texto': '#222B36', 'texto_secundario': '#5F6B76', 'cuadricula': '#C9D1D9',
                 'linea_cero': '#5F6B76', 'borde_barra': '#FFFFFF'}

PALETA_ACTIVA = None   # se actualiza cada vez que se generan los gráficos


# ---------- 2) LECTURA Y VALIDACIÓN DE PALETAS CSS ----------
PATRON_HEX = re.compile(r'#([0-9a-fA-F]{8}|[0-9a-fA-F]{6}|[0-9a-fA-F]{4}|[0-9a-fA-F]{3})(?![0-9a-fA-F])')


def _contraste_con_blanco(color_hex):
    r, g, b = [c / 12.92 if c <= 0.04045 else ((c + 0.055) / 1.055) ** 2.4 for c in mcolors.to_rgb(color_hex)]
    return 1.05 / (0.2126 * r + 0.7152 * g + 0.0722 * b + 0.05)


def extraer_colores_css(codigo_css):
    """Extrae los colores HEX de un texto CSS (p. ej. exportado de Coolors).
    El texto NUNCA se ejecuta: solo se buscan patrones #RGB, #RRGGBB o #RRGGBBAA.
    Devuelve (lista_de_colores, lista_de_avisos) o lanza ValueError con un mensaje claro."""
    texto = str(codigo_css or '')
    if not texto.strip():
        raise ValueError('❌ El campo está vacío. Pegue el código CSS de su paleta '
                         '(por ejemplo: --azul: #003049;).')
    colores = []
    for coincidencia in PATRON_HEX.findall(texto):
        h = coincidencia.upper()
        if len(h) in (3, 4):                 # #RGB / #RGBA  -> #RRGGBB
            h = ''.join(c * 2 for c in h[:3])
        h = '#' + h[:6]                      # se descarta la transparencia (AA)
        if h not in colores:                 # Coolors repite los colores en varias secciones
            colores.append(h)
    if not colores:
        pista = (' Se detectaron colores en formato rgb()/hsl(): en Coolors elija Export → CSS y '
                 'pegue la sección "CSS HEX".') if re.search(r'(rgb|hsl)a?\(', texto, re.I) else ''
        raise ValueError('❌ No se encontraron colores HEX válidos en el texto pegado. '
                         'Se esperan valores como #1F4E79 o #1F4E79FF.' + pista)
    avisos = []
    claros = [c for c in colores if _contraste_con_blanco(c) < 1.25]
    if claros:
        colores = [c for c in colores if c not in claros]
        avisos.append('Se omitieron colores demasiado claros para un fondo blanco: ' + ', '.join(claros))
    if len(colores) < MINIMO_COLORES:
        raise ValueError(f'❌ La paleta solo tiene {len(colores)} color(es) utilizable(s) y se necesitan '
                         f'al menos {MINIMO_COLORES}.' + (' ' + avisos[0] if avisos else ''))
    return colores, avisos


def construir_paleta(nombre, codigo_css=None):
    """Devuelve la paleta lista para aplicar: {'nombre', 'colores', 'avisos'}."""
    if nombre in PALETAS_PREDETERMINADAS:
        return {'nombre': nombre, 'colores': list(PALETAS_PREDETERMINADAS[nombre]['colores']), 'avisos': []}
    colores, avisos = extraer_colores_css(codigo_css)
    return {'nombre': 'Personalizada', 'colores': colores, 'avisos': avisos}


def _mezclar(color_hex, destino, proporcion):
    base, meta = np.array(mcolors.to_rgb(color_hex)), np.array(mcolors.to_rgb(destino))
    return mcolors.to_hex(base + (meta - base) * proporcion).upper()


def asignar_colores(paises, colores):
    """Asigna un color fijo a cada país (el mismo en los cuatro gráficos).
    Si hay más países que colores, se derivan tonos claros/oscuros y se cambia el estilo de línea."""
    variantes = [(None, 0, '-'), ('#FFFFFF', 0.45, '--'), ('#000000', 0.35, ':'), ('#FFFFFF', 0.70, '-.')]
    mapa_color, mapa_estilo = {}, {}
    for i, pais in enumerate(paises):
        destino, proporcion, estilo = variantes[(i // len(colores)) % len(variantes)]
        base = colores[i % len(colores)]
        mapa_color[pais] = base if destino is None else _mezclar(base, destino, proporcion)
        mapa_estilo[pais] = estilo
    return mapa_color, mapa_estilo


# ---------- 3) GENERACIÓN DE LOS CUATRO GRÁFICOS ----------
def generar_graficos(paleta):
    """Dibuja los 4 gráficos con la paleta indicada. No modifica ningún dato ni cálculo."""
    global PALETA_ACTIVA
    PALETA_ACTIVA = paleta
    apoyo = COLORES_APOYO

    # Orden fijo de países (de mayor a menor valor total) -> color estable por país
    paises = pivot_value.sum().sort_values(ascending=False).index.tolist()
    paises += [p for p in pivot_share.columns if p not in paises]
    if 'growth_data' in globals():
        paises += [p for p in growth_data['partner_country'] if p not in paises]
    color_pais, estilo_pais = asignar_colores(paises, paleta['colores'])

    def dar_formato(ax, titulo, etiqueta_y):
        ax.set_title(titulo, fontsize=14, fontweight='bold', color=apoyo['texto'])
        ax.set_ylabel(etiqueta_y, color=apoyo['texto_secundario'])
        ax.set_xlabel(ax.get_xlabel(), color=apoyo['texto_secundario'])
        ax.tick_params(colors=apoyo['texto_secundario'])
        ax.grid(True, alpha=0.5, color=apoyo['cuadricula'])
        ax.set_axisbelow(True)
        ax.margins(y=0.12)   # espacio para que las etiquetas no choquen con el título

    def dibujar_lineas(tabla, ax, marcador):
        tabla.plot(kind='line', marker=marcador, ax=ax, linewidth=2, markersize=6,
                   color=[color_pais[c] for c in tabla.columns],
                   style=None)
        for linea, pais in zip(ax.get_lines(), tabla.columns):
            linea.set_linestyle(estilo_pais[pais])
        ax.set_xticks(list(tabla.index))
        ax.legend(bbox_to_anchor=(1.02, 1), loc='upper left', frameon=False,
                  labelcolor=apoyo['texto'])

    # Gráficos mejorados con valores
    fig, axes = plt.subplots(2, 2, figsize=(16, 14))
    fig.patch.set_facecolor('white')

    # Market Share Trends con valores en puntos clave
    dibujar_lineas(pivot_share, axes[0, 0], 'o')
    dar_formato(axes[0, 0], 'Market Share Evolution (%)', 'Market Share (%)')

    # Añadir valores para Brasil y Colombia (principales)
    for country in ['Brasil', 'Colombia']:
        if country in pivot_share.columns:
            for year in pivot_share.index:
                value = pivot_share.loc[year, country]
                if not pd.isna(value):
                    axes[0, 0].annotate(f'{value:.1f}%',
                                        (year, value),
                                        textcoords="offset points",
                                        xytext=(0, 10),
                                        ha='center', fontsize=8, color=apoyo['texto'])

    # Trade Value Evolution
    dibujar_lineas(pivot_value / 1000, axes[0, 1], 's')
    dar_formato(axes[0, 1], 'Trade Value Evolution (Million USD)', 'Million USD')

    # Market Share actual con valores
    latest_year = df['year'].max()
    latest_data = market_evolution_top[market_evolution_top['year'] == latest_year].nlargest(8, 'market_share')
    serie_share = latest_data.set_index('partner_country')['market_share']
    bars = serie_share.plot(kind='bar', ax=axes[1, 0], width=0.7,
                            color=[color_pais[p] for p in serie_share.index],
                            edgecolor=apoyo['borde_barra'], linewidth=1.5)
    dar_formato(axes[1, 0], f'Market Share {latest_year} (%)', 'Market Share (%)')
    axes[1, 0].tick_params(axis='x', rotation=45)
    axes[1, 0].grid(False, axis='x')

    # Añadir valores encima de las barras
    for i, (country, value) in enumerate(serie_share.items()):
        axes[1, 0].text(i, value + 0.5, f'{value:.1f}%', ha='center', va='bottom',
                        fontsize=9, color=apoyo['texto'])

    # CAGR con valores
    serie_cagr = None
    if 'growth_data' in globals():
        serie_cagr = growth_data.set_index('partner_country')['cagr'].replace([np.inf, -np.inf], np.nan)
    if serie_cagr is not None and serie_cagr.notna().any():
        bars2 = serie_cagr.plot(kind='bar', ax=axes[1, 1], width=0.7,
                                color=[color_pais[p] for p in serie_cagr.index],
                                edgecolor=apoyo['borde_barra'], linewidth=1.5)
        axes[1, 1].axhline(y=0, color=apoyo['linea_cero'], linestyle='-', linewidth=1, alpha=0.8)

        # Añadir valores encima de las barras
        for i, (country, value) in enumerate(serie_cagr.items()):
            if pd.isna(value):
                continue
            axes[1, 1].text(i, value + 0.5 if value > 0 else value - 1, f'{value:.1f}%',
                            ha='center', va='bottom' if value > 0 else 'top',
                            fontsize=9, color=apoyo['texto'])
    else:
        axes[1, 1].text(0.5, 0.5, 'CAGR no disponible\n(se requieren al menos dos años de datos)',
                        ha='center', va='center', transform=axes[1, 1].transAxes,
                        color=apoyo['texto_secundario'], fontsize=11)
        axes[1, 1].set_xticks([])
    dar_formato(axes[1, 1], 'CAGR % (Compound Annual Growth Rate)', 'CAGR (%)')
    axes[1, 1].tick_params(axis='x', rotation=45)
    axes[1, 1].grid(False, axis='x')

    fig.suptitle(f'Paleta: {paleta["nombre"]}', fontsize=10, color=apoyo['texto_secundario'],
                 x=0.99, y=0.995, ha='right')
    plt.tight_layout()
    plt.show()
    return fig


# ---------- 4) INTERFAZ DE SELECCIÓN ----------
def _muestra_html(colores):
    cuadros = ''.join(f'<span title="{c}" style="display:inline-block;width:34px;height:20px;'
                      f'background:{c};margin-right:3px;border-radius:3px"></span>' for c in colores)
    return f'<div style="margin:4px 0 2px 0">{cuadros}</div>'


def selector_de_paleta():
    """Muestra el selector (3 paletas + personalizada) y genera los gráficos al pulsar el botón."""
    try:
        import ipywidgets as widgets
        from IPython.display import display, clear_output
    except ImportError:
        return _selector_por_texto()

    opciones = list(PALETAS_PREDETERMINADAS) + [OPCION_PERSONALIZADA]
    selector = widgets.RadioButtons(options=opciones, value=PALETA_INICIAL,
                                    layout=widgets.Layout(width='auto'))
    vista_previa = widgets.HTML()
    campo_css = widgets.Textarea(
        placeholder='Pegue aquí el código CSS de su paleta. Ejemplo (Coolors → Export → CSS):\n'
                    ':root {\n  --color-1: #780000;\n  --color-2: #c1121f;\n  --color-3: #003049;\n}',
        layout=widgets.Layout(width='620px', height='150px', display='none'))
    boton = widgets.Button(description='Generar gráficos', button_style='primary', icon='paint-brush')
    salida = widgets.Output()

    def actualizar_vista(_=None):
        if selector.value == OPCION_PERSONALIZADA:
            campo_css.layout.display = 'block'
            vista_previa.value = ('<i>Pegue el código CSS en el recuadro y pulse '
                                  '<b>Generar gráficos</b>. Se necesitan al menos '
                                  f'{MINIMO_COLORES} colores HEX.</i>')
        else:
            campo_css.layout.display = 'none'
            info = PALETAS_PREDETERMINADAS[selector.value]
            vista_previa.value = (_muestra_html(info['colores']) +
                                  f'<i>{info["descripcion"]}</i><br><small>{" · ".join(info["colores"])}</small>')

    def al_pulsar(_=None):
        with salida:
            clear_output(wait=True)
            try:
                paleta = construir_paleta(selector.value, campo_css.value)
            except ValueError as error:
                print(error)
                print('ℹ️ No se generaron los gráficos. Corrija el código CSS o elija una paleta predeterminada.')
                return
            for aviso in paleta['avisos']:
                print('⚠️', aviso)
            print(f'🎨 Paleta aplicada: {paleta["nombre"]} → {", ".join(paleta["colores"])}')
            generar_graficos(paleta)

    selector.observe(actualizar_vista, names='value')
    boton.on_click(al_pulsar)
    actualizar_vista()
    display(widgets.VBox([widgets.HTML('<b>🎨 Elija la paleta de colores de los gráficos</b>'),
                          selector, vista_previa, campo_css, boton, salida]))
    al_pulsar()   # primera ejecución con la paleta inicial


def _selector_por_texto():
    """Alternativa para entornos sin ipywidgets: menú numerado."""
    opciones = list(PALETAS_PREDETERMINADAS) + [OPCION_PERSONALIZADA]
    print('🎨 Elija la paleta de colores de los gráficos:')
    for i, nombre in enumerate(opciones, 1):
        detalle = PALETAS_PREDETERMINADAS.get(nombre, {}).get('descripcion', 'Use su propia paleta (Coolors u otra).')
        print(f'  {i}. {nombre}: {detalle}')
    while True:
        respuesta = input(f'👉 Escriba el número de la opción (1-{len(opciones)}): ').strip()
        if respuesta.isdigit() and 1 <= int(respuesta) <= len(opciones):
            nombre = opciones[int(respuesta) - 1]
            break
        print('   ⚠️ Opción no válida.')
    while True:
        css = None
        if nombre == OPCION_PERSONALIZADA:
            print('👉 Pegue el código CSS de la paleta y presione Enter en una línea vacía para terminar:')
            lineas = []
            while True:
                linea = input()
                if not linea.strip():
                    break
                lineas.append(linea)
            css = '\n'.join(lineas)
        try:
            paleta = construir_paleta(nombre, css)
            break
        except ValueError as error:
            print(error)
    for aviso in paleta['avisos']:
        print('⚠️', aviso)
    print(f'🎨 Paleta aplicada: {paleta["nombre"]} → {", ".join(paleta["colores"])}')
    generar_graficos(paleta)


selector_de_paleta()
4.6. Generar Insights estratégicos o "Descubrimientos relevantes"
Nota: Reemplazar "producto analizado" (Line 2) y "Producto" (Line 35) por el número de partida y nombre del producto analizado antes de proceder con la ejecución del script.

# Generar insights automáticos
def generate_insights(df, market_evolution_top, product_name="producto analizado"):
    latest_year = df['year'].max()
    first_year = df['year'].min()

    print(f"📊 INSIGHTS ESTRATÉGICOS - {product_name.upper()}")
    print("="*60)

    # Top 3 países
    top3 = market_evolution_top[market_evolution_top['year'] == latest_year].nlargest(3, 'market_share')
    print(f"\n🏆 TOP 3 PAÍSES ({latest_year}):")
    for i, (_, row) in enumerate(top3.iterrows(), 1):
        print(f"  {i}. {row['partner_country']}: {row['market_share']:.1f}% market share")

    # Mayor crecimiento
    growth_analysis = market_evolution_top.groupby('partner_country').apply(
        lambda x: x['trade_value_usd'].iloc[-1] / x['trade_value_usd'].iloc[0] - 1
    ).sort_values(ascending=False)

    print(f"\n📈 MAYOR CRECIMIENTO ({first_year}-{latest_year}):")
    print(f"  {growth_analysis.index[0]}: +{growth_analysis.iloc[0]:.1%}")

    # Oportunidades
    declining = growth_analysis[growth_analysis < 0]
    if len(declining) > 0:
        print(f"\n⚠️  MERCADOS EN DECLIVE:")
        for country, decline in declining.head(3).items():
            print(f"  {country}: {decline:.1%}")

    # Tamaño total del mercado
    total_market = df[df['year'] == latest_year]['trade_value_usd'].sum()
    print(f"\n💰 TAMAÑO TOTAL DEL MERCADO ({latest_year}): ${total_market/1e9:.1f}B USD")

# Define the product_name variable
product_name = "Producto"

# Aplicar análisis con nombre de producto dinámico
generate_insights(df, market_evolution_top, product_name)

4.7. Generación de Promts para análisis posteriores
Nota: Reemplazar "market_name" (Line 7) por el mercado objetivo (país-destino) antes de proceder con la ejecución del script.

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

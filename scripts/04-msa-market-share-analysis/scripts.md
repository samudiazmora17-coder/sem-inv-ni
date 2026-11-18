# ==============================================================
# 4.2. CARGAR ARCHIVOS FUENTE (TRADE MAP)
# Detección flexible de columnas (español / inglés) + asignación manual
# ==============================================================

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
    texto = serie.astype(str).str.replace(' ', '', regex=False).str.replace(' ', '', regex=False).str.strip()
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

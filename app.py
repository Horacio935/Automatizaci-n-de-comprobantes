import io

import streamlit as st
import pandas as pd
import streamlit.components.v1 as components

# --------------------------------------------------------------------
#  Funciones auxiliares
# --------------------------------------------------------------------

def cargar_planilla(archivo_excel: io.BytesIO) -> pd.DataFrame:
    """
    Detecta automáticamente la hoja que empieza con 'PLANILLA SECT'
    y la carga aunque el sector sea diferente.
    """
    # Obtener lista de hojas del archivo
    xls = pd.ExcelFile(archivo_excel)
    hojas = xls.sheet_names

    # Buscar hoja que empiece con 'PLANILLA SECT'
    hoja_encontrada = None
    for h in hojas:
        if str(h).strip().upper().startswith("PLANILLA SECT"):
            hoja_encontrada = h
            break

    if hoja_encontrada is None:
        raise ValueError("No se encontró ninguna hoja que empiece con 'PLANILLA SECT'.")

    # Cargar la hoja encontrada
    df = pd.read_excel(
        archivo_excel,
        sheet_name=hoja_encontrada,
        header=5  # misma estructura
    )

    # Filtrar solo filas válidas
    df = df[df["Corr"].notna()].copy()
    df = df[df["NOMBRE DEL EMPLEADO"].notna()]
    df = df[df["NOMBRE DEL EMPLEADO"].astype(str).str.strip() != "nan"]
    df = df[df["NOMBRE DEL EMPLEADO"].astype(str).str.strip() != ""]

    # Limpieza
    df["Corr"] = df["Corr"].astype(int)
    df["NOMBRE DEL EMPLEADO"] = df["NOMBRE DEL EMPLEADO"].astype(str).str.strip()

    return df


def safe_float(value):
    try:
        if pd.isna(value):
            return 0.0
        return float(value)
    except:
        return 0.0

def fila_a_diccionario(row: pd.Series, periodo_texto: str) -> dict:
    """
    Convierte una fila de la planilla en un diccionario listo para
    usar en el comprobante.
    """
    # Datos base
    sueldo_base_mensual = safe_float(row["SALARIO BASE MENSUAL"])
    devengado = safe_float(row["SUELDO BASE SEGÚN DIAS TRABAJADOS"])
    total_ingresos = safe_float(row["TOTAL OTROS INGRESOS"])

    # Desglose de otros ingresos (columnas bajo OTROS INGRESOS)
    bono_puesto = safe_float(row.get("Unnamed: 20"))      # BONO PUESTO
    bono_trabajo = safe_float(row.get("Unnamed: 21"))     # BONO POR TRABAJO
    asueto = safe_float(row.get("Unnamed: 22"))           # ASUETO
    horas_extras = safe_float(row.get("Unnamed: 23"))     # HORAS EXTRAS (monto)

    # Otros ingresos totales = total_ingresos - devengado (por si hay algo mas)
    otros_ingresos_total = total_ingresos - devengado

    datos = {
        # Datos generales
        "periodo": periodo_texto,
        "boleta_numero": int(row["Corr"]),
        "nombre": row["NOMBRE DEL EMPLEADO"],
        "puesto": row["NOMBRE DE LA INSTALACION"],
        "dias_trabajados": int(row["DIAS LABORADOS"]),
        "sueldo_base_mensual": sueldo_base_mensual,

        # Ingresos
        "devengado": devengado,
        "bono_puesto": bono_puesto,
        "bono_trabajo": bono_trabajo,
        "asueto": asueto,
        "horas_extras": horas_extras,
        "otros_ingresos_total": otros_ingresos_total,
        "total_ingresos": total_ingresos,

        # Descuentos / egresos
        "igss": safe_float(row.get("DESCUENTOS")),
        "seguro_vida": safe_float(row.get("Unnamed: 14")),
        "anticipo": safe_float(row.get("Unnamed: 15")),
        "prestamo": safe_float(row.get("Unnamed: 16")),
        "pgr": safe_float(row.get("Unnamed: 17")),
        "otros_descuentos": safe_float(row.get("Unnamed: 18")),
        "total_egresos": safe_float(row["TOTAL EGRESOS"]),

        # Liquido
        "liquido": safe_float(row["TOTAL A RECIBIR"]),
    }

    return datos



def formato_moneda(valor: float) -> str:
    """Devuelve el numero formateado como Q#.## con dos decimales."""
    return f"Q {valor:,.2f}"


def renderizar_comprobante(datos: dict) -> str:
    """
    Genera el HTML COMPLETO (pagina) del comprobante, con un boton
    para imprimir / guardar como PDF.
    """
    f = formato_moneda

    html = f"""<!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8" />
        <title>Comprobante de pago</title>
        <style>
        body {{
            font-family: Arial, sans-serif;
            background-color: #f5f5f5;
        }}
        .recibo {{
            width: 800px;
            margin: 20px auto;
            padding: 24px;
            border: 1px solid #333;
            background-color: #ffffff;
            font-size: 13px;
        }}
        .recibo h2, .recibo h3 {{
            text-align: center;
            margin: 4px 0;
        }}
        .recibo-table {{
            width: 100%;
            border-collapse: collapse;
            margin-top: 8px;
        }}
        .recibo-table td, .recibo-table th {{
            border: 1px solid #555;
            padding: 4px 6px;
        }}
        .section-title {{
            margin-top: 8px;
            font-weight: bold;
            text-decoration: underline;
        }}
        .right {{ text-align: right; }}
        .center {{ text-align: center; }}

        /* Boton que no se imprime */
        .no-print {{
            display: inline-block;
            margin: 10px auto;
            padding: 8px 16px;
            background-color: #222;
            color: #fff;
            border-radius: 4px;
            border: none;
            cursor: pointer;
            font-size: 13px;
        }}
        .no-print:hover {{
            background-color: #444;
        }}

        @media print {{
            .no-print {{
                display: none;
            }}
            body {{
                background-color: #ffffff;
            }}
        }}
        </style>
    </head>
    <body>

        <div style="text-align:center;">
            <button class="no-print" onclick="window.print()">
                Generar comprobante en PDF
            </button>
        </div>

        <div class="recibo">
            <h2>PROTECTOR S.A.</h2>
            <h3>Boleta de Liquidacion No. {datos["boleta_numero"]}</h3>
            <p><b>Periodo:</b> {datos["periodo"]}</p>

            <table class="recibo-table">
                <tr>
                    <th>Nombre completo</th>
                    <td colspan="3">{datos["nombre"]}</td>
                </tr>
                <tr>
                    <th>Puesto / Instalacion</th>
                    <td>{datos["puesto"]}</td>
                    <th>Dias trabajados</th>
                    <td class="center">{datos["dias_trabajados"]}</td>
                </tr>
                <tr>
                    <th>Sueldo base mensual</th>
                    <td class="right">{f(datos["sueldo_base_mensual"])}</td>
                    <th>Boleta por el</th>
                    <td class="center">100%</td>
                </tr>
            </table>

            <p class="section-title">INGRESOS</p>
            <table class="recibo-table">
                <tr>
                    <th>Concepto</th>
                    <th class="right">Monto</th>
                </tr>
                <tr>
                    <td>Devengado (segun dias trabajados)</td>
                    <td class="right">{f(datos["devengado"])}</td>
                </tr>
                <tr>
                    <td>Bono por puesto</td>
                    <td class="right">{f(datos["bono_puesto"])}</td>
                </tr>
                <tr>
                    <td>Bono por trabajo / productividad</td>
                    <td class="right">{f(datos["bono_trabajo"])}</td>
                </tr>
                <tr>
                    <td>Asueto</td>
                    <td class="right">{f(datos["asueto"])}</td>
                </tr>
                <tr>
                    <td>Horas extras</td>
                    <td class="right">{f(datos["horas_extras"])}</td>
                </tr>
                <tr>
                    <td><b>Otros ingresos (total)</b></td>
                    <td class="right"><b>{f(datos["otros_ingresos_total"])}</b></td>
                </tr>
                <tr>
                    <th>Total de ingresos</th>
                    <th class="right">{f(datos["total_ingresos"])}</th>
                </tr>
            </table>

            <p class="section-title">EGRESOS / DESCUENTOS</p>
            <table class="recibo-table">
                <tr>
                    <th>Concepto</th>
                    <th class="right">Monto</th>
                </tr>
                <tr><td>IGSS</td><td class="right">{f(datos["igss"])}</td></tr>
                <tr><td>Seguro de vida</td><td class="right">{f(datos["seguro_vida"])}</td></tr>
                <tr><td>Anticipos</td><td class="right">{f(datos["anticipo"])}</td></tr>
                <tr><td>Prestamos</td><td class="right">{f(datos["prestamo"])}</td></tr>
                <tr><td>P.G.R.</td><td class="right">{f(datos["pgr"])}</td></tr>
                <tr><td>Otros descuentos</td><td class="right">{f(datos["otros_descuentos"])}</td></tr>
                <tr>
                    <th>Total de egresos</th>
                    <th class="right">{f(datos["total_egresos"])}</th>
                </tr>
            </table>

            <p class="section-title">LIQUIDO A RECIBIR</p>
            <table class="recibo-table">
                <tr>
                    <th>Liquido a recibir</th>
                    <th class="right">{f(datos["liquido"])}</th>
                </tr>
            </table>

            <p style="margin-top: 24px;">
                He revisado y recibido de conformidad el importe neto indicado y acepto
                como buena esta liquidacion.
            </p>

            <br><br>
            <table class="recibo-table">
                <tr>
                    <td class="center">Firma de recibido</td>
                </tr>
                <tr>
                    <td style="height: 40px;"></td>
                </tr>
                <tr>
                    <td class="center">{datos["nombre"]}</td>
                </tr>
            </table>
        </div>

    </body>
    </html>
    """
    return html



# --------------------------------------------------------------------
#  App de Streamlit
# --------------------------------------------------------------------

st.title("Generador de comprobantes de pago")

st.write(
    "1. Sube el archivo de Excel de la planilla.\n"
    "2. Escribe el texto del periodo (ej. 'DEL 01 AL 15 DE NOVIEMBRE 2025').\n"
    "3. Selecciona el empleado y genera el comprobante."
)

archivo = st.file_uploader("Sube la planilla en Excel", type=["xlsx"])

if archivo is not None:
    try:
        df = cargar_planilla(archivo)
    except Exception as e:
        st.error(f"Error al leer el archivo: {e}")
    else:
        st.success(f"Planilla cargada. Empleados encontrados: {len(df)}")

        # Campo para que tu papa ponga el periodo como lo quiera ver en el recibo
        periodo = st.text_input(
            "Texto del periodo (aparece en el comprobante)",
            value="DEL 01 AL 15 DE NOVIEMBRE DE 2025"
        )

        nombres = df["NOMBRE DEL EMPLEADO"].tolist()
        nombre_sel = st.selectbox("Selecciona un empleado", nombres)

        if nombre_sel and st.button("Generar comprobante"):
            fila = df[df["NOMBRE DEL EMPLEADO"] == nombre_sel].iloc[0]
            datos = fila_a_diccionario(fila, periodo)
            html = renderizar_comprobante(datos)

            # Mostrar el HTML en un iframe con su propio boton de "Generar comprobante en PDF"
            components.html(html, height=900, scrolling=True)

            st.info(
                "Para guardarlo como PDF, usa Ctrl+P (o Imprimir) en el navegador "
                "y elige 'Guardar como PDF'."
            )

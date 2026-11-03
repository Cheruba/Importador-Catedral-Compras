import io
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Importador de Compras", page_icon="📊", layout="centered")

st.title("📊 Convertidor de Comprobantes AFIP a Excel")
st.write("Ingresá el CUIT de la empresa y subí el archivo CSV descargado de AFIP.")

# 1. Campo para ingresar el CUIT de la empresa
cuit_empresa = st.text_input("CUIT de la empresa (sin guiones o con guiones):", placeholder="Ej: 30719436281")

# 2. Componente para subir el archivo CSV
uploaded_file = st.file_uploader("Seleccioná el archivo CSV de AFIP", type=["csv"])

if uploaded_file is not None:
    if not cuit_empresa.strip():
        st.warning("⚠️ Recordá escribir el CUIT en la casilla superior para que el archivo descargado lleve su número al inicio.")

    try:
        # Cargar el CSV subido
        df_csv = pd.read_csv(uploaded_file, sep=';', encoding='utf-8')

        def a_numero(valor):
            if pd.isna(valor) or valor == '':
                return 0.0
            val_str = str(valor).replace('.', '').replace(',', '.').strip()
            try:
                return float(val_str)
            except ValueError:
                return 0.0

        # Construcción del DataFrame para la solapa Comprobantes
        df_excel = pd.DataFrame()

        # Conversión de Fecha a formato Fecha real de Excel
        fechas_convertidas = pd.to_datetime(df_csv['Fecha de Emisión'], errors='coerce')
        df_excel['Fecha'] = fechas_convertidas
        df_excel['FechaContable'] = fechas_convertidas

        df_excel['Proveedor'] = df_csv['Nro. Doc. Emisor']
        df_excel['Comprobante'] = df_csv['Tipo de Comprobante']
        df_excel['Sucursal'] = df_csv['Punto de Venta']
        df_excel['NroComprobante'] = df_csv['Número Desde']
        df_excel['CondCompra'] = 0
        df_excel['Moneda'] = None  # En blanco
        df_excel['Cotizacion'] = df_csv['Tipo Cambio'].apply(a_numero).replace(0, 1.0)
        df_excel['ClasificacionCF'] = 1
        df_excel['ImputacionCF'] = 1

        # Netos Gravados
        df_excel['NG al 21%'] = df_csv['Imp. Neto Gravado IVA 21%'].apply(a_numero)
        df_excel['NG al 27%'] = df_csv['Imp. Neto Gravado IVA 27%'].apply(a_numero)
        df_excel['NG al 10,5%'] = df_csv['Imp. Neto Gravado IVA 10,5%'].apply(a_numero)
        df_excel['NG al 5%'] = df_csv['Imp. Neto Gravado IVA 5%'].apply(a_numero)
        df_excel['NG al 2,5%'] = df_csv['Imp. Neto Gravado IVA 2,5%'].apply(a_numero)

        # IVA
        df_excel['IVA al 21%'] = df_csv['IVA 21%'].apply(a_numero)
        df_excel['IVA al 27%'] = df_csv['IVA 27%'].apply(a_numero)
        df_excel['IVA al 10,5%'] = df_csv['IVA 10,5%'].apply(a_numero)
        df_excel['IVA al 5%'] = df_csv['IVA 5%'].apply(a_numero)
        df_excel['IVA al 2,5%'] = df_csv['IVA 2,5%'].apply(a_numero)

        # CÁLCULO DE EXENTO: Si el Tipo de Comprobante es 11, 12 o 13 (Comprobantes C), el Total pasa a Exento
        def calcular_exento(row):
            tipo_comp = row['Tipo de Comprobante']
            imp_exentas = a_numero(row['Imp. Op. Exentas'])
            imp_total = a_numero(row['Imp. Total'])
            
            if tipo_comp in [11, 12, 13]:
                return imp_total
            return imp_exentas

        df_excel['Exento'] = df_csv.apply(calcular_exento, axis=1)
        df_excel['NoAlcanzado'] = df_csv['Imp. Neto No Gravado'].apply(a_numero) + df_csv['Imp. Neto Gravado IVA 0%'].apply(a_numero)
        df_excel['ImpuestosInternos'] = 0.0
        df_excel['PercepcionesIVA'] = 0.0
        df_excel['RegimenPercIVA'] = None
        df_excel['PercepcionesIB'] = df_csv['Otros Tributos'].apply(a_numero)
        df_excel['PercepcionesIG'] = 0.0
        df_excel['PercepcionesBP'] = 0.0

        # Totales y CAE
        df_excel['TotalComprobante'] = df_csv['Imp. Total'].apply(a_numero)
        df_excel['CAI/CAE/COE'] = df_csv['Cód. Autorización']
        df_excel['VtoCAI/CAE/COE'] = None

        # Solapa Proveedores
        df_proveedores = df_csv[['Nro. Doc. Emisor', 'Denominación Emisor']].drop_duplicates()
        df_proveedores.columns = ['Nro de Documento', 'Razon Social']

        # Mostrar métricas en pantalla
        col1, col2 = st.columns(2)
        col1.metric("Comprobantes procesados", len(df_excel))
        col2.metric("Monto Total", f"${df_excel['TotalComprobante'].sum():,.2f}")

        # Crear el archivo Excel en memoria
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl', date_format='YYYY-MM-DD', datetime_format='YYYY-MM-DD') as writer:
            df_excel.to_excel(writer, sheet_name='Comprobantes', index=False)
            df_proveedores.to_excel(writer, sheet_name='Proveedores', index=False)
        output.seek(0)

        # Construir el nombre del archivo final
        cuit_limpio = str(cuit_empresa).strip().replace('-', '')
        if cuit_limpio:
            nombre_archivo_salida = f"{cuit_limpio}-Importar_Compras.xlsx"
        else:
            nombre_archivo_salida = "Importar_Compras.xlsx"

        st.info(f"📁 El archivo se descargará con el nombre: **{nombre_archivo_salida}**")

        # Botón de descarga
        st.download_button(
            label="📥 Descargar Excel Listo",
            data=output,
            file_name=nombre_archivo_salida,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.success("¡Archivo listo para descargar!")

    except Exception as e:
        st.error(f"Error al procesar el archivo: {e}")

import io
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Importador de Compras", page_icon="📊", layout="centered")

st.title("📊 Convertidor de Comprobantes AFIP a Excel")
st.write("Ingresá el CUIT de la empresa y subí el archivo CSV descargado de AFIP.")

# Recuadro para ingresar el CUIT de la empresa (con tu CUIT configurado por defecto)
cuit_empresa = st.text_input("CUIT de la empresa:", value="30719436281")

# Componente para subir el archivo CSV
uploaded_file = st.file_uploader("Seleccioná el archivo CSV de AFIP", type=["csv"])

if uploaded_file is not None:
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
        df_excel['Fecha'] = df_csv['Fecha de Emisión']
        df_excel['FechaContable'] = df_csv['Fecha de Emisión']
        df_excel['Proveedor'] = df_csv['Nro. Doc. Emisor']
        df_excel['Comprobante'] = df_csv['Tipo de Comprobante']
        df_excel['Sucursal'] = df_csv['Punto de Venta']
        df_excel['NroComprobante'] = df_csv['Número Desde']
        df_excel['CondCompra'] = 0
        df_excel['Moneda'] = df_csv['Moneda']
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

        # No alcanzados / Exentos / Percepciones
        df_excel['Exento'] = df_csv['Imp. Op. Exentas'].apply(a_numero)
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

        # Mostrar métricas rápidas en pantalla
        col1, col2 = st.columns(2)
        col1.metric("Comprobantes procesados", len(df_excel))
        col2.metric("Monto Total", f"${df_excel['TotalComprobante'].sum():,.2f}")

        # Crear el archivo Excel en memoria
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_excel.to_excel(writer, sheet_name='Comprobantes', index=False)
            df_proveedores.to_excel(writer, sheet_name='Proveedores', index=False)
        output.seek(0)

        # Generar nombre dinámico del archivo según el CUIT ingresado
        cuit_limpio = str(cuit_empresa).strip().replace('-', '') if cuit_empresa else "Importar"
        nombre_archivo_salida = f"{cuit_limpio}-Importar_Compras.xlsx"

        # Botón para descargar el archivo procesado
        st.download_button(
            label="📥 Descargar Excel Listo",
            data=output,
            file_name=nombre_archivo_salida,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.success(f"¡Archivo procesado! El nombre generado es: **{nombre_archivo_salida}**")

    except Exception as e:
        st.error(f"Error al procesar el archivo: {e}")

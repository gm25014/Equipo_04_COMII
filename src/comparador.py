from collections import defaultdict
import pandas as pd
from openpyxl import load_workbook
from openpyxl.chart import BarChart, Reference
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer

# 1. Cargar el archivo de cotizaciones
df = pd.read_excel("data/raw/cotizaciones.xlsx")

# 2. Calcular los puntajes según los pesos (60% Precio, 25% Plazo de Entrega, 15% Condiciones de Pago)
max_precio = df['Precio'].max()
max_tiempo = df['PlazoEntrega'].max()
max_pago = df['CondicionesPago'].max()

df['Score_Precio'] = (1 - (df['Precio'] / max_precio)) * 60
df['Score_Tiempo'] = (1 - (df['PlazoEntrega'] / max_tiempo)) * 25
df['Score_Pago'] = (df['CondicionesPago'] / max_pago) * 15

# 3. Calcular el puntaje final total
df['Score_Total'] = df['Score_Precio'] + df['Score_Tiempo'] + df['Score_Pago']

# 4. Mostrar en la terminal el mejor proveedor por cada producto
print("¡Cotización procesada con éxito!\n")
for producto in df['Producto'].unique():
    sub_df = df[df['Producto'] == producto]
    ganador = sub_df.loc[sub_df['Score_Total'].idxmax()]
    print(f"Producto: {ganador['Producto']} --> Proveedor recomendado: {ganador['Proveedor']} (Puntaje: {ganador['Score_Total']:.2f})")

# 5. Guardar el resultado completo en la carpeta output
df.to_excel("output/resultado_cotizaciones.xlsx", index=False)
print("\n¡Archivo de resultados guardado con éxito en la carpeta output!")


def agregar_grafico_puntajes(ruta_archivo="output/resultado_cotizaciones.xlsx"):
    libro = load_workbook(ruta_archivo)
    hoja = libro.active

    encabezados = {
        str(celda.value).strip(): celda.column
        for celda in hoja[1]
        if celda.value is not None
    }
    if "Proveedor" not in encabezados or "Score_Total" not in encabezados:
        raise ValueError("No se encontraron las columnas 'Proveedor' y 'Score_Total'.")

    puntajes = defaultdict(lambda: [0.0, 0])
    for fila in range(2, hoja.max_row + 1):
        proveedor = hoja.cell(fila, encabezados["Proveedor"]).value
        puntaje = hoja.cell(fila, encabezados["Score_Total"]).value

        if proveedor is None or puntaje is None:
            continue

        nombre = str(proveedor).strip()
        if not nombre:
            continue
        if not isinstance(puntaje, (int, float)):
            raise ValueError(
                f"Puntaje no numérico para el proveedor '{nombre}' en la fila {fila}."
            )

        puntajes[nombre][0] += puntaje
        puntajes[nombre][1] += 1

    if not puntajes:
        raise ValueError("No hay puntajes por proveedor para graficar.")

    columna_proveedor = 11
    columna_puntaje = 12
    hoja.cell(1, columna_proveedor, "Proveedor")
    hoja.cell(1, columna_puntaje, "Puntaje promedio")

    for fila, (proveedor, (total, cantidad)) in enumerate(puntajes.items(), start=2):
        hoja.cell(fila, columna_proveedor, proveedor)
        hoja.cell(fila, columna_puntaje, total / cantidad)

    grafico = BarChart()
    grafico.type = "bar"
    grafico.style = 10
    grafico.title = "Puntaje promedio por proveedor"
    grafico.x_axis.title = "Puntaje promedio"
    grafico.y_axis.title = "Proveedor"
    grafico.y_axis.scaling.orientation = "maxMin"

    datos = Reference(
        hoja,
        min_col=columna_puntaje,
        min_row=1,
        max_row=len(puntajes) + 1,
    )
    categorias = Reference(
        hoja,
        min_col=columna_proveedor,
        min_row=2,
        max_row=len(puntajes) + 1,
    )
    grafico.add_data(datos, titles_from_data=True)
    grafico.set_categories(categorias)
    grafico.height = 8
    grafico.width = 16
    hoja.add_chart(grafico, "N2")

    libro.save(ruta_archivo)


# Ejecutar la función del gráfico
agregar_grafico_puntajes()


def generar_pdf_top_proveedores(df_resultados):
    """
    Genera un archivo PDF con los 5 proveedores con mayor puntaje.
    """
    # Nota: Tu columna principal de puntaje se llama 'Score_Total' en este script
    top_5 = df_resultados.sort_values(by="Score_Total", ascending=False).head(5)
    
    ruta_pdf = "output/top_5_proveedores.pdf"
    doc = SimpleDocTemplate(ruta_pdf, pagesize=letter)
    story = []
    styles = getSampleStyleSheet()
    
    story.append(Paragraph("<b>Reporte Oficial: Top 5 Proveedores (Taller 3)</b>", styles['Title']))
    story.append(Spacer(1, 15))
    
    for idx, row in top_5.iterrows():
        texto = f"<b>Producto:</b> {row['Producto']} <br/><b>Proveedor Recomendado:</b> {row['Proveedor']} <br/><b>Puntaje Total:</b> {row['Score_Total']:.2f}"
        story.append(Paragraph(texto, styles['Normal']))
        story.append(Spacer(1, 10))
        
    doc.build(story)
    print("¡Archivo PDF generado con éxito en la carpeta output/!")


# Ejecutar la función del PDF al final de todo
generar_pdf_top_proveedores(df)
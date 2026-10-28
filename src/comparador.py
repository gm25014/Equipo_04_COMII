from collections import defaultdict
from pathlib import Path
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from zipfile import BadZipFile
from xml.sax.saxutils import escape

import pandas as pd
from openpyxl import load_workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.utils.exceptions import InvalidFileException
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer


COLUMNAS_REQUERIDAS = {
    "Producto",
    "Proveedor",
    "Precio",
    "PlazoEntrega",
    "CondicionesPago",
}


def obtener_raiz_aplicacion():
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[1]


def obtener_archivo_inicial():
    archivo_externo = (
        obtener_raiz_aplicacion() / "data" / "raw" / "cotizaciones.xlsx"
    )
    if archivo_externo.is_file():
        return archivo_externo

    if getattr(sys, "frozen", False):
        carpeta_temporal = getattr(sys, "_MEIPASS", None)
        if carpeta_temporal:
            archivo_empaquetado = (
                Path(carpeta_temporal) / "data" / "raw" / "cotizaciones.xlsx"
            )
            if archivo_empaquetado.is_file():
                return archivo_empaquetado
    return archivo_externo


def calcular_puntajes(df):
    columnas_faltantes = COLUMNAS_REQUERIDAS.difference(df.columns)
    if columnas_faltantes:
        faltantes = ", ".join(sorted(columnas_faltantes))
        raise ValueError(f"El archivo no contiene las columnas requeridas: {faltantes}.")
    if df.empty:
        raise ValueError("El archivo de cotizaciones no contiene registros.")

    max_precio = df["Precio"].max()
    max_tiempo = df["PlazoEntrega"].max()
    max_pago = df["CondicionesPago"].max()
    if max_precio == 0 or max_tiempo == 0 or max_pago == 0:
        raise ValueError("Los valores máximos de precio, plazo y pago deben ser distintos de cero.")

    df = df.copy()
    df["Score_Precio"] = (1 - (df["Precio"] / max_precio)) * 60
    df["Score_Tiempo"] = (1 - (df["PlazoEntrega"] / max_tiempo)) * 25
    df["Score_Pago"] = (df["CondicionesPago"] / max_pago) * 15
    df["Score_Total"] = (
        df["Score_Precio"] + df["Score_Tiempo"] + df["Score_Pago"]
    )
    return df


def agregar_grafico_puntajes(ruta_archivo):
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


def generar_pdf_top_proveedores(df_resultados, ruta_pdf):
    top_5 = df_resultados.sort_values(
        by="Score_Total", ascending=False
    ).head(5)
    documento = SimpleDocTemplate(str(ruta_pdf), pagesize=letter)
    contenido = []
    estilos = getSampleStyleSheet()
    contenido.append(
        Paragraph(
            "<b>Reporte Oficial: Top 5 Proveedores</b>",
            estilos["Title"],
        )
    )
    contenido.append(Spacer(1, 15))

    for _, fila in top_5.iterrows():
        texto = (
            f"<b>Producto:</b> {escape(str(fila['Producto']))}"
            f"<br/><b>Proveedor recomendado:</b> {escape(str(fila['Proveedor']))}"
            f"<br/><b>Puntaje total:</b> {fila['Score_Total']:.2f}"
        )
        contenido.append(Paragraph(texto, estilos["Normal"]))
        contenido.append(Spacer(1, 10))

    documento.build(contenido)


def procesar_cotizaciones(ruta_entrada):
    ruta_salida = obtener_raiz_aplicacion() / "output"
    ruta_salida.mkdir(parents=True, exist_ok=True)
    df = calcular_puntajes(pd.read_excel(ruta_entrada))

    ruta_excel = ruta_salida / "resultado_cotizaciones.xlsx"
    ruta_pdf = ruta_salida / "top_5_proveedores.pdf"
    df.to_excel(ruta_excel, index=False)
    agregar_grafico_puntajes(ruta_excel)
    generar_pdf_top_proveedores(df, ruta_pdf)
    return ruta_excel, ruta_pdf, df


def iniciar_interfaz():
    raiz = tk.Tk()
    raiz.title("Comparador de cotizaciones")
    raiz.minsize(720, 480)

    contenedor = ttk.Frame(raiz, padding=20)
    contenedor.pack(fill="both", expand=True)
    contenedor.columnconfigure(0, weight=1)
    contenedor.rowconfigure(4, weight=1)

    ttk.Label(
        contenedor,
        text="Comparador de cotizaciones",
        font=("Segoe UI", 18, "bold"),
    ).grid(row=0, column=0, columnspan=2, sticky="w")
    ttk.Label(
        contenedor,
        text="Selecciona un archivo Excel para calcular y comparar las cotizaciones.",
    ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(4, 16))

    archivo_inicial = obtener_archivo_inicial()
    ruta_archivo = tk.StringVar(
        value=str(archivo_inicial) if archivo_inicial.is_file() else ""
    )
    ttk.Entry(contenedor, textvariable=ruta_archivo).grid(
        row=2,
        column=0,
        sticky="ew",
        padx=(0, 8),
    )

    def seleccionar_archivo():
        seleccion = filedialog.askopenfilename(
            title="Seleccionar archivo de cotizaciones",
            initialdir=str(archivo_inicial.parent),
            initialfile=archivo_inicial.name,
            filetypes=[("Archivos de Excel", "*.xlsx"), ("Todos los archivos", "*.*")],
        )
        if seleccion:
            ruta_archivo.set(seleccion)

    ttk.Button(
        contenedor,
        text="Examinar...",
        command=seleccionar_archivo,
    ).grid(row=2, column=1, sticky="ew")

    estado = tk.StringVar(value="Listo para procesar.")
    etiqueta_estado = ttk.Label(contenedor, textvariable=estado)
    etiqueta_estado.grid(row=3, column=0, columnspan=2, sticky="w", pady=12)

    tabla = ttk.Treeview(
        contenedor,
        columns=("producto", "proveedor", "puntaje"),
        show="headings",
        height=8,
    )
    tabla.heading("producto", text="Producto")
    tabla.heading("proveedor", text="Proveedor recomendado")
    tabla.heading("puntaje", text="Puntaje")
    tabla.column("producto", width=220)
    tabla.column("proveedor", width=260)
    tabla.column("puntaje", width=100, anchor="e")
    tabla.grid(row=4, column=0, columnspan=2, sticky="nsew")
    barra_desplazamiento = ttk.Scrollbar(
        contenedor,
        orient="vertical",
        command=tabla.yview,
    )
    tabla.configure(yscrollcommand=barra_desplazamiento.set)
    barra_desplazamiento.grid(row=4, column=2, sticky="ns")

    salidas = tk.StringVar(value="")
    ttk.Label(
        contenedor,
        textvariable=salidas,
        justify="left",
        wraplength=680,
    ).grid(row=5, column=0, columnspan=2, sticky="w", pady=(12, 4))

    def procesar():
        archivo = Path(ruta_archivo.get().strip())
        if not archivo.is_file():
            messagebox.showerror(
                "Archivo no válido",
                "Selecciona un archivo de cotizaciones existente en formato .xlsx.",
                parent=raiz,
            )
            return
        if archivo.suffix.lower() != ".xlsx":
            messagebox.showerror(
                "Formato no compatible",
                "El archivo debe estar en formato .xlsx.",
                parent=raiz,
            )
            return

        for elemento in tabla.get_children():
            tabla.delete(elemento)
        salidas.set("")
        boton_procesar.configure(state="disabled")
        estado.set("Procesando cotizaciones...")
        raiz.update_idletasks()
        try:
            ruta_excel, ruta_pdf, resultados = procesar_cotizaciones(archivo)
        except (
            BadZipFile,
            ImportError,
            InvalidFileException,
            OSError,
            TypeError,
            ValueError,
        ) as error:
            estado.set("No se pudo completar el procesamiento.")
            messagebox.showerror(
                "Error al procesar",
                str(error),
                parent=raiz,
            )
        else:
            for elemento in tabla.get_children():
                tabla.delete(elemento)
            mejores = resultados.nlargest(5, "Score_Total")
            for _, fila in mejores.iterrows():
                tabla.insert(
                    "",
                    "end",
                    values=(
                        fila["Producto"],
                        fila["Proveedor"],
                        f"{fila['Score_Total']:.2f}",
                    ),
                )
            estado.set("Proceso completado correctamente.")
            salidas.set(
                f"Excel: {ruta_excel}\nPDF: {ruta_pdf}"
            )
        finally:
            boton_procesar.configure(state="normal")

    boton_procesar = ttk.Button(
        contenedor,
        text="Procesar cotizaciones",
        command=procesar,
    )
    boton_procesar.grid(row=6, column=0, sticky="w", pady=(12, 0))

    raiz.mainloop()


if __name__ == "__main__":
    iniciar_interfaz()
import pandas as pd

# 1. Cargar el archivo de cotizaciones
df = pd.read_excel("../data/raw/cotizaciones.xlsx")

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
df.to_excel("../output/resultado_cotizaciones.xlsx", index=False)
print("\n¡Archivo de resultados guardado con éxito en la carpeta output!")
# Taller 3: Evaluación y Automatización de Cotizaciones de Proveedores

## Descripción del Proyecto
Este proyecto automatiza la evaluación y selección de cotizaciones de proveedores utilizando Python y la librería pandas. Aplica un modelo de ponderación basado en tres criterios clave:
- *Precio*: 60% (ranking inverso: menor precio obtiene mayor puntaje).
- *Plazo de Entrega*: 25% (ranking inverso: menor tiempo de entrega obtiene mayor puntaje).
- *Condiciones de Pago*: 15% (ranking directo: mejores condiciones obtienen mayor puntaje).

## Estructura del Repositorio
- data/raw/: Contiene el archivo Excel original con las cotizaciones de entrada.
- output/: Contiene el archivo Excel resultante con los puntajes calculados y el mejor proveedor seleccionado.
- src/: Contiene el script principal de Python (comparador.py).
- docs/: Documentación del proyecto (README.md y bitacora_prompts.md).
- requirements.txt: Lista de dependencias y librerías necesarias.

## Requisitos e Instalación
1. Asegúrate de tener Python instalado.
2. Instala la librería requerida ejecutando en tu terminal:
   ```bash
   pip install -r requirements.txt
 
## Nuevas Funcionalidades (Taller 3)
1. *Gráfico en Excel:* El script genera y añade automáticamente un gráfico de barras horizontales en el archivo de resultados (output/resultado_cotizaciones.xlsx) que muestra el puntaje promedio por proveedor.
2. *Reporte en PDF:* Se implementó la función generar_pdf_top_proveedores utilizando la librería reportlab para exportar de manera automática un reporte oficial en PDF (output/top_5_proveedores.pdf) con los 5 mejores proveedores clasificados por su Score_Total.

## Ejecución del Script
Para procesar las cotizaciones, generar los gráficos y exportar el reporte en PDF, ejecuta el siguiente comando en tu terminal:
```bash
python src/comparador.py

### Pasos finales para guardarlo:
1. Pega esto al final de tu archivo *README.md*.
2. Presiona *Ctrl + S* para guardarlo.
3. Vuelve a tu terminal por última vez para subirlo a GitHub con estos comandos:
   ```bash
   git add docs/README.md
   git commit -m "Actualizado el README con las caracteristicas de graficos y PDF"
   git push
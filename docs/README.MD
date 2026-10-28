# Taller 3: Evaluación y Automatización de Cotizaciones de Proveedores

## Descripción del Proyecto
Este proyecto automatiza la evaluación y selección de cotizaciones de proveedores utilizando Python y la librería pandas. Aplica un modelo de ponderación basado en tres criterios clave:
- Precio: 60% (ranking inverso: menor precio obtiene mayor puntaje).
- Plazo de Entrega: 25% (ranking inverso: menor tiempo de entrega obtiene mayor puntaje).
- Condiciones de Pago: 15% (ranking directo: mejores condiciones obtienen mayor puntaje).

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
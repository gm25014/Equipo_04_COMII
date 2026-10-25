# Bitácora de Prompts - Taller 3

## 1. Estructura y Creación del Proyecto
- *Prompt:* Ayúdame a estructurar las carpetas y archivos necesarios para el Taller 3 de automatización de cotizaciones en Python usando pandas.
- *Resultado:* Definición de la arquitectura del repositorio (src/, data/raw/, output/, docs/).

## 2. Lógica del Comparador de Cotizaciones
- *Prompt:* Diseña un script en Python que lea el archivo de cotizaciones con pandas, calcule los puntajes ponderados (60% precio, 25% plazo, 15% pago) usando ranking inverso para precio y tiempo, y guarde el mejor proveedor.
- *Resultado:* Creación del código principal en src/comparador.py.

## 3. Corrección de Rutas y Verificación
- *Prompt:* Tengo un FileNotFoundError al ejecutar el script desde la raíz, ayúdame a ajustar las rutas relativas para que lea correctamente desde data/raw/cotizaciones.xlsx.
- *Resultado:* Ajuste de las rutas relativas en comparador.py y ejecución exitosa del script con salida en output/.
# Taller 3: Evaluación y Automatización de Cotizaciones de Proveedores

## Descripción del Proyecto
Este proyecto ofrece una interfaz gráfica para evaluar y comparar cotizaciones de proveedores utilizando Python y la librería pandas. Aplica un modelo de ponderación basado en tres criterios clave:
- Precio: 60% (ranking inverso: menor precio obtiene mayor puntaje).
- Plazo de Entrega: 25% (ranking inverso: menor tiempo de entrega obtiene mayor puntaje).
- Condiciones de Pago: 15% (ranking directo: mejores condiciones obtienen mayor puntaje).

## Estructura del Repositorio
- data/raw/: Contiene el archivo Excel original con las cotizaciones de entrada.
- output/: Contiene el archivo Excel resultante con los puntajes y el reporte PDF de los cinco mejores proveedores.
- src/: Contiene la interfaz gráfica y la lógica principal en comparador.py.
- docs/: Documentación del proyecto (README.md y bitacora_prompts.md).
- requirements.txt: Lista de dependencias y librerías necesarias.

## Requisitos e Instalación
Se requiere Python 3 instalado en Windows y disponible mediante el lanzador
`py` o el comando `python`. `iniciar.bat` reutiliza las dependencias instaladas
en Python; si no están disponibles, crea el entorno virtual local `.venv` e
instala lo indicado en `requirements.txt`. En ese caso, la primera ejecución
necesita conexión a Internet para descargar las librerías.

## Ejecución de la interfaz gráfica en Windows
1. Abre la carpeta del proyecto.
2. Ejecuta `fluji.exe` o `iniciar.bat` desde el Explorador de archivos. El iniciador abre el ejecutable empaquetado cuando está disponible; si no, usa Python.
3. La interfaz abre con `data/raw/cotizaciones.xlsx` seleccionado si el archivo existe. Usa **Examinar...** para elegir otro archivo Excel `.xlsx`.
4. Pulsa **Procesar cotizaciones**. La tabla muestra los cinco mejores resultados y la interfaz indica dónde se guardaron el Excel y el reporte PDF en `output/`.

### Crear o actualizar `fluji.exe`

Ejecuta `empaquetar.bat` desde la carpeta raíz del proyecto. El proceso prepara `.venv` si hace falta, instala los requisitos de compilación y crea `fluji.exe` en esa misma carpeta. El ejecutable incluye una copia de ejemplo del Excel; mantiene la estructura del proyecto y utiliza los archivos de `data/raw/` y `output/` junto al ejecutable.

Para ejecutar sin empaquetar, usa `iniciar.bat` sin `fluji.exe` o ejecuta desde la raíz:

```powershell
.\.venv\Scripts\python.exe .\src\comparador.py
```

La interfaz usa Tkinter, incluido en las instalaciones habituales de Python para Windows; no requiere una dependencia adicional en `requirements.txt`. La primera compilación requiere conexión a Internet para instalar PyInstaller.

Si se creó el entorno virtual, para usarlo desde una terminal activa
`.venv\Scripts\activate` antes de ejecutar otros comandos Python.

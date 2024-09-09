# IMAGENX_ExcelFunction
Complemento de Excel para insertar imágenes en celdas a partir de una URL.

## Descripción
Este complemento de Excel permite insertar imágenes en una celda específica a partir de una URL, ajustando el tamaño de la imagen para que se adapte a la celda.

## Idiomas de Función
IMAGEX (Inglés) / IMAGENX (Español).

## Características Principales
- **Inserción de Imágenes desde URL**: Puedes insertar imágenes directamente desde una URL o desde una celda que contenga la URL.
- **Ajuste Automático**: Las imágenes se ajustan automáticamente al tamaño de la celda de destino, manteniendo la relación de aspecto.
- **Fácil Integración**: Con una sintaxis clara y sencilla, IMAGENX se integra perfectamente en tus hojas de cálculo, permitiéndote manejar imágenes con facilidad.

## Descargar el Complemento
Puedes descargar el complemento IMAGEX.xlam desde el siguiente enlace:
[Descargar IMAGENX.xlam]([https://github.com/danraxai/IMAGENX_ExcelFunction/blob/main/IMAGENX.xlam](https://github.com/danraxai/IMAGEX_ExcelFunction/blob/main/IMAGEX.xlam))

## Instalación
1. Descargar el Archivo:
   - Descarga el archivo .xlam desde el enlace proporcionado.
2. Abrir Excel:
   - Abre Excel y ve a **Archivo > Opciones > Complementos**.
3. Administrar Complementos:
   - En la parte inferior de la ventana, selecciona **Complementos de Excel** en el menú desplegable y haz clic en **Ir...**.
4. Agregar el Complemento:
   - Haz clic en **Examinar...** y navega hasta el archivo .xlam que descargaste.
   - Selecciona el archivo y haz clic en **Aceptar**.
5. Activar el Complemento:
   - Asegúrate de que la casilla junto a IMAGENX esté marcada y haz clic en **Aceptar**.

## Sintaxis

=IMAGENX(celdaOUrl, celdaDestino)

- **celdaOUrl**: Una celda que contiene la URL de la imagen o la URL directamente como texto.
- **celdaDestino**: La celda donde se insertará la imagen y se ajustará automáticamente al tamaño de la celda.

## Ejemplos de Uso
### Ejemplo 1: Insertar una Imagen desde una URL en una Celda
Supongamos que tienes la URL de una imagen "https://ejemplo.com/imagen.jpg" y quieres insertarla en la celda A1. Puedes usar la función IMAGENX de la siguiente manera:

=IMAGENX("https://ejemplo.com/imagen.jpg", A1)

Esto insertará la imagen en la celda A1, ajustando su tamaño automáticamente para que se adapte a la celda.

### Ejemplo 2: Usar una URL desde una Celda
Si la URL está almacenada en la celda B1, puedes insertar la imagen en la celda C1 con la siguiente fórmula:

=IMAGENX(B1, C1)

Esto insertará la imagen de la URL en la celda C1.

## Beneficios
- **Automatización**: Evita insertar manualmente imágenes una por una. IMAGENX permite automatizar el proceso directamente desde tus fórmulas de Excel.
- **Ajuste de Tamaño**: No necesitas preocuparte por redimensionar las imágenes, ya que IMAGENX ajusta automáticamente la imagen para que encaje en la celda de destino.
- **Fácil de Usar**: Solo necesitas una URL y una celda de destino. IMAGENX se encarga del resto.

## Conclusión
La función IMAGENX es una herramienta poderosa para insertar y gestionar imágenes directamente en Excel. Ya sea que estés creando reportes visuales, insertando logotipos o utilizando imágenes en tus análisis, IMAGENX te proporciona una manera rápida y eficiente de hacerlo. Mejora tu flujo de trabajo y añade un toque visual a tus hojas de cálculo de manera simple y automatizada.

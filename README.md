# Crazy Form Bot

Este script en Python automatiza el reto **Crazy Form** de Arena RPA. Lee un archivo Excel con los datos de 10 usuarios y llena el formulario dinámico en la página web de forma automática y resiliente a cambios en los campos.

# ¿Qué hace el script?

- Lee el archivo `datos.xlsx` (10 filas de datos).  
- Abre la página del reto y hace clic en "Iniciar Reto".  
- Llena y envía el formulario para cada fila del Excel.  
- Espera 30 segundos al finalizar antes de cerrar el navegador.

---

## Librerías necesarias

```bash
pip install pandas selenium openpyxl



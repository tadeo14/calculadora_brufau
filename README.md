# Calculadora Brufau

Aplicación Streamlit para herramientas de compras de Brufau Sanitarios.
Incluye: calculadora de porcentajes, descuentos escalonados y listas dolarizadas con importación/exportación de Excel.

## Requisitos

- Python 3.11+ instalado
- Conexión a internet para instalar dependencias
- Windows PowerShell o CMD

## Instalación

1. Abre una terminal en el directorio del proyecto:

   ```powershell
   cd "c:\Users\sanch\OneDrive\Documents\PROYECTOS BRUFAU\calculadora_brufau"
   ```

2. Crea y activa un entorno virtual recomendado:

   PowerShell:
   ```powershell
   python -m venv .venv
   .\.venv\Scripts\Activate.ps1
   ```

   CMD:
   ```cmd
   python -m venv .venv
   .venv\Scripts\activate
   ```

3. Instala las dependencias:

   ```powershell
   pip install -r requirements.txt
   ```

## Ejecución

1. Ejecuta la aplicación con Streamlit:

   ```powershell
   streamlit run app.py
   ```

2. Abre el navegador en la URL que Streamlit muestre, normalmente:

   ```text
   http://localhost:8501
   ```

## Archivos principales

- `app.py`: aplicación principal de Streamlit.
- `db.py`: inicialización y acceso a la base de datos SQLite local.
- `io_xlsx.py`: funciones para parsear Excel y exportar listas dolarizadas.
- `requirements.txt`: dependencias del proyecto.

## Notas

- La base de datos SQLite se crea automáticamente en `brufau.db` en el mismo directorio.
- Si no tienes Streamlit, se instala junto a las demás dependencias desde `requirements.txt`.
- Al cerrar la terminal, el entorno virtual puede desactivarse con:

  ```powershell
  deactivate
  ```

# DescuentosStock

Sistema de escritorio para actualizar descuentos en una plantilla de Excel a partir de pedidos mayoristas, canjes y la opción de Mercado Libre desde un menú interactivo.

## ¿Qué hace el sistema? 🚀

El programa permite:

- leer archivos de pedidos desde carpetas específicas;
- actualizar la columna de descuentos de una plantilla llamada Plantilla.xlsx;
- mover los archivos ya procesados a la carpeta Descontados por fecha;
- generar respaldos en la carpeta Descuentos;
- registrar en una pestaña llamada No Encontrados los productos que no existen en la plantilla.

## Requisitos 📋

- Python 3.10 a 3.15
- Poetry
- Microsoft Excel para abrir y editar la plantilla
- Conexión a la red o acceso a los archivos que se van a procesar

## Dependencias principales

El proyecto utiliza:

- pandas
- openpyxl
- pdfplumber
- colorama
- xlrd
- xlutils

## Instalación local

1. Clonar el repositorio:

```bash
git clone https://github.com/Kniball4726/DescuentosStock.git
```

2. Ingresar a la carpeta del proyecto:

```bash
cd DescuentosStock
```

3. Instalar dependencias:

```bash
poetry install
```

4. Colocar la plantilla base en la raíz del proyecto con el nombre Plantilla.xlsx.

5. Ejecutar la aplicación:

```bash
poetry run python app.py
```

## Estructura esperada del sistema

Al ejecutar la aplicación por primera vez, se crean automáticamente las carpetas:

- Canjes
- Descuentos
- Descontados
- Mayoristas
- Mercado Libre

Además, el proceso espera los siguientes archivos:

- Mayoristas: archivos PDF
- Canjes: archivos Excel (.xlsx)
- Mercado Libre: archivos compatibles con la opción del menú, aunque su procesamiento no está activo en la versión actual
- Plantilla.xlsx: archivo base donde se registran los descuentos

## Flujo de uso

1. Colocar los archivos a procesar en las carpetas correspondientes.
2. Ejecutar la aplicación.
3. Elegir una opción del menú:
   - 1. Descontar Mayoristas: procesa los PDFs y actualiza los descuentos en la plantilla.
   - 2. Descontar Canjes: procesa los archivos Excel de canjes y suma las cantidades sobre la plantilla.
   - 3. Descontar Mercado libre: opción disponible en el menú, pero su procesamiento no está implementado en esta versión.
   - 4. Guardar Descuentos: crea un backup en la carpeta Descuentos y limpia la columna de descuentos para el siguiente ciclo.
   - 5. Salir.
4. Revisar la pestaña No Encontrados si algunos códigos no existen en la plantilla.

## Generación de ejecutable

Para crear un ejecutable de producción:

```bash
poetry run pyinstaller --onefile app.py
```

El ejecutable se generará en la carpeta dist.

## Autor

Gregory Rodriguez

- [@kniball4726](https://github.com/kniball4726)

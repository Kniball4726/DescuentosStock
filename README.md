# Sistema de descuento Mayoristas, Canjes y Mercado libre
Automatización para el descuento de productos vendidos en Mayoristas, mercado libre y canjes)

## Comenzando 🚀

Estas instrucciones te permitirán obtener una copia del proyecto en funcionamiento en tu máquina local para propósitos de desarrollo y pruebas.

Prerrequisitos 📋
  
  - Instalación de Python
  - Instalación de Visual Studio Code (VSC)
  - Instalaciòn de gestor de paquetes (Poetry)
  - Clonación de repositorio
    
1.- Instalar python desde la pagina oficial:

[Descarga de python](https://www.python.org/)

2.- Verificar instalacion de python

````bash
  python --version
````

3.- Verificar la instalación del gestor de paquetes PIP

```bash
  pip --version
```

4.- Instalación de editor de codigo Visual Studio Code

[Descarga de VSC](https://code.visualstudio.com/)

5.- Instalación de Gestor de paquetes Poetry

  - Desde una terminal powershell dentro de VSC ejecuta
  
```bash
  (Invoke-WebRequest -Uri https://install.python-poetry.org -UseBasicParsing).Content | py
```

  - Reinicia la aplicacón y ejecuta


```bash
  poetry --version
```

## Instalación local

Abrir visual estudio code y abrir una terminal con Ctrl+Shift+ñ

  - Hacer un clon del repositorio:

```bash[
  git clone https://github.com/Kniball4726/DescuentosStock.git
```

  - Desde el terminal entrar en la carpeta DescuentosStock

```bash
  cd DescuentosStock
```

  - Ejecutar el siguiente comando
  - Este comando instalara todas las dependencias del proyecto utilizadas en la programación

```bash
  poetry install
```

  - Crear ejecutable para uso en producción

```bash
  poetry run pyinstaller --onefile app.py
```

  - Se genera una carpeta llamada "dist" dentro del proyecto, esta carpeta contiene un archivo llamada "app.exe", la cual va a se el ejecutable final para que use el usuario.

## Forma de uso (Usuario final)

  - Se debe colocar el archivo Plantilla.xlsx donde sea que coloque el ejecutable para que tenga un funcionamiento correcto.
  
  - Al ejecutar el Script por primera vez se van a crear cuatro carpetas junto al ejecutable las cuales tendran por nombre "Canjes", "Descontados", "Mayoristas" y "Mercado Libre".
  
  - Al crear estas carpetas se debe colocar el archivo a descontar con formato .pdf dentro de la carpeta "Mayoristas", Archivos de canjes con formato .xls o xlsx dentro de la carpeta "Canjes" y archivos .doc o .docx en la carpeta "Mercado Libre"
  
  - El programa va a desplegar un menú interactivo con cinco opciones:
  
      - Opción 1: Hace el descuento de todos los archivos .pdf de pedidos mayoristas cargados en la carpeta "Mayoristas" generando la carga de los productos dentro del archivo Plantilla.xlsx
      - Opción 2: Hace el descuento de todos los archivos .xls o .xlsx correspondiente a los canjes cargados en la carpeta "Canjes" sumando los productos ya descontados de mayoristas en el archivo Plantilla.xlsx
      - Opción 3: Hace el descuento de todos los archivos .doc o .docx de los pedidos de Mercado Libre cargados en la carpeta "Mercado Libre" sumando los productos ya descontados de mayoristas y canjes en el archivo Plantilla.xlsx
      - Opción 4: Guardar descuento, es de vital importancia una vez terminado los descuentos hacer este paso para cerrar el proceso completo de manera eficiente, al ejecutar esta funcion, el programa hace un respaldo del descuento consolidado del dia dentro de la carpeta "Descuentos" con la fecha actual y formato .xlsx
      - Opción 5: Salir del programa y finalizar, se cierra la consola interactiva.
   

## Autor

  Gregory Rodriguez - Trabajo inicial, Desarrollo y documentación

- [@kniball4726](https://github.com/kniball4726)

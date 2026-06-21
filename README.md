Sistema de descuento Mayoristas, Canjes y Mercado libre
Automatización para el descuento de productos vendidos en Mayoristas, mercado libre y canjes)

Comenzando 🚀
Estas instrucciones te permitirán obtener una copia del proyecto en funcionamiento en tu máquina local para propósitos de desarrollo y pruebas.

Prerrequisitos 📋

1.- Instalar python desde la pagina oficial:

Descarga de python
2.- Verificar instalacion de python

  python --version
3.- Verificar la instalación del gestor de paquetes PIP

  pip --version
4.- Instalación de editor de codigo Visual Studio Code

Descarga de VSC
Instalación local
Abrir visual estudio code y abrir una terminal con Ctrl+Shift+ñ

hacer un clon del repositorio

  git clone https://github.com/kniball4726/Python.git
Desde el terminal entrar en la carpeta Python

  cd python
Estando dentro del proyecto desde terminal se debe crear un entorno virtual

  python -m venv .venv
Para activar el entorno virtual se debe entrar en la carpeta .venv/Scrips y correr Activate de la siguiente manera

  cd .\.venv\Scripts\
  .\activate
Volvemos a la carpeta raiz de nuestro proyecto

  cd ../..
Se deben instalar las dependencias utilizadas en el proyecto para que funciona de manera optima

  pip install -r requirements.txt
para hacer un despliegue con ejecutable desde windows se debe usar:

    pyinstaller --onefile app.py
al correr este comando se crearan dos carpetas denominadas build y dist

dentro de la carpeta dist encontraremos un archivo llamado app.exe

este ejecutable de windows puede utilizarse directamente desde el escritorio de su pc otorgandole un acceso directo, se puede modificar el nombre y el icono

Forma de uso
Al ejecutar el Script por primera vez se van a crear dos carpetas junto al ejecutable las cuales tendran por nombre "Descontar" y "Descontados".
Al crear estas carpetas se debe colocar el archivo a descontar con formato .docx dentro de la carpeta "Descontar" y se debe correr el script nuevamente.
Se debe correr el Script una vez colocado el archivo y este va a analizar este archivo y va a crear una carpeta con el nombre de la fecha actual Ejemplo: "2026-05.02" y dentro de esta carpeta se va a crear un archivo llamado "Consolidado_(fecha actual).xlsx" siguiendo con el ejemplo "Consolidado_2026-05-02.xlsx".
Una vez el Script realice el descuento informa a traves de consola y a su vez el archivo.docx pasa a la carpeta "Descontados".
Si abrimos el archivo de excel "Consolidado_2026-05-02.xlsx" estara el descuento correspondiente de los productos señalados en el .docx indicando cantidad y productos.
Cabe destacar que el descuento al hacerse una vez por fecha, no se podra hacer de nuevo ya que se generaria un conflicto de fechas, en caso de tener que hacer otro descuento se debera renombrar la carpeta generada con la fecha para poder asi hacer otro descuento del mismo día.
Documentación
En la carpeta raiz se encuentra un archivo llamado documentacion.py al ejecutar este archivo por terminal se observa la documentación completa de todas las funciones utilizadas en la aplicación

Autor
Gregory Rodriguez - Trabajo inicial, Desarrollo y documentación

@kniball4726

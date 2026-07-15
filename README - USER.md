# Sistema de descuento Mayoristas, Canjes y Mercado libre

Automatización para el descuento de productos vendidos en Mayoristas, mercado libre y canjes

## Comenzando 🚀

Estas instrucciones te permitirán hacer el descuento de productos de pedidos maoristas y canjes

## Prerrequisitos 📋

1.- Descargar pedidos a preparar desde Flexxus:

- Se conecta a la red a traves de la puerta de enlace procom en la barra de estado de la pc, se ubica una figura de mundo, boton derecho del mouse y conectar.

- Una vez conectado se ingresa al flexxus con usuario y clave asignada a traves del acceso directo "WebSupplements12.rdp"

- En la pestaña Ventas, presionar el boton Pedido, seleccionar "Plantilla de nota de Pedidos"

- Seleccionar fecha, preferiblemente con un mes de antelación.

- Van a aparecer los pedidos a preparar

- Se le da doble click en pedido seleccionado y sale un formulario que informa los productos del pedido.

- Se presiona el boton imprimir y sale una pantalla de interface

- Se le da click a la impresoray se seleccionar "Microsoft Print to PDF(Redirected xxx)"

- Boton aceptar y se procede a guardar el remito en la carpeta correspondiente la cual se encuentra en el "Escritorio/Descuentos/Mayoristas"

2.- Realizar el descuento:

- En el escritorio ejecutamos el acceso directo con el nombre "Dewcontar"

- Al iniciar el programa se despliega un menú con:

    1.- Descontar Mayoristas
    2.- Descontar canjes
    3.- Descontar Mercado Libre
    4.- Guardar Descuentos (Backup + Limpiar Plantilla)
    5.- Salir

## Forma de uso (Usuario final)

- Se debe colocar el archivo Plantilla.xlsx donde sea que coloque el ejecutable para que tenga un funcionamiento correcto.

- Al ejecutar el Script por primera vez se van a crear cuatro carpetas junto al ejecutable las cuales tendrán por nombre "Canjes", "Descuentos", "Mayoristas" y "Mercado Libre".

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

# DseFacturasToCsv v1.0
## Programa para transformar facturas desde un excel y grabarlas en un csv para su importacion

### Desarrollado por Carlos Clemente (06-2024)

### Control de versiones
 - Version 1.0 - Primera version funcional
<br>

**Instrucciones:**
 - Partiendo de un listado de facturas en Excel, se hace una
   transformacion de los datos para generar un fichero .csv que pueda importarse en la contagen
 - En la primera version se ha desarrollado para un cliente especifico, pero admite multiples formatos generando nuevas clases
   con el proceso de transformacion necesario
 - Se pueden establecer la hoja y la celda en la que se insertarán los datos; si no se indican se insertaran en la hoja 1 y celda A1
 - En la ejecucion se debe pasar como parametro el tipo de conversion, asi como el fichero de entrada y el de salida
 - Si no se pasa el fichero de salida, se generará uno con el mismo nombre del de la entrada en formato .csv
 - El fichero de salida sera un .csv separado por punto y coma
<br>

**Uso:** 
* dsefacturastocsv tipo entrada salida
<br>

**Parametros de ejecucion:** 
```
	tipo		Identifica la transformacion a realizar. Esta formado por una letra y dos numeros ('E'mitidas, 'R'ecibidas)
	entrada		Fichero excel de entrada
	salida		Fichero csv donde grabar los datos (opcional)
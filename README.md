# DseImfacex v1.3
## Programa para importar un listado de facturas desde un excel personalizando las columnas de datos

### Desarrollado por Carlos Clemente (06-2024)

### Control de versiones
 - Version 1.0	- Primera version funcional.
 - Version 1.1	- Añadida funcionalidad para agrupar por fechas (tipo E01).
 - Version 1.2	- Añadidos campos en facturas emitidas para recargo de IVA y retencion. 
				- Revision procesoAlcasal para ajustes de varios campos.
				- Añadida funcionalidad para obtener fichero de errores en el proceso.
 - Version 1.3	- Añadida funcionalidad para importar facturas emitidas desde un Excel indicando en que columna esta cada campo
				- Modificado el pase de parametros en la ejecucion
<br>

**Instrucciones:**
 - Partiendo de un listado de facturas en Excel, se hace una transformacion de los datos para generar
   un fichero .csv que pueda importarse en la contagen
 - En la primera version se ha desarrollado para un cliente especifico (proceso E01), pero admite multiples formatos
   generando nuevas clases con el proceso de transformacion necesario
 - Si no se pasa el fichero de salida, se generará uno con el mismo nombre del de la entrada en formato .csv
 - El fichero de salida sera un .csv separado por punto y coma
 - El fichero de configuracion debe tener en cada linea la columna y nombre del campo separado por punto y coma
 - El fichero Excel debe tener una fila con una cabecera, que es la que se indica en el parameto 'fila'
 - Se puede indicar de forma opcional el numero de hoja en la que estan los datos (por defecto la 1)
<br>

**Uso:** 
* dseimfacex entrada=ficheroEntrada.xlsx salida=ficheroSalida.csv proceso=tipoProceso configuracion=ficheroConfiguracion.txt fila=1 hoja=1
<br>

**Parametros de ejecucion:** 
```
ficheroEntrada	Fichero excel de entrada
ficheroSalida	Fichero csv en el que se grabara el resultado (opcional)
proceso			Identifica el tipo de proceso a realizar. Esta formado por una letra y dos numeros ('E'mitidas, 'R'ecibidas)
				Para la importacion 'estandar' de Diagram sera 'E00' para emitidas y 'R00' para recibidas
configuracion	Fichero de configuracion donde se indica en que columna esta cada uno de los campos a importar
fila			Fila que contiene la cabecera de las columnas
hoja			(Opcional) Hoja en la que estan los datos, si no se pasa se toma por defecto la 1
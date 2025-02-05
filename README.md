# DseImfacex v1.5.1.0
## Programa para importar un listado de facturas desde un excel personalizando las columnas de datos

### Desarrollado por Carlos Clemente (01-2025)

### Control de versiones
 - Version 1.0	- Primera version funcional.
 - Version 1.1	- Añadida funcionalidad para agrupar por fechas (tipo E01).
 - Version 1.2	- Añadidos campos en facturas emitidas para recargo de IVA y retencion. 
				- Revision procesoAlcasal para ajustes de varios campos.
				- Añadida funcionalidad para obtener fichero de errores en el proceso.
 - Version 1.3	- Añadida funcionalidad para importar facturas emitidas desde un Excel indicando en que columna esta cada campo
				- Modificado el pase de parametros en la ejecucion
 - Version 1.4	- Modificado para pasar un guion con los parametros y configuracion de columnas
 - Version 1.5	- Incluida biblioteca 'UtilidadesDiagram' como estatica
				- Añadida propiedad 'ficheroFactura' a la clase 'Facturas'
<br>

**Instrucciones:**
 - Partiendo de un listado de facturas en Excel, se hace una transformacion de los datos para generar
   un fichero .csv que pueda importarse en la contagen
 - En la primera version se ha desarrollado para un cliente especifico (proceso E01), pero admite multiples formatos
   generando nuevas clases con el proceso de transformacion necesario
 - Si no se pasa el fichero de salida, se generará uno con el mismo nombre del de la entrada en formato .csv
 - El fichero de salida sera un .csv separado por punto y coma
 - En el guion se pasan dos zonas con [parametros] y [columnas] seguidas cada una con los valores correspondientes
 - Cada parametro debe pasarse en una linea como 'clave=valor'
 - Las configuracion de columnas deben pasarse en cada linea como 'columna=nombreCampo'
 - El fichero Excel debe tener una fila con una cabecera, que es la que se indica en el parameto 'fila'
 - Se puede indicar de forma opcional el numero de hoja en la que estan los datos (por defecto la 1)
<br>

**Uso:**
dseimfacex guion.txt
* Ejemplo de guion:
	* [parametros]
	* entrada=emitidasDiagram.xlsx
	* salida=salida_emitidasDiaram.csv
	* proceso=E00
	* fila=4
	* [columnas]
	* G=fechaFactura
	* C=periodoFactura
	* D=serieFactura
	* E=numeroFactura
	* H=nifFactura
<br>

**Parametros de ejecucion:** 
```
entrada		Fichero excel de entrada
salida		Fichero csv en el que se grabara el resultado (opcional)
proceso		Identifica el tipo de proceso a realizar. Esta formado por una letra y dos numeros ('E'mitidas, 'R'ecibidas)
			Para la importacion 'estandar' de Diagram sera 'E00' para emitidas y 'R00' para recibidas
fila		Fila que contiene la cabecera de las columnas
hoja		(Opcional) Hoja en la que estan los datos, si no se pasa se toma por defecto la 1
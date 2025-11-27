# DseImfacex v2.0.1.0
## Programa para importar un listado de facturas o un balance desde un excel personalizando las columnas de datos

### Desarrollado por Carlos Clemente (04-2025)

### Control de versiones
 - Version 1.0	- Primera version funcional.
 - Version 1.1	- Añadida funcionalidad para agrupar por fechas (tipo E01).
 - Version 1.2	- Añadidos campos en facturas emitidas para recargo de IVA y retencion. 
				- Revision procesoAlcasal para ajustes de varios campos.
				- Añadida funcionalidad para obtener fichero de errores en el proceso.
 - Version 1.3	- Añadida funcionalidad para importar facturas emitidas desde un Excel indicando en que columna esta cada campo
				- Modificado el pase de parametros en la ejecucion
 - Version 1.4	- Añadida funcionalidad para pasar un guion con los parametros y configuracion de columnas
 - Version 1.5	- Incluida biblioteca 'UtilidadesDiagram' como estatica
				- Añadida propiedad 'ficheroFactura' a la clase 'Facturas' (fichero PDF para archivar en el GAD)
				- Añadido metodo 'RecibidasAlcasal' para procesar las recibidas de este cliente
 - Version 1.6	- Añadido metodo 'ChequeoIntegridadFacturas' para comprobar si las facturas son correctas (cuotas de IVA calculadas y total factura)
 - Version 1.7	- Modificado metodo 'ProcesoAlcasal' para incluir nuevas series de ingresos del cliente Alcasal
 - Version 2.0	- Añadida funcionalidad para generar un diario a partir de un balance
 - Version 2.1	- Añadida funcionalidad para que los importes del balance esten en una columna y no importar cuentas sin movimientos
 - Version 2.2	- Añadida funcionalidad para convertir libros de Excel 97-2003 a Excel 2007


<br>
<b>Instrucciones:</b>

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
 - En el caso de importacion de facturas, de cada factura se realiza un chequeo de las cuotas de IVA, cuota de IRPF y total factura, generando un fichero con los errores encontrados.
 - En el caso de importacion de un balance, debe indicarse el parametro 'longitud' para que solo lea las cuentas con esa longitud (evita leer todos los niveles)
<br>

<b>Uso:</b>
dseimfaex dsclave guion.txt

* <u>Parametros de ejecucion:</u>
	* dsclave: Clave de ejecucion del programa
	* guion.txt: Fichero con los datos de ejecucion
<br>
```
* Ejemplo de guion de facturas:
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
	
* Ejemplo de guion de balance 2 columnas:
	* [parametros]
	* entrada=balance.xlsx
	* salida=salida_balance.csv
	* proceso=BAL
	* fila=4
	* longitud=12
	* movimientos=S
	* [columnas]
	* A=Cuenta
	* B=Descripcion
	* G=ImporteDebe
	* H=ImporteHaber
	
* Ejemplo de guion de balance 1 columna:
	* [parametros]
	* entrada=balance.xlsx
	* salida=salida_balance.csv
	* proceso=BAL
	* fila=4
	* longitud=12
	* columna=S
	* [columnas]
	* A=Cuenta
	* B=Descripcion
	* G=ImporteDebe

*Contenido campos:
	entrada		Fichero excel de entrada
	salida		Fichero csv en el que se grabara el resultado (opcional)
	proceso		Identifica el tipo de proceso a realizar. Esta formado por una letra y dos numeros ('E'mitidas, 'R'ecibidas)
				Para la importacion 'estandar' de Diagram sera 'E00' para emitidas y 'R00' para recibidas
				En el caso de importacion de balances sera fijo 'BAL'
	fila		Fila que contiene la cabecera de las columnas
	hoja		(Opcional) Hoja en la que estan los datos, si no se pasa se toma por defecto la 1
	longitud	(solo para balances) Longitud de las cuentas a procesar (evita generar apuntes de todos los niveles que tenga el balance)
	columna		(solo para balances) Indica si los saldos de las cuentas estan en una sola columna (positivos al debe / negativos al haber)
	movimientos	(solo para balances) Indica si recoger las cuentas sin movimientos (defecto N)
	
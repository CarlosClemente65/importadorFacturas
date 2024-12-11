using System;
using System.Collections.Generic;
using System.Text;

namespace importadorFacturas.Metodos
{
    public class ProcesoDiagram
    {
        //Metodo para procesar los datos de las facturas de Diagram
        public StringBuilder ProcesarFacturas(string ficheroEntrada)
        {
            //Devuelve el resultado si hay algun error
            StringBuilder resultado = new StringBuilder();

            Procesos proceso = new Procesos();

            //Carga los campos por defecto a exporta
            Facturas.MapeoFacturas();

            //Lee el fichero con la configuracion de columnas
            proceso.LeerConfiguracionColumnas(Program.ficheroConfiguracion);

            //Instancia una nueva lista de facturas
            Facturas.ListaFacturas = new List<Facturas>();

            //Carga los datos del excel para procesarlos
            var datosExcel = proceso.LeerExcel(ficheroEntrada);

            var numFila = 0; //Controla la fila en la que se ha podido producir un error
            var numColumna = 0;//Controla la columna en la que se ha podido producir un error
            int numeroFactura = 1;

            //Procesado de los datos leidos
            try
            {
                //Procesa cada fila
                foreach(var fila in datosExcel)
                {
                    //Instancia una nueva factura
                    var factura = new Facturas();

                    //Asigna el numero de contador a la factura
                    factura.contador = numeroFactura;

                    //Procesa cada columna y asigna los valores a las propiedades de la clase
                    foreach(var columna in Facturas.mapeoColumnas)
                    {
                        // columna.Key es el índice de la columna
                        // columna.Value es el nombre de la propiedad
                        if(fila.TryGetValue(columna.Key, out var valorCelda))
                        {
                            // Obtener la propiedad en la clase
                            var propiedad = typeof(Facturas).GetProperty(columna.Value);

                            if(propiedad != null && propiedad.CanWrite)
                            {
                                // Verificar si el valor es null antes de convertir y asignar
                                object valorPropiedad = null;

                                if(!string.IsNullOrEmpty(valorCelda))
                                {
                                    //Comprobar si el valorCelda es un numero para evitar algun error al confundir numeros con fechas.
                                    bool esNumero = double.TryParse(valorCelda, out double numero);

                                    //Si no es un numero, intenta convertirlo a una fecha
                                    if(!esNumero && DateTime.TryParse(valorCelda, out DateTime fecha))
                                    {
                                        valorPropiedad = fecha.ToString("dd.MM.yyyy");
                                    }
                                    else
                                    {
                                        // Convertir el valor al tipo de la propiedad y asignarlo
                                        valorPropiedad = Convert.ChangeType(valorCelda, propiedad.PropertyType);
                                    }
                                }
                                propiedad.SetValue(factura, valorPropiedad);
                            }
                        }
                    }

                    //Añade la linea de la factura con sus campos a la lista de facturas y aumenta el contador
                    Facturas.ListaFacturas.Add(factura);
                    numeroFactura++;
                }
                return resultado;
            }

            catch(Exception ex)
            {
                resultado.AppendLine($"Error al procesar los datos en la fila {numFila} y columna {numColumna}. Revise la estructura");
                resultado.AppendLine($"{ex.Message}");
                return resultado;
            }
        }
    }
}
using System;
using System.Text;

namespace importadorFacturas.Metodos
{
    public class ProcesoDiagram
    {
        public StringBuilder EmitidasDiagram(string ficheroEntrada)
        {
            //Metodo para procesar los datos de las emitidas de Diagram

            //Devuelve el resultado si hay algun error
            StringBuilder resultado = new StringBuilder();

            Procesos proceso = new Procesos();
            int filaInicio = 1; //Hay que pasar la fila de la cabecera para contar las columnas
            int columnaInicio = 1; //Los datos empiezan en la columna 1
            int columnaFinal = 66; //Para no tener que procesar todas las columnas se lee hasta la 66 que tiene el total factura

            var datosExcel = proceso.leerExcel(ficheroEntrada, filaInicio, columnaInicio, columnaFinal);

            var numFila = 0; //Controla la fila en la que se ha podido producir un error
            var numColumna = 0;//Controla la columna en la que se ha podido producir un error

            //Proceso de los datos leidos
            try
            {
                foreach(var fila in datosExcel)
                {
                    //Instanciacion de la clase para cada linea
                    var factura = new Facturas();

                    //Asignar valores a las propiedades
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
                                // Convertir el valor al tipo de la propiedad y asignarlo
                                var valorPropiedad = Convert.ChangeType(valorCelda, propiedad.PropertyType);
                                propiedad.SetValue(factura, valorPropiedad);
                            }
                        }
                    }

                    //Añade la linea de la factura con sus campos a la lista de facturas
                    Facturas.ListaFacturas.Add(factura);
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
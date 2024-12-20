﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;

namespace importadorFacturas.Metodos
{
    public class ProcesoDiagram
    {
        //Crea un diccionario con las propiedades de la clase Facturas
        private static readonly Dictionary<string, PropertyInfo> propiedadesClaseFacturas = typeof(Facturas).GetProperties().Where(p => p.CanWrite).ToDictionary(p => p.Name, p => p);

        //Metodo para procesar los datos de las facturas de Diagram
        public StringBuilder ProcesarFacturas()
        {
            //Almacena en resultado si hay algun error
            StringBuilder resultado = new StringBuilder();

            //Carga los campos por defecto a exportar
            Facturas.MapeoFacturas();

            //Instancia una nueva lista de facturas
            Facturas.ListaFacturas = new List<Facturas>();

            //Carga los datos del excel para procesarlos
            var datosExcel = Program.proceso.LeerExcel();

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
                    numFila++; //Se actualiza el valor de la fila para el control de errores

                    //Asigna el numero de contador a la factura
                    factura.contador = numeroFactura;

                    //Procesa cada columna y asigna los valores a las propiedades de la clase
                    foreach(var columna in Facturas.MapeoColumnas)
                    {
                        // columna.Value es el nombre de la propiedad de la clase 'Facturas'
                        // columna.Key es el índice de la columna
                        numColumna = columna.Key; // Se actualiza el valor de la columna para el control de errores

                        AsignarValor(fila, factura, columna);
                    }

                    //Añade la linea de la factura con sus campos a la lista de facturas y aumenta el contador
                    Facturas.ListaFacturas.Add(factura);
                    numeroFactura++; //Se actualiza el numero de la factura
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

        //Metodo para la asignacion del valor de cada celda a la propiedad de la clase que le corresponde
        private static void AsignarValor(Dictionary<int, string> fila, Facturas factura, KeyValuePair<int, string> columna)
        {
            if(fila.TryGetValue(columna.Key, out var valorCelda))
            {
                // Obtener la propiedad desde el diccionario
                if(propiedadesClaseFacturas.TryGetValue(columna.Value, out var propiedad))
                {
                    if(propiedad.CanWrite)
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
        }
    }
}
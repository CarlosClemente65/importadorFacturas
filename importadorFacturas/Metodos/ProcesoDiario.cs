using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using CsvHelper;
using CsvHelper.Configuration;
using DocumentFormat.OpenXml.Spreadsheet;

namespace importadorFacturas.Metodos
{
    public class ProcesoDiario
    {
        //Crea un diccionario con las propiedades de la clase Facturas
        private static readonly Dictionary<string, PropertyInfo> propiedadesClaseDiario = typeof(Diario).GetProperties().Where(p => p.CanWrite).ToDictionary(p => p.Name, p => p);


        public StringBuilder ProcesarDiario()
        {
            // Almacena los errores si se producen
            StringBuilder resultado = new StringBuilder();

            // Crea una lista con los apuntes
            Diario.ApuntesDiario = new List<Diario>();

            int numLinea = 0;

            // Procesando los datos leidos
            try
            {
                // Carga los datos del excel para procesarlos
                var datosExcel = Program.proceso.LeerExcel();

                foreach(var fila in datosExcel)
                {
                    // Crea una nueva linea del diario
                    var lineaDiario = new Diario();

                    foreach(var columna in Diario.MapeoColumnasDiario)
                    {
                        var nombreColumna = columna.Value;
                        var valorCelda = fila[columna.Key]?.ToString().Trim() ?? "";

                        // Solo procesa las lineas que tengan la longitud de cuenta pasada en parametros
                        if(nombreColumna == "Cuenta")
                        {
                            if(valorCelda.Length != Configuracion.LongitudCuenta)
                            {
                                lineaDiario = null;
                                break;
                            }
                        }


                        AsignarValor(fila, lineaDiario, columna);

                        if(lineaDiario.ImporteDebe != 0)
                        {
                            lineaDiario.CuentaDebe = lineaDiario.Cuenta;
                            lineaDiario.Signo = 'D';
                        }
                        else if(lineaDiario.ImporteHaber != 0)
                        {
                            lineaDiario.CuentaHaber = lineaDiario.Cuenta;
                            lineaDiario.Signo = 'H';
                        }
                    }
                    if(lineaDiario != null)
                    {
                        numLinea++;
                        lineaDiario.Apunte = numLinea;
                        Diario.ApuntesDiario.Add(lineaDiario);
                    }
                }
            }

            catch(InvalidOperationException ex)
            {
                resultado.AppendLine($"Error al procesar los datos.");
                resultado.AppendLine($"{ex.Message}");
                return resultado;
            }

            catch(Exception ex)
            {
                resultado.AppendLine($"Error al procesar los datos en la fila {numLinea}. Revise la estructura");
                resultado.AppendLine($"{ex.Message}");
                return resultado;
            }

            return resultado;
        }


        private static void AsignarValor(Dictionary<int, string> fila, Diario lineaDiario, KeyValuePair<int, string> columna)
        {
            //Se pasa como parametros la fila entera para luego obtener el valor de la columna, la instancia de la factura para ir añadiendo propiedades, y la columna que se va a procesar del mapeo de columnas.

            //Intenta obtener el valor de la celda que tiene la columna pasada en el diccionario
            if(fila.TryGetValue(columna.Key, out var valorCelda))
            {
                // Obtener el tipo de la propiedad de la columna pasada
                if(propiedadesClaseDiario.TryGetValue(columna.Value, out var propiedad))
                {
                    if(propiedad.CanWrite)
                    {
                        // Inicializa el valor antes de asignarlo
                        object valorPropiedad = null;

                        //Valida que haya algun dato en la celda antes de asignarlo
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
                        propiedad.SetValue(lineaDiario, valorPropiedad);
                    }
                }
            }
        }

        // Metodo para grabar el csv

    }
}

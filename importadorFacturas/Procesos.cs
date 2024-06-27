using ClosedXML.Excel;
using System.Collections.Generic;
using System.Linq;
using CsvHelper.Configuration;
using CsvHelper;
using System.Globalization;
using System.IO;
using System.Text;
using System;

namespace importadorFacturas
{
    public class Procesos
    {
        public List<Dictionary<int, string>> leerExcel(string fichero,int filaInicio, int columnaInicio, int columnaFinal = 1, int hojaExcel = 1)
        {
            //Metodo para hacer la lectura del Excel y pasarlo a una lista
            //Recibe por parametro el fichero excel a leer, asi como la fila y columna desde la que empezar a leer los datos. La columna final y la hoja de excel se pueden pasar por parametro, si no se pondra por defecto 1
            //La fila debe incluir la cabecera para el procesado posterior (luego se omite en la salida)

            var datosExcel = new List<Dictionary<int, string>>();

            //Se ajusta el numero de la fila y columna de inicio ya que ClosedXML usa base 0
            filaInicio--;
            columnaInicio--;


            using (var libro = new XLWorkbook(fichero))
            {
                var hoja = libro.Worksheet(hojaExcel);

                //Obtiene la cabecera para determinar el numero de columnas
                var cabecera = hoja.Row(filaInicio + 1)
                                    .CellsUsed()
                                    .Select(cell => cell.Address.ColumnNumber)
                                    .Where(colNum => colNum <= columnaFinal)
                                    .ToList();

                //Almacena las filas con datos desde la fila de inicio
                var filas = hoja.RowsUsed().Where(r => r.RowNumber() > filaInicio + 1); //Se suma 1 para saltar la cabecera

                foreach (var fila in filas)
                {
                    //Procesa cada columna en la fila y almacena el valor en la lista datosFilas
                    var datosFilas = new Dictionary<int, string>();
                    foreach (var columna in cabecera)
                    {
                        var cell = fila.Cell(columna);
                        datosFilas[columna] = cell.GetValue<string>();
                    }
                    datosExcel.Add(datosFilas);
                }
            }

            return datosExcel;
        }

        public StringBuilder grabarCsv<T>(string ficheroSalida, List<T> datos)
        {
            var resultado = new StringBuilder();

            try
            {
                //Metodo para grabar el fichero de salida en csv
                var csvConfig = new CsvConfiguration(CultureInfo.CurrentCulture)
                {
                    //Se almacena sin cabecera y con el separador de punto y coma
                    HasHeaderRecord = false,
                    Delimiter = ";"
                };

                using (var writer = new StreamWriter(ficheroSalida, false, System.Text.Encoding.UTF8))
                {
                    using (var csv = new CsvWriter(writer, csvConfig))
                    {
                        csv.WriteRecords(datos);
                    }
                }
                return resultado;
            }
            catch (Exception ex)
            {
                resultado.AppendLine($"Error al grabar el fichero de salida");
                resultado.AppendLine(ex.Message);
                return resultado;
            }
        }

    }
}

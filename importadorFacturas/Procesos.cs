using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Reflection;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Spreadsheet;
using CsvHelper.Configuration;
using CsvHelper;
using System.Globalization;
using System.IO;

namespace importadorFacturas
{
    public class Procesos
    {
        public List<Dictionary<int, string>> leerExcel(string fichero, int filaInicio, int columnaInicio)
        {
            //metodo para hacer la lectura del Excel y pasarlo a una lista
            //Recibe por parametro el fichero excel a leer, asi como la fila y columna desde la que empezar a leer los datos.
            //La fila debe incluir la cabecera para el procesado posterior (luego se omite en la salida)

            var datosExcel = new List<Dictionary<int, string>>();

            //Se ajusta el numero de la fila y columna de inicio ya que ClosedXML usa base 0
            filaInicio--;
            columnaInicio--;

            using (var libro = new XLWorkbook(fichero))
            {
                var hoja = libro.Worksheet(1);

                //Obtiene la cabecera para determinar el numero de columnas
                var cabecera = hoja.Row(filaInicio+1).CellsUsed().Select(cell => cell.Address.ColumnNumber).ToList();

                //Almacena las filas con datos desde la fila de inicio
                var filas = hoja.RowsUsed().Where(r => r.RowNumber() > filaInicio + 1); //Se suma 1 para saltar la cabecera

                foreach(var fila in filas)
                {
                    var datosFilas = new Dictionary<int, string>();

                    //Saltamos la fila de inicio que sera la cabecera de los datos
                    foreach(var columna in cabecera)
                    {
                        var cell = fila.Cell(columna);
                        datosFilas[columna] = cell.GetValue<string>();
                    }
                    datosExcel.Add(datosFilas);
                }
            }

            return datosExcel;
        }

        public void grabarCsv<T>(string ficheroSalida, List<T> datos)
        {
            //Metodo para grabar el fichero de salida en csv
            var csvConfig = new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                HasHeaderRecord = false,
                Delimiter = ";"
            };

            using (var writer = new StreamWriter(ficheroSalida))
            {
                using (var csv = new CsvWriter(writer, csvConfig))
                {
                    csv.WriteRecords(datos);
                }
            }
        }

    }
}

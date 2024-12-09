using ClosedXML.Excel;
using System.Collections.Generic;
using System.Linq;
using CsvHelper.Configuration;
using System.Globalization;
using System.IO;
using System.Text;
using System;
using static importadorFacturas.Facturas;

namespace importadorFacturas
{
    public class Procesos
    {
        public List<Dictionary<int, string>> leerExcel(string fichero, int filaInicio, int columnaInicio, int columnaFinal = 1, int hojaExcel = 1)
        {
            //Metodo para hacer la lectura del Excel y pasarlo a una lista
            //Recibe por parametro el fichero excel a leer, asi como la fila y columna desde la que empezar a leer los datos. La columna final y la hoja de excel se pueden pasar por parametro, si no se pondra por defecto 1
            //La fila debe incluir la cabecera para el procesado posterior (luego se omite en la salida)

            var datosExcel = new List<Dictionary<int, string>>();

            //Se ajusta el numero de la fila y columna de inicio ya que ClosedXML usa base 0
            filaInicio--;
            columnaInicio--;


            using(var libro = new XLWorkbook(fichero))
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
                if(Program.tipoProceso == "E00")
                {
                    //ColumnasAexportar = filas.ToArray();
                }

                foreach(var fila in filas)
                {
                    //Procesa cada columna en la fila y almacena el valor en la lista datosFilas
                    var datosFilas = new Dictionary<int, string>();
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

        public StringBuilder grabarCsv<T>(string ficheroSalida, List<T> datos, string[] camposAexportar)
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

                //var encoding = new System.Text.UTF8Encoding(false);
                using(var writer = new StreamWriter(ficheroSalida, false, Encoding.Default))
                {
                    foreach(var dato in datos)
                    {
                        var fila = new List<string>();

                        // Filtrar las propiedades que están en el array propiedadesExportables y ordenar por el atributo OrdenCsv
                        var propiedades = typeof(T).GetProperties()
                            .Where(prop => camposAexportar.Contains(prop.Name)) // Filtra por el nombre de propiedad
                            .Where(prop => prop.GetCustomAttributes(typeof(OrdenCsvAttribute), false).Any()) // Asegura que la propiedad tenga el atributo
                            .OrderBy(prop => ((OrdenCsvAttribute)prop.GetCustomAttributes(typeof(OrdenCsvAttribute), false).First()).Orden); // Ordena por el valor del atributo


                        // Recorre las propiedades y añade solo las que tienen valor
                        foreach(var propiedad in propiedades)
                        {
                            var valor = propiedad.GetValue(dato);
                            //Añade el valor de la propiedad a la fila
                            if(valor != null)
                            {
                                fila.Add(valor.ToString());
                            }
                            else
                            {
                                fila.Add(""); // Si no tiene valor, se añade una celda vacía
                            }
                        }

                        // Escribe la fila al archivo
                        if(fila.Count > 0)
                        {
                            writer.WriteLine(string.Join(csvConfig.Delimiter, fila));
                        }
                    }
                }
                return resultado;
            }
            catch(Exception ex)
            {
                resultado.AppendLine($"Error al grabar el fichero de salida");
                resultado.AppendLine(ex.Message);
                return resultado;
            }
        }

        public static Dictionary<int, string> LeerCsv(string rutaCsv, out string[] propiedades)
        {
            var mapeoColumnas = new Dictionary<int, string>();
            var listaPropiedades = new List<string>();

            // Leer el archivo línea por línea
            foreach(var linea in File.ReadLines(rutaCsv))
            {
                if(string.IsNullOrWhiteSpace(linea)) continue; // Saltar líneas vacías

                var partes = linea.Split(';');
                if(partes.Length != 2) continue; // Saltar líneas mal formateadas

                string letraColumna = partes[0].Trim();
                string propiedad = partes[1].Trim();

                // Convertir la letra de columna a número
                int numeroColumna = LetraAColumna(letraColumna);
                if(numeroColumna <= 0) continue; // Saltar letras inválidas

                // Almacenar en el diccionario y la lista de propiedades
                mapeoColumnas[numeroColumna] = propiedad;
                listaPropiedades.Add(propiedad);
            }

            propiedades = listaPropiedades.ToArray();

            return mapeoColumnas;
        }

        private static int LetraAColumna(string letraColumna)
        {
            // Método para convertir letras de columna a número
            int columna = 0;
            foreach(char letra in letraColumna.ToUpper())
            {
                if(letra < 'A' || letra > 'Z') return -1; // Caracter inválido
                columna = columna * 26 + (letra - 'A' + 1);
            }
            return columna;
        }

    }
}

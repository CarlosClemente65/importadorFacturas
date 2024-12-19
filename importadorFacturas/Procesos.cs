using ClosedXML.Excel;
using System.Collections.Generic;
using System.Linq;
using CsvHelper.Configuration;
using System.Globalization;
using System.IO;
using System.Text;
using System;

namespace importadorFacturas
{
    public class Procesos
    {
        //Metodo para hacer la lectura del Excel y pasarlo a una lista
        public List<Dictionary<int, string>> LeerExcel(Configuracion parametros)
        {
            //Recibe por parametro el fichero excel a leer y la hoja de excel se puede pasar por parametro, si no se pondra por defecto 1
            //La fila debe incluir la cabecera para el procesado posterior (luego se omite en la salida)

            //La fila de inicio se pasa por parametro y se almacena en una propiedad de la clase 'Program'
            int filaInicio = parametros.FilaInicio;
            int hojaExcel = parametros.HojaExcel;
            string ficheroExcel = parametros.FicheroEntrada;

            var datosExcel = new List<Dictionary<int, string>>();

            //Se ajusta el numero de la fila y columna de inicio ya que ClosedXML usa base 0
            filaInicio--;

            using(var libro = new XLWorkbook(ficheroExcel))
            {
                var hoja = libro.Worksheet(hojaExcel);

                //Obtiene la cabecera para determinar el numero de columnas
                //var cabecera2 = new List<int>(Facturas.mapeoColumnas.Keys);

                var cabecera = hoja.Row(filaInicio + 1)
                                    .Cells()
                                    .Select(c => c.Address.ColumnNumber)
                                    .ToList();

                //Almacena las filas con datos desde la fila de inicio
                var filas = hoja.RowsUsed().Where(r => r.RowNumber() > filaInicio + 1); //Se suma 1 para saltar la cabecera

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

        //Metodo para grabar el fichero de salida en formato csv
        public StringBuilder GrabarCsv<T>(string ficheroSalida, List<T> datos, string[] camposAexportar)
        {
            //Variable que recoje los posibles errores
            var resultado = new StringBuilder();

            try
            {
                //Configuracion de la salida csv para que no ponga cabecera y el separador sea un punto y coma
                var csvConfig = new CsvConfiguration(CultureInfo.CurrentCulture)
                {
                    //Se almacena sin cabecera y con el separador de punto y coma
                    HasHeaderRecord = false,
                    Delimiter = ";"
                };

                using(var writer = new StreamWriter(ficheroSalida, false, Encoding.Default))
                {
                    //Procesado de cada fila
                    foreach(var dato in datos)
                    {
                        var fila = new List<string>();

                        // Filtrar las propiedades que están en la variable de 'camposAexportar y ordenar por el atributo OrdenCsv
                        var propiedades = typeof(T).GetProperties()
                            .Where(prop => camposAexportar.Contains(prop.Name)) // Filtra por el nombre de propiedad
                            .Where(prop => prop.GetCustomAttributes(typeof(Facturas.OrdenCsvAttribute), false).Any()) // Asegura que la propiedad tenga el atributo
                            .OrderBy(prop => ((Facturas.OrdenCsvAttribute)prop.GetCustomAttributes(typeof(Facturas.OrdenCsvAttribute), false).First()).Orden); // Ordena por el valor del atributo


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
                                fila.Add(""); // Si no tiene valor, se añade un valor vacío
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

        //Metodo para leer el fichero con la configuracion de columnas
        public void LeerConfiguracionColumnas(string rutaCsv)
        {
            //Inicializa la lista de propiedades y el mapeo de columnas
            var listaPropiedades = new List<string>();
            Facturas.mapeoColumnas = new Dictionary<int, string>();

            //Leer el archivo línea por línea
            foreach(var linea in File.ReadLines(rutaCsv))
            {
                if(string.IsNullOrWhiteSpace(linea)) continue; //Salta líneas vacías
                
                //Divide la cadena por el primer punto y coma que encuentra
                (string letraColumna, string propiedad) = Program.utiles.DivideCadena(linea, ';');

                // Convertir la letra de columna a número
                int numeroColumna = LetraAColumna(letraColumna);
                if(numeroColumna <= 0) continue; // Saltar letras inválidas

                // Almacenar en el diccionario y la lista de propiedades
                Facturas.mapeoColumnas[numeroColumna] = propiedad;
                listaPropiedades.Add(propiedad);
            }
        }

        //Metodo para convertir cada letra de la configuracion en el numero de columna
        private int LetraAColumna(string letraColumna)
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

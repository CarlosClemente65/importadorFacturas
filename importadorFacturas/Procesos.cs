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
        //Metodo para chequear si existe el fichero pasado
        private static string ChequeoFichero(string fichero)
        {
            string resultado = string.Empty;
            if(!File.Exists(fichero))
            {
                resultado = $"Error. No existe el fichero {fichero}";
            }
            return resultado;
        }

        //Metodo para hacer la lectura del Excel y pasarlo a una lista
        public List<Dictionary<int, string>> LeerExcel()
        {
            //Recibe por parametro el fichero excel a leer y la hoja de excel se puede pasar por parametro, si no se pondra por defecto 1
            //La fila debe incluir la cabecera para el procesado posterior (luego se omite en la salida)

            //La fila de inicio se pasa por parametro y se almacena en una propiedad de la clase 'Program'
            int filaInicio = Configuracion.FilaInicio;
            int hojaExcel = Configuracion.HojaExcel;
            string ficheroExcel = Configuracion.FicheroEntrada;

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
        public StringBuilder GrabarCsv<T>(List<T> datos, string[] camposAexportar)
        {
            string ficheroSalida = Configuracion.FicheroSalida;
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

        //Metodo para hacer la carga del guion
        public bool CargarGuion(string ficheroGuion)
        {
            using(var contenidoGuion = new StreamReader(ficheroGuion))
            {
                string linea;
                bool procesaParametros = false;
                bool procesaColumnas = false;

                while((linea = contenidoGuion.ReadLine()) != null)
                {
                    linea = linea.Trim();

                    if(string.IsNullOrWhiteSpace(linea))
                    {
                        continue;
                    }

                    if(linea.Equals("[parametros]", StringComparison.OrdinalIgnoreCase))
                    {
                        procesaParametros = true;
                        procesaColumnas = false;
                        continue;
                    }

                    if(linea.Equals("[columnas]", StringComparison.OrdinalIgnoreCase))
                    {
                        procesaParametros = false;
                        procesaColumnas = true;
                        continue;
                    }

                    if(procesaParametros)
                    {
                        Configuracion.parametros.Add(linea);
                    }
                    else if(procesaColumnas)
                    {
                        Configuracion.columnas.Add(linea);
                    }
                }
            }

            if(ProcesarParametros()) //Si no ha habido errores procesa las columnas
            {
                LeerConfiguracionColumnas(Configuracion.columnas);
                return true; //Se devuelve true porque no ha habido errores
            }

            return false; //Si en el procesado de parametros ha habido algun error

        }

        private bool ProcesarParametros()
        {
            /* Parametros:
             * entrada=fichero de entrada
             * salida=fichero de salida (opcional)
             * proceso=tipo de proceso a ejecutar
             * fila=fila de la cabecera con los nombres de las columnas. Nota: tiene que haber una fila que tenga todas las columnas rellenas para luego poder procesar bien los datos
             * hoja=hoja en la que estan los datos (opcional)
             * */

            //Almacena los errores que se puedan producir
            string chequeo = string.Empty;

            //Procesa las lineas
            foreach(string linea in Configuracion.parametros)
            {
                //Separa el parametro y su valor
                (string parametro, string valor) = Program.utiles.DivideCadena(linea, '=');

                switch(parametro)
                {
                    //Fichero entrada
                    case "entrada":
                        chequeo = ChequeoFichero(valor);

                        if(!string.IsNullOrEmpty(chequeo))
                        {
                            Program.utiles.GrabarFichero(Configuracion.FicheroErrores, chequeo);
                            return false;
                        }
                        Configuracion.FicheroEntrada = valor;

                        if(string.IsNullOrEmpty(Configuracion.FicheroSalida))
                        {
                            Configuracion.FicheroSalida = $"salida_{Path.GetFileNameWithoutExtension(Configuracion.FicheroEntrada)}.csv";
                            Program.utiles.ControlFicheros(Configuracion.FicheroSalida);
                        }
                        Configuracion.FicheroErrores = Path.Combine(Path.GetDirectoryName(Configuracion.FicheroEntrada), "errores.txt");
                        Program.utiles.ControlFicheros(Configuracion.FicheroErrores);

                        break;

                    //Fichero salida
                    case "salida":
                        Configuracion.FicheroSalida = valor;
                        Program.utiles.ControlFicheros(Configuracion.FicheroSalida);
                        break;

                    //Tipo de proceso
                    case "proceso":
                        //Se valida que el tipo de proceso sea alguno de los definidos en la clase 'Configuracion'
                        if(Enum.TryParse<Configuracion.TiposProceso>(valor, out Configuracion.TiposProceso _tipoProceso))
                        {
                            Configuracion.TipoProceso = _tipoProceso.ToString();
                        }
                        else
                        {
                            Program.utiles.GrabarFichero(Configuracion.FicheroErrores, $"Error. Tipo de proceso {valor} incorrecto");
                            return false;
                        }
                        break;

                    //Fila de cabecera de columnas
                    case "fila":
                        //Se valida que la fila sea un numero mayor que 0
                        if(int.TryParse(valor, out int _fila) && _fila > 0)
                        {
                            Configuracion.FilaInicio = _fila;
                        }
                        else
                        {
                            Program.utiles.GrabarFichero(Configuracion.FicheroErrores, $"Error. Fila {valor} incorrecta");
                            return false;
                        }
                        break;

                    //Hoja en la que estan los datos (opcional)
                    case "hoja":
                        //Se valida que la hoja sea un numero mayor que 0
                        if(int.TryParse(valor, out int _hoja) && _hoja > 0)
                        {
                            Configuracion.HojaExcel = _hoja;
                        }
                        else
                        {
                            Program.utiles.GrabarFichero(Configuracion.FicheroErrores, $"Error. Hoja {valor} incorrecta");
                            return false;
                        }
                        break;
                }
            }
            return true;
        }

        //Metodo para leer el fichero con la configuracion de columnas
        private void LeerConfiguracionColumnas(List<string> lineas)
        {
            //Devuelve el control si no se han pasado las columnas (proceso Alcasal)
            if(lineas.Count == 0)
            {
                return;
            }

            //Inicializa la lista de propiedades y el mapeo de columnas
            var listaPropiedades = new List<string>();
            Facturas.MapeoColumnas = new Dictionary<int, string>();

            //Leer el archivo línea por línea
            foreach(var linea in lineas)
            {
                //Divide la cadena por el primer punto y coma que encuentra
                (string letraColumna, string propiedad) = Program.utiles.DivideCadena(linea, ';');

                // Convertir la letra de columna a número
                int numeroColumna = LetraAColumna(letraColumna);
                if(numeroColumna <= 0) continue; // Saltar letras inválidas

                // Almacenar en el diccionario y la lista de propiedades
                Facturas.MapeoColumnas[numeroColumna] = propiedad;
                listaPropiedades.Add(propiedad);
            }

            //Se añade el campo 'contador' para que se incluya en el fichero de salida
            Facturas.ColumnasAexportar = new List<string> { "contador" };
            Facturas.ColumnasAexportar.AddRange(listaPropiedades);
        }
    }
}

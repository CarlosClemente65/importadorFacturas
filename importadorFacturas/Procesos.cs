using ClosedXML.Excel;
using System.Collections.Generic;
using System.Linq;
using CsvHelper.Configuration;
using System.Globalization;
using System.IO;
using System.Text;
using System;
using UtilidadesDiagram;
using DocumentFormat.OpenXml.Drawing.Diagrams;

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

            var datosExcel = new List<Dictionary<int, string>>();

            //Se ajusta el numero de la fila y columna de inicio ya que ClosedXML usa base 0
            filaInicio--;

            using(var libro = new XLWorkbook(Configuracion.FicheroEntrada))
            {
                var hoja = libro.Worksheet(Configuracion.HojaExcel);

                //Carga las columnas de la fila de cabecera para procesarlas
                var cabecera = hoja.Row(filaInicio + 1)
                                    .Cells()
                                    .Select(c => c.Address.ColumnNumber)
                                    .ToList();

                //Almacena las filas con datos desde la fila de inicio
                var filas = hoja.RowsUsed().Where(r => r.RowNumber() > filaInicio + 1); //Se suma 1 para saltar la cabecera

                foreach(var fila in filas)
                {
                    //Procesa cada columna en la fila y almacena el valor en la lista datosExcel
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

                using(var writer = new StreamWriter(Configuracion.FicheroSalida, false, Encoding.Default))
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

                //Procesa todas las lineas validas
                while((linea = contenidoGuion.ReadLine()) != null)
                {
                    linea = linea.Trim();

                    //No procesa las lineas vacias
                    if(string.IsNullOrWhiteSpace(linea))
                    {
                        continue;
                    }

                    //Controla si es la zona de parametros
                    if(linea.Equals("[parametros]", StringComparison.OrdinalIgnoreCase))
                    {
                        procesaParametros = true;
                        procesaColumnas = false;
                        continue;
                    }

                    //Controla si es la zona de columnas
                    if(linea.Equals("[columnas]", StringComparison.OrdinalIgnoreCase))
                    {
                        procesaParametros = false;
                        procesaColumnas = true;
                        continue;
                    }

                    //Se añade la linea segun si esta en la zona de parametros o de columnas
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

            //Carga los parametros a las propiedades de la clase 'Configuracion'
            if(ProcesarParametros())
            {
                //Si no ha habido errores procesa las columnas
                LeerConfiguracionColumnas(Configuracion.columnas);
                return true; //Se devuelve true porque no ha habido errores
            }

            return false; //Si en el procesado de parametros ha habido algun error devuelve false

        }

        //Metodo para procesar los parametros del guion
        private bool ProcesarParametros()
        {
            //Variable para almacenar los errores
            string chequeo = string.Empty;

            //Procesa las lineas
            foreach(string linea in Configuracion.parametros)
            {
                //Separa el parametro y su valor
                (string parametro, string valor) = Utilidades.DivideCadena(linea, '=');

                switch(parametro)
                {
                    //Fichero entrada
                    case "entrada":
                        chequeo = ChequeoFichero(valor); //Devuelve los errores que se hayan producido

                        if(!string.IsNullOrEmpty(chequeo))
                        {
                            Utilidades.GrabarFichero(Configuracion.FicheroErrores, chequeo);
                            return false;
                        }
                        Configuracion.FicheroEntrada = valor;

                        //Graba el defecto del fichero de salida segun el nombre del fichero de entrada
                        if(string.IsNullOrEmpty(Configuracion.FicheroSalida))
                        {
                            Configuracion.FicheroSalida = $"salida_{Path.GetFileNameWithoutExtension(Configuracion.FicheroEntrada)}.csv";
                            Utilidades.ControlFicheros(Configuracion.FicheroSalida);
                        }

                        //Graba el defecto del fichero de errores segun el nombre del fichero de entrada
                        Configuracion.FicheroErrores = Path.Combine(Path.GetDirectoryName(Configuracion.FicheroEntrada), "errores.txt");
                        Utilidades.ControlFicheros(Configuracion.FicheroErrores);

                        break;

                    //Fichero salida
                    case "salida":
                        Configuracion.FicheroSalida = valor;
                        Utilidades.ControlFicheros(Configuracion.FicheroSalida);
                        break;

                    //Tipo de proceso
                    case "proceso":
                        //Se valida que el tipo de proceso sea alguno de los definidos en 'Configuracion.TiposProceso'
                        if(Enum.TryParse<Configuracion.TiposProceso>(valor, out Configuracion.TiposProceso _tipoProceso))
                        {
                            Configuracion.TipoProceso = _tipoProceso.ToString();
                        }
                        else
                        {
                            Utilidades.GrabarFichero(Configuracion.FicheroErrores, $"Error. Tipo de proceso {valor} incorrecto");
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
                            Utilidades.GrabarFichero(Configuracion.FicheroErrores, $"Error. Fila {valor} incorrecta");
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
                            Utilidades.GrabarFichero(Configuracion.FicheroErrores, $"Error. Hoja {valor} incorrecta");
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

            //Genera la lista de las columnas a exportar segun el defecto
            Facturas.MapeoFacturas();
            Facturas.ColumnasAexportar = new List<string> { "contador" };
            Facturas.ColumnasAexportar.AddRange(Facturas.MapeoColumnas.Values);

            //Crea una nueva instancia para cargar las columnas que vienen en la configuracion
            Facturas.MapeoColumnas = new Dictionary<int, string>();

            //Procesa las lineas
            foreach(var linea in lineas)
            {
                //Divide la cadena por el simbolo igual 
                (string letraColumna, string propiedad) = Utilidades.DivideCadena(linea, '=');

                // Convertir la letra de columna a número
                int numeroColumna = LetraAColumna(letraColumna);
                if(numeroColumna <= 0) continue; // Saltar letras inválidas

                // Almacenar en el diccionario el numero de columna y el nombre del campo
                Facturas.MapeoColumnas[numeroColumna] = propiedad;
            }
        }

        public List<string[]> LeerCsv(string ficheroEntrada)
        {
            List<string[]> datosEntrada = new List<string[]>();

            try
            {
                foreach(string linea in File.ReadLines(ficheroEntrada))
                {
                    string[] elementos = linea.Split(';');
                    datosEntrada.Add(elementos);
                }

                return datosEntrada;
            }

            catch(Exception ex)
            {
                throw new Exception($"No se ha podido leer el fichero de entrada.\n{ex.Message}");
            }
        }

        //Permite chequear si hay alguna cuota de IVA que no esta bien calculada
        public string ChequeoCuotaIva<T>(T factura, Func<T, decimal> obtenerBase, Func<T, float> obtenerPorentaje, Func<T, decimal> obtenerCuota, int numFila) where T : Facturas
        {
            string resultado = string.Empty;
            //Almacena los valores de las propiedades segun la base, porcentaje y cuota pasada al metodo
            decimal baseFactura = obtenerBase(factura);
            float porcentajeIva = obtenerPorentaje(factura);
            decimal cuotaIva = obtenerCuota(factura);

            decimal cuotaCalculada = Math.Round(baseFactura * (decimal)porcentajeIva / 100, 2);//Calculo de la cuota que correspnde al tipo pasado

            //Si la cuota calculada difiere en mas o menos 5 centimos, genera el error en la factura.
            if(Math.Abs(cuotaIva - cuotaCalculada) > 0.05m)
            {
                resultado = $"\t- La cuota de IVA calculada al {porcentajeIva}% no es correcta. Cuota calculada: {cuotaCalculada} - Cuota del fichero: {cuotaIva}";
            }

            return resultado;
        }

        //Metodo para chequear si las bases y cuotas son correctas con los porcentajes de IVA correspondientes
        public StringBuilder ChequeoIntegridadFacturas<T>(List<T> facturas) where T : Facturas
        {
            StringBuilder resultado = new StringBuilder();
            int numLinea = 1;

            foreach(var factura in facturas)
            {
                decimal totalFacturaCalculado = 0; //Acumula el total factura de cada campo
                string resultadoChequeo = string.Empty;
                bool flag = false; //Control para poner la cabecera de los chequeos de cada factura.

                for(int i = 1; i <= 10; i++)
                {
                    // Usamos reflexión para obtener las propiedades baseFacturaX, porcentajeIvaX y cuotaIvaX
                    var baseFacturaProp = factura.GetType().GetProperty($"baseFactura{i}");
                    var porcentajeIvaProp = factura.GetType().GetProperty($"porcentajeIva{i}");
                    var cuotaIvaProp = factura.GetType().GetProperty($"cuotaIva{i}");
                    var cuotaRecargoProp = factura.GetType().GetProperty($"cuotaRecargo{i}");

                    if(baseFacturaProp != null && porcentajeIvaProp != null && cuotaIvaProp != null)
                    {
                        // Obtener los valores de las propiedades mediante reflexión
                        decimal baseFactura = (decimal)baseFacturaProp.GetValue(factura);
                        float porcentajeIva = (float)porcentajeIvaProp.GetValue(factura);
                        decimal cuotaIva = (decimal)cuotaIvaProp.GetValue(factura);
                        decimal cuotaRecargo = (decimal)cuotaRecargoProp.GetValue(factura);
                        totalFacturaCalculado += baseFactura + cuotaIva + cuotaRecargo; //Se acumula cada campo al total de factura calculado

                        // Llamamos al método de chequeo y acumulamos los errores en el StringBuilder
                        resultadoChequeo = ChequeoCuotaIva(factura, f => baseFactura, f => porcentajeIva, f => cuotaIva, numLinea);

                        // Si hay algún error, lo agregamos al resultado final
                        if(!string.IsNullOrEmpty(resultadoChequeo))
                        {
                            if(!flag) //Permite añadir una cabecera por cada factura
                            {
                                resultado.AppendLine($"\nDescuadres en la factura de la linea {numLinea} del proveedor {factura.nombreFactura} y fecha {factura.fechaFactura}:");
                                flag = true;
                            }
                            resultado.AppendLine(resultadoChequeo);
                        }
                    }
                }

                //Chequeo si cuadra la cuota de IRPF con el porcentaje pasado (solo si se han pasado la base y porcentaje)
                if(factura.baseIrpf != 0 && factura.porcentajeIrpf != 0)
                {
                    //Calcula la cuota de IRPF que corresponde 
                    decimal cuotaIrpfCalculada = Math.Round(factura.baseIrpf * (decimal)factura.porcentajeIrpf / 100, 2);

                    //Permite una diferencia en mas/menos 5 centimos
                    if(Math.Abs(cuotaIrpfCalculada - factura.cuotaIrpf) > 0.05m)
                    {
                        resultado.AppendLine($"\t- La cuota de IRPF calculada al {factura.porcentajeIrpf}% no es correcta. Cuota calculada: {cuotaIrpfCalculada} - Cuota del fichero: {factura.cuotaIrpf}. Revise si el porcentaje informado es correcto.");
                    }
                }

                //Chequeo si cuadra el total factura calculado con el pasado en el fichero
                if(totalFacturaCalculado - factura.cuotaIrpf != factura.totalFactura)
                {
                    resultado.AppendLine($"\t- El total de factura calculado no es correcto. Importe calculado: {totalFacturaCalculado} - Importe de la factura: {factura.totalFactura}. Revise si falta alguna base o cuota de IVA");
                }
                numLinea++;
                flag = false;
            }
            return resultado;
        }
    }
}

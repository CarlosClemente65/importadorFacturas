using CsvHelper;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text;


namespace importadorFacturas
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string ficheroEntrada = string.Empty;
            string ficheroSalida = string.Empty;
            string tipoProceso = string.Empty;
            
            if (args.Length == 0)
            {
                return;
            }

            //Se pueden pasar 3 parametros: el primero es el tipo de proceso que se puede usar para futuras importaciones, el segundo es el fichero excel a leer, y el tercero es el fichero de salida aunque este es opcional

            tipoProceso = args[0];

            ficheroEntrada = args[1];
            if (!File.Exists(ficheroEntrada)) return;

            ficheroSalida = Path.ChangeExtension(ficheroEntrada, "csv");
            if (args.Length > 2)
            {
                ficheroSalida = args[2];
                if (File.Exists(ficheroSalida)) File.Delete(ficheroSalida);
            }

            string ficheroErrores = Path.Combine(Path.GetDirectoryName(ficheroEntrada), "errores.txt");
            if (File.Exists(ficheroErrores)) File.Delete(ficheroErrores);

            procesarFichero(tipoProceso, ficheroEntrada, ficheroSalida);

        }


        private static void procesarFichero(string tipoProceso, string ficheroEntrada, string ficheroSalida)
        {
            StringBuilder resultado = new StringBuilder();
            //Metodo para leer el fichero Excel y procesar los datos segun el tipo pasado por parametro

            //Nota: el tipo de proceso debe ser una letra (E para ventas y R para compras) seguido de dos numeros (hasta 99 importaciones diferentes)

            Procesos proceso = new Procesos();

            switch (tipoProceso)
            {
                case "E01":
                    //Facuras emitidas de Alcasal (cliente de Raiña Asesores) tiquet 5863-37

                    //Inicializa campos
                    procesoAlcasal metodo = new procesoAlcasal();

                    //Procesar los datos del fichero de Excel
                    resultado = metodo.emitidasAlcasar(ficheroEntrada);

                    //Carga los datos procesados para pasarlos al csv
                    List<facturasEmitidas> datosProcesados = facturasEmitidas.obtenerDatos();
                    if (datosProcesados.Count > 0)
                    {
                        resultado = proceso.grabarCsv(ficheroSalida, datosProcesados);
                    }

                    break;
            }

            //Grabar el registro de errores si se ha producido alguno
            if (resultado.Length > 0)
            {
                string rutaLog = Path.Combine(Path.GetDirectoryName(ficheroEntrada), "errores.txt");
                File.WriteAllText(rutaLog, resultado.ToString(), Encoding.UTF8);
            }
        }
    }
}

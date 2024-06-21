using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CsvHelper;
using CsvHelper.Configuration;


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

            //Solo se pasan 2 parametros: el primero es el tipo de proceso que se puede usar para futuras importaciones, el segundo es el fichero excel a leer
            //Nota: el tipo de proceso debe ser una letra (E para ventas y R para compras) seguido de dos numeros (hasta 99 importaciones diferentes)
            tipoProceso = args[0];

            ficheroEntrada = args[1];
            if (!File.Exists(ficheroEntrada)) return;

            ficheroSalida = Path.ChangeExtension(ficheroEntrada, "csv");
            if (args.Length > 2)
            {
                ficheroSalida = args[2];
                if (File.Exists(ficheroSalida)) File.Delete(ficheroSalida);
            }

            procesarFichero(tipoProceso, ficheroEntrada, ficheroSalida);

        }


        private static void procesarFichero(string tipoProceso, string ficheroEntrada, string ficheroSalida)
        {
            //Metodo para leer el fichero Excel y procesar los datos

            Procesos proceso = new Procesos();

            switch (tipoProceso)
            {
                case "E01":
                    procesoAlcasar metodo = new procesoAlcasar();

                    metodo.emitidasAlcasar(ficheroEntrada);

                    List<ingresosAlcasar> datosProcesados = ingresosAlcasar.obtenerDatos();
                    proceso.grabarCsv(ficheroSalida, datosProcesados);
                    
                    break;
            }
        }
    }
}

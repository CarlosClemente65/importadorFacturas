using importadorFacturas.Metodos;
using System.Collections.Generic;
using System.IO;
using System.Text;


namespace importadorFacturas
{
    internal class Program
    {
        static string ficheroEntrada = string.Empty;
        static string ficheroSalida = string.Empty;
        static string ficheroErrores = string.Empty;
        static string tipoProceso = string.Empty;

        //Instanciacion de las utilidades para acceso a los metodos
        public static Utilidades utiles = new Utilidades();
        static void Main(string[] args)
        {
            if(args.Length == 0)
            {
                return;
            }

            //Se pueden pasar 3 parametros: el primero es el tipo de proceso que se puede usar para futuras importaciones, el segundo es el fichero excel a leer, y el tercero es el fichero de salida aunque este es opcional

            tipoProceso = args[0];

            ficheroEntrada = args[1];
            if(!File.Exists(ficheroEntrada)) return;

            ficheroSalida = Path.ChangeExtension(ficheroEntrada, "csv");
            if(args.Length > 2)
            {
                ficheroSalida = args[2];
            }
            utiles.ControlFicheros(ficheroSalida);

            ficheroErrores = Path.Combine(Path.GetDirectoryName(ficheroEntrada), "errores.txt");
            utiles.ControlFicheros(ficheroErrores);

            procesarFichero(tipoProceso, ficheroEntrada, ficheroSalida);

        }


        private static void procesarFichero(string tipoProceso, string ficheroEntrada, string ficheroSalida)
        {
            //Metodo para leer el fichero Excel y procesar los datos segun el tipo pasado por parametro
            //Nota: el tipo de proceso debe ser una letra (E para emitidas y R para recibidas) seguido de dos numeros (hasta 99 importaciones diferentes de cada tipo)

            //Variable que recoge el texto devuelto en el metodo si se ha producido algun error en el procesado
            StringBuilder resultado = new StringBuilder();

            Procesos proceso = new Procesos();

            switch(tipoProceso)
            {
                case "E00":
                    //Facturas emitidas con formato 5 IVAs de diagram (pendiente de desarrollo)
                    //Inicializa campos
                    ProcesoDiagram procesoDiagram = new ProcesoDiagram();
                    /*
                     * TO_DO. Pendiente desarrollar (hacer similar al E01)
                     */
                    break;

                case "E01":
                    //Facuras emitidas de Alcasal (cliente de Raiña Asesores) tiquet 5863-37

                    //Inicializa campos
                    procesoAlcasal metodo = new procesoAlcasal();

                    //Procesar los datos del fichero de Excel
                    resultado = metodo.emitidasAlcasar(ficheroEntrada);

                    //Carga los datos procesados para pasarlos al csv
                    List<EmitidasE01> datosProcesados = EmitidasE01.ObtenerDatos();
                    if(datosProcesados.Count > 0)
                    {
                        //Array de propiedades a exportar de este tipo
                        string []camposAexportar = EmitidasE01.PropiedadesAexportar;
                        resultado = proceso.grabarCsv(ficheroSalida, datosProcesados, camposAexportar);
                    }

                    break;

                case "R00":
                    //Facturas recibidas con formato 5 IVAs de diagram (pendiente de desarrollo)

                    break;

                default:
                    //Si no se pasa un tipo de proceso correcto, se graba el fichero de errores.
                    utiles.GrabarFichero(ficheroErrores, $"El tipo de proceso {tipoProceso} no es correcto");
                    break;
            }

            //Grabar el registro de errores si se ha producido alguno
            if(resultado.Length > 0)
            {
                utiles.GrabarFichero(ficheroErrores, resultado.ToString());
            }
        }
    }
}

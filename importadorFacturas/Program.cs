using importadorFacturas.Metodos;
using System.Collections.Generic;
using System.IO;
using System.Text;


namespace importadorFacturas
{
    internal class Program
    {
        public static string ficheroEntrada = string.Empty;
        public static string ficheroColumnas = string.Empty;
        public static string ficheroSalida = string.Empty;
        public static string ficheroErrores = "errores.txt";
        public static string tipoProceso = string.Empty;

        //Instanciacion de las utilidades para acceso a los metodos
        public static Utilidades utiles = new Utilidades();
        static void Main(string[] args)
        {
            if(args.Length == 0)
            {
                return;
            }

            //Se pueden pasar 4 parametros: el primero es el tipo de proceso que se puede usar para futuras importaciones, el segundo es el fichero excel a leer, el tercero es el fichero de salida aunque este es opcional y el cuarto es la configuracion de columnas del excel

            tipoProceso = args[0];

            ficheroEntrada = args[1];
            if(!File.Exists(ficheroEntrada))
            {
                File.WriteAllText(ficheroErrores, "No existe el fichero de entrada");
                return;
            }
            ficheroSalida = Path.ChangeExtension(ficheroEntrada, "csv");
            ficheroErrores = Path.Combine(Path.GetDirectoryName(ficheroEntrada), "errores.txt");
            utiles.ControlFicheros(ficheroErrores);

            int argumentos = args.Length;

            switch(argumentos)
            {
                case 3:
                    ficheroSalida = args[2];
                    utiles.ControlFicheros(ficheroSalida);
                    break;

                case 4:
                    ficheroColumnas = args[3];
                    if(!File.Exists(ficheroColumnas))
                    {
                        File.WriteAllText(ficheroErrores, "No existe el fichero de configuracion de columnas");
                        return;
                    }

                    break;
            }

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
                    resultado = procesoDiagram.EmitidasDiagram(ficheroEntrada);
                    List<Facturas> facturasDiagram = Facturas.obtenerDatos();
                    if(facturasDiagram.Count > 0)
                    {
                        //Array de propiedades a exportar de este tipo
                        string[] camposAexportar = Facturas.ColumnasAexportar;
                        resultado = proceso.grabarCsv(ficheroSalida, facturasDiagram, camposAexportar);
                    }
                    break;

                case "E01":
                    //Facuras emitidas de Alcasal (cliente de Raiña Asesores) tiquet 5863-37

                    //Inicializa campos
                    procesoAlcasal metodo = new procesoAlcasal();

                    //Procesar los datos del fichero de Excel
                    resultado = metodo.emitidasAlcasar(ficheroEntrada);

                    //Carga los datos procesados para pasarlos al csv
                    List<EmitidasE01> facturasAlcasal = EmitidasE01.ObtenerDatos();
                    if(facturasAlcasal.Count > 0)
                    {
                        //Array de propiedades a exportar de este tipo
                        string[] camposAexportar = EmitidasE01.PropiedadesAexportar;
                        resultado = proceso.grabarCsv(ficheroSalida, facturasAlcasal, camposAexportar);
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

using System;
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
        public static int filaInicio = 1;

        //Instanciacion de las utilidades para acceso a los metodos
        public static Metodos.Utilidades utiles = new Metodos.Utilidades();
        static void Main(string[] args)
        {
            if(args.Length == 0)
            {
                return;
            }

            /* Parametros:
             * entrada = fichero de entrada
             * salida = fichero de salida (opcional)
             * columnas = fichero con la configuracion de columnas que tiene el fichero de entrada (solo para el proceso de Diagram)
             * proceso = tipo de proceso a ejecutar
             * fila = fila de la cabecera con los nombres de las columnas. Nota: tiene que haber una fila que tenga todas las columnas rellenas para luego poder procesar bien los datos
             * */

            ProcesarArgumentos(args);

            procesarFichero();
        }



        //Metodo para leer el fichero Excel y procesar los datos segun el tipo pasado por parametro
        private static void procesarFichero()
        {
            //Nota: el tipo de proceso debe ser una letra (E para emitidas y R para recibidas) seguido de dos numeros dejando como reservados para Diagram el '00' (hasta 99 importaciones diferentes de cada tipo)

            //Variable que recoge el texto devuelto en el metodo si se ha producido algun error en el procesado
            StringBuilder resultado = new StringBuilder();

            Procesos proceso = new Procesos();

            switch(tipoProceso)
            {
                //Facturas emitidas con formato diagram
                case "E00":
                    //Inicializa campos
                    Metodos.ProcesoDiagram procesoE00 = new Metodos.ProcesoDiagram();
                    resultado = procesoE00.ProcesarFacturas(ficheroEntrada);
                    List<Facturas> facturasE00 = Facturas.ObtenerFacturas();
                    if(facturasE00.Count > 0)
                    {
                        //Array de propiedades a exportar de este tipo
                        string[] camposAexportar = Facturas.ColumnasAexportar;
                        resultado = proceso.GrabarCsv(ficheroSalida, facturasE00, camposAexportar);
                    }
                    break;

                //Facuras emitidas de Alcasal (cliente de Raiña Asesores) tiquet 5863-37
                case "E01":
                    //Inicializa campos
                    procesoAlcasal procesoE01 = new procesoAlcasal();

                    //Procesar los datos del fichero de Excel
                    resultado = procesoE01.EmitidasAlcasar(ficheroEntrada);

                    //Carga los datos procesados para pasarlos al csv
                    List<EmitidasE01> facturasE01 = EmitidasE01.ObtenerFacturasE01();
                    if(facturasE01.Count > 0)
                    {
                        //Array de propiedades a exportar de este tipo
                        string[] camposAexportar = Facturas.ColumnasAexportar;//EmitidasE01.PropiedadesAexportar;
                        resultado = proceso.GrabarCsv(ficheroSalida, facturasE01, camposAexportar);
                    }

                    break;

                //Facturas recibidas con formato diagram
                case "R00":
                    //Inicializa campos
                    Metodos.ProcesoDiagram procesoR00 = new Metodos.ProcesoDiagram();
                    resultado = procesoR00.ProcesarFacturas(ficheroEntrada);
                    List<Facturas> facturasR00 = Facturas.ObtenerFacturas();
                    if(facturasR00.Count > 0)
                    {
                        //Array de propiedades a exportar de este tipo
                        string[] camposAexportar = Facturas.ColumnasAexportar;
                        resultado = proceso.GrabarCsv(ficheroSalida, facturasR00, camposAexportar);
                    }
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

        private static void ProcesarArgumentos(string[] argumentos)
        {
            foreach(string argumento in argumentos)
            {
                var partes = argumento.Split('=');
                string nombre = string.Empty;
                string valor = string.Empty;

                if(partes.Length == 2)
                {
                    nombre = partes[0];
                    valor = partes[1];
                }

                switch(nombre)
                {
                    case "entrada":
                        ficheroEntrada = valor;
                        if(!File.Exists(ficheroEntrada))
                        {
                            File.WriteAllText(ficheroErrores, "No existe el fichero de entrada");
                            return;
                        }

                        if(string.IsNullOrEmpty(ficheroSalida)) ficheroSalida = $"salida_{Path.GetFileNameWithoutExtension(ficheroEntrada)}.csv";
                        ficheroErrores = Path.Combine(Path.GetDirectoryName(ficheroEntrada), "errores.txt");
                        utiles.ControlFicheros(ficheroErrores);

                        break;

                    case "salida":
                        ficheroSalida = valor;
                        utiles.ControlFicheros(ficheroSalida);
                        break;

                    case "columnas":
                        ficheroColumnas = valor;
                        if(!File.Exists(ficheroColumnas))
                        {
                            File.WriteAllText(ficheroErrores, "No existe el fichero de configuracion de columnas");
                            return;
                        }
                        break;

                    case "proceso":
                        tipoProceso = valor;
                        break;

                    case "fila":
                        filaInicio = Convert.ToInt32(valor);
                        break;

                }
            }
        }
    }
}

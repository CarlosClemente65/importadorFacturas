using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using DocumentFormat.OpenXml.Bibliography;
using UtilesDiagram;


namespace importadorFacturas
{
    internal class Program
    {
        public static Configuracion Parametros = new Configuracion();

        public static Configuracion.TiposProceso TiposProceso { get; private set; }

        //Instanciacion de las utilidades para acceso a los metodos
        public static UtilidadesDiagram utiles = new UtilidadesDiagram();

        public static Procesos proceso = new Procesos();

        static void Main(string[] args)
        {
            if(args.Length == 0)
            {
                return;
            }

            /* Parametros:
             * entrada = fichero de entrada
             * salida = fichero de salida (opcional)
             * configuracion = fichero con la configuracion de columnas que tiene el fichero de entrada (solo para el proceso de Diagram)
             * proceso = tipo de proceso a ejecutar
             * fila = fila de la cabecera con los nombres de las columnas. Nota: tiene que haber una fila que tenga todas las columnas rellenas para luego poder procesar bien los datos
             * */

            if(!ProcesarArgumentos(args))
            {
                procesarFichero(Parametros);
            }
        }



        //Metodo para leer el fichero Excel y procesar los datos segun el tipo pasado por parametro
        private static void procesarFichero(Configuracion parametros)
        {
            //Nota: el tipo de proceso debe ser una letra (E para emitidas y R para recibidas) seguido de dos numeros dejando como reservados para Diagram el '00' (hasta 99 importaciones diferentes de cada tipo)

            //Variable que recoge el texto devuelto en el metodo si se ha producido algun error en el procesado
            StringBuilder resultado = new StringBuilder();

            Procesos proceso = new Procesos();

            switch(Parametros.TipoProceso)
            {
                //Facturas emitidas con formato diagram
                case "E00":
                    //Inicializa campos
                    Metodos.ProcesoDiagram procesoE00 = new Metodos.ProcesoDiagram();
                    resultado = procesoE00.ProcesarFacturas(parametros);
                    List<Facturas> facturasE00 = Facturas.ObtenerFacturas();
                    if(facturasE00.Count > 0)
                    {
                        //Array de propiedades a exportar de este tipo
                        string[] camposAexportar = Facturas.ColumnasAexportar;
                        resultado = proceso.GrabarCsv(Parametros.FicheroSalida, facturasE00, camposAexportar);
                    }
                    break;

                //Facuras emitidas de Alcasal (cliente de Raiña Asesores) tiquet 5863-37
                case "E01":
                    //Inicializa campos
                    ProcesoAlcasal procesoE01 = new ProcesoAlcasal();

                    //Procesar los datos del fichero de Excel
                    resultado = procesoE01.EmitidasAlcasar(Parametros);

                    //Carga los datos procesados para pasarlos al csv
                    List<EmitidasE01> facturasE01 = EmitidasE01.ObtenerFacturasE01();
                    if(facturasE01.Count > 0)
                    {
                        //Array de propiedades a exportar de este tipo
                        string[] camposAexportar = Facturas.ColumnasAexportar;
                        resultado = proceso.GrabarCsv(Parametros.FicheroSalida, facturasE01, camposAexportar);
                    }

                    break;

                //Facturas recibidas con formato diagram
                case "R00":
                    //Inicializa campos
                    Metodos.ProcesoDiagram procesoR00 = new Metodos.ProcesoDiagram();
                    resultado = procesoR00.ProcesarFacturas(parametros);
                    List<Facturas> facturasR00 = Facturas.ObtenerFacturas();
                    if(facturasR00.Count > 0)
                    {
                        //Array de propiedades a exportar de este tipo
                        string[] camposAexportar = Facturas.ColumnasAexportar;
                        resultado = proceso.GrabarCsv(Parametros.FicheroSalida, facturasR00, camposAexportar);
                    }
                    break;

                default:
                    //Si no se pasa un tipo de proceso correcto, se graba el fichero de errores.
                    utiles.GrabarFichero(Parametros.FicheroErrores, $"El tipo de proceso {Parametros.TipoProceso} no es correcto");
                    break;
            }

            //Grabar el registro de errores si se ha producido alguno
            if(resultado.Length > 0)
            {
                utiles.GrabarFichero(Parametros.FicheroErrores, resultado.ToString());
            }
        }

        private static bool ProcesarArgumentos(string[] argumentos)
        {
            bool errores = false;
            foreach(string argumento in argumentos)
            {
                //Separa el argumento y su valor
                (string nombre, string valor) = utiles.DivideCadena(argumento, '=');

                string chequeo = string.Empty;
                switch(nombre)
                {
                    //Fichero entrada
                    case "entrada":
                        chequeo = ChequeoFichero(valor);

                        if(!string.IsNullOrEmpty(chequeo))
                        {
                            utiles.GrabarFichero(Parametros.FicheroErrores, chequeo);
                            return true;
                        }
                        Parametros.FicheroEntrada = valor;

                        if(string.IsNullOrEmpty(Parametros.FicheroSalida))
                        {
                            Parametros.FicheroSalida = $"salida_{Path.GetFileNameWithoutExtension(Parametros.FicheroEntrada)}.csv";
                            utiles.ControlFicheros(Parametros.FicheroSalida);
                        }

                        Parametros.FicheroErrores = Path.Combine(Path.GetDirectoryName(Parametros.FicheroEntrada), "errores.txt");
                        utiles.ControlFicheros(Parametros.FicheroErrores);

                        break;

                    //Fichero salida
                    case "salida":
                        Parametros.FicheroSalida = valor;
                        utiles.ControlFicheros(Parametros.FicheroSalida);
                        break;

                    //Fichero configuracion
                    case "configuracion":
                        chequeo = ChequeoFichero(valor);

                        if(!string.IsNullOrEmpty(chequeo))
                        {
                            utiles.GrabarFichero(Parametros.FicheroErrores, chequeo);
                            return true;
                        }
                        Parametros.FicheroConfiguracion = valor;
                        break;

                    //Tipo de proceso
                    case "proceso":
                        //Se valida que el tipo de proceso sea alguno de los definidos en la clase 'Configuracion'
                        if(Enum.TryParse<Configuracion.TiposProceso>(valor, out Configuracion.TiposProceso _tipoProceso))
                        {
                            Parametros.TipoProceso = _tipoProceso.ToString();
                        }
                        else
                        {
                            utiles.GrabarFichero(Parametros.FicheroErrores, $"Error. Tipo de proceso {valor} incorrecto");
                            return true;
                        }
                        break;

                    //Fila de cabecera de columnas
                    case "fila":
                        //Se valida que la fila sea un numero mayor que 0
                        if(int.TryParse(valor, out int _fila) && _fila > 0)
                        {
                            Parametros.FilaInicio = _fila;
                        }
                        else
                        {
                            utiles.GrabarFichero(Parametros.FicheroErrores, $"Error. Fila {valor} incorrecta");
                            return true;
                        }
                        break;

                    //Hoja en la que estan los datos (opcional)
                    case "hoja":
                        //Se valida que la hoja sea un numero mayor que 0
                        if(int.TryParse(valor, out int _hoja) && _hoja > 0)
                        {
                            Parametros.HojaExcel = _hoja;
                        }
                        else
                        {
                            utiles.GrabarFichero(Parametros.FicheroErrores, $"Error. Hoja {valor} incorrecta");
                            return true;
                        }
                        break;
                }
            }
            return errores;
        }

        private static string ChequeoFichero(string fichero)
        {
            string resultado = string.Empty;
            if(!File.Exists(fichero))
            {
                resultado = $"Error. No existe el fichero {fichero}";
            }
            return resultado;
        }
    }
}

using System.Collections.Generic;
using System.IO;
using System.Text;
using UtilidadesDiagram;


namespace importadorFacturas
{
    internal class Program
    {
        public static Configuracion.TiposProceso TiposProceso { get; private set; }

        public static Procesos proceso = new Procesos();

        static void Main(string[] args)
        {
            //Controla que se pase como argumento el guion
            if(args.Length == 0)
            {
                return;
            }

            string ficheroGuion = args[0];
            
            //Controla que exista el fichero con el guion
            if(!File.Exists(ficheroGuion))
            {
                Utilidades.GrabarFichero(Configuracion.FicheroErrores, $"Error. No existe el fichero {ficheroGuion}");
                return;
            }

            //Si no se han producido errores al cargar el guion, se procesa el fichero.
            if(proceso.CargarGuion(ficheroGuion))
            {
                procesarFichero();
            }
        }



        //Metodo para leer el fichero Excel y procesar los datos segun el tipo pasado por parametro
        private static void procesarFichero()
        {
            //Nota: el tipo de proceso debe ser una letra (E para emitidas y R para recibidas) seguido de dos numeros dejando como reservados para Diagram el '00' (hasta 99 importaciones diferentes de cada tipo)

            //Variable que recoge el texto devuelto en el metodo si se ha producido algun error en el procesado
            StringBuilder resultado = new StringBuilder();

            switch(Configuracion.TipoProceso)
            {
                //Facturas emitidas con formato diagram
                case "E00":
                    //Inicializa campos
                    Metodos.ProcesoDiagram procesoE00 = new Metodos.ProcesoDiagram();
                    resultado = procesoE00.ProcesarFacturas();
                    List<Facturas> facturasE00 = Facturas.ObtenerFacturas();
                    if(facturasE00.Count > 0)
                    {
                        resultado = proceso.GrabarCsv(facturasE00, Facturas.ColumnasAexportar.ToArray());
                    }
                    break;

                //Facuras emitidas de Alcasal (cliente de Raiña Asesores) tiquet 5863-37
                case "E01":
                    //Inicializa campos
                    ProcesoAlcasal procesoE01 = new ProcesoAlcasal();

                    //Procesar los datos del fichero de Excel
                    resultado = procesoE01.EmitidasAlcasar();

                    //Carga los datos procesados para pasarlos al csv
                    List<EmitidasE01> facturasE01 = EmitidasE01.ObtenerFacturasE01();
                    if(facturasE01.Count > 0)
                    {
                        resultado = proceso.GrabarCsv(facturasE01, Facturas.ColumnasAexportar.ToArray());
                    }

                    break;

                //Facturas recibidas con formato diagram
                case "R00":
                    //Inicializa campos
                    Metodos.ProcesoDiagram procesoR00 = new Metodos.ProcesoDiagram();
                    resultado = procesoR00.ProcesarFacturas();
                    List<Facturas> facturasR00 = Facturas.ObtenerFacturas();
                    if(facturasR00.Count > 0)
                    {
                        resultado = proceso.GrabarCsv(facturasR00, Facturas.ColumnasAexportar.ToArray());
                    }
                    break;

                default:
                    //Si no se pasa un tipo de proceso correcto, se graba el fichero de errores.
                    Utilidades.GrabarFichero(Configuracion.FicheroErrores, $"El tipo de proceso {Configuracion.TipoProceso} no es correcto");
                    break;
            }

            //Grabar el registro de errores si se ha producido alguno
            if(resultado.Length > 0)
            {
                Utilidades.GrabarFichero(Configuracion.FicheroErrores, resultado.ToString());
            }
        }
    }
}

using System.Collections.Generic;

namespace importadorFacturas
{
    public static class Configuracion
    {
        public static string FicheroEntrada {  get; set; }
        public static string FicheroSalida { get; set; }
        //public static string FicheroConfiguracion { get; set; }
        public static string FicheroErrores { get; set; } = "errores.txt";
        public static string TipoProceso { get; set; }
        public static int FilaInicio { get; set; } = 1;
        public static int HojaExcel { get; set; } = 1;

        public static List<string> parametros = new List<string>();
        public static List<string> columnas = new List<string>();

        //Detalle de los tipos de proceso validos que estan implementados.
        public enum TiposProceso
        {
            E00, //Emitidas de Diagram
            E01, //Emitidas de Alcasal
            R00 //Recibidas de Diagram
        }
    }
}

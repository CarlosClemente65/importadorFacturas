using System;
using System.Collections.Generic;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace importadorFacturas
{
    public class Configuracion
    {
        public string FicheroEntrada {  get; set; }
        public string FicheroSalida { get; set; }
        public string FicheroConfiguracion { get; set; }
        public string FicheroErrores { get; set; } = "errores.txt";
        public string TipoProceso { get; set; }
        public int FilaInicio { get; set; } = 1;
        public int HojaExcel { get; set; } = 1;

        //Detalle de los tipos de proceso validos que estan implementados.
        public enum TiposProceso
        {
            E00, //Emitidas de Diagram
            E01, //Emitidas de Alcasal
            R00 //Recibidas de Diagram
        }

    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Permissions;
using System.Text;
using System.Threading.Tasks;

namespace importadorFacturas
{
    public class ingresosAlcasar
    {
        //Clase que representa la estructura de datos que finalmente se generara en el fichero de salida
        public string tipoFactura { get; set; }
        public string fechaFactura { get; set; }
        public string serieFactura { get; set; }
        public string numeroFactura { get; set; }
        public string nifFactura { get; set; }
        public string nombreFactura { get; set; }
        public string apellidoFactura { get; set; }
        public string direccionFactura { get; set; }
        public string codPostalFactura { get; set; }
        public float baseFactura { get; set; }
        public float porcentajeIva { get; set; }
        public float cuotaIva { get; set; }
        public float totalFactura { get; set; }
        public string fechaFraAgrupada { get; set; }
        public string primerNumero { get; set; }
        public string ultimoNumero { get; set; }
        public int contadorFacturas { get; set; }

        public static List<ingresosAlcasar> ListaIngresos { get; set; } = new List<ingresosAlcasar>();

        public static List<ingresosAlcasar> obtenerDatos()
        {
            return ListaIngresos;
        }

    }
}

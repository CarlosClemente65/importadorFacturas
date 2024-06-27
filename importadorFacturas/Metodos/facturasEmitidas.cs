using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Permissions;
using System.Text;
using System.Threading.Tasks;

namespace importadorFacturas
{
    public class facturasEmitidas
    {
        //Clase que representa la estructura de datos que finalmente se generara en el fichero de salida
        public string fechaFactura { get; set; }
        public string serieFactura { get; set; }
        public string numeroFactura { get; set; }
        public string referenciaFactura { get; set; }
        public decimal baseFactura { get; set; }
        public float porcentajeIva { get; set; }
        public decimal cuotaIva { get; set; }
        public float porcentajeRecargo { get; set; }
        public decimal cuotaRecargo { get; set; }
        public decimal baseRetencion { get; set; }
        public float porcentajeRetencion { get; set; }
        public decimal cuotaRetencion { get; set; }
        public decimal totalFactura { get; set; }
        public string primerNumero { get; set; }
        public string ultimoNumero { get; set; }
        public int contadorFacturas { get; set; }
        public string nifFactura { get; set; }
        public string apellidoFactura { get; set; }
        public string nombreFactura { get; set; }
        public string direccionFactura { get; set; }
        public string codPostalFactura { get; set; }

        public static List<facturasEmitidas> ListaIngresos { get; set; } = new List<facturasEmitidas>();

        public static List<facturasEmitidas> obtenerDatos()
        {
            return ListaIngresos;
        }

    }
}

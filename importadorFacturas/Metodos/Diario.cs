using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace importadorFacturas.Metodos
{
    // Clase que representa la estructura de datos que finalmente se generara en el fichero de salida para crear los asientos en la importacion de balances.
    public class Diario
    {
        // Propiedades del diario que se generaran en cada columna
        public int Apunte { get; set; }
        public string Cuenta { get; set; }
        public char Signo { get; set; }
        public decimal Importe { get; set; }
        public string CuentaDebe {  get; set; }
        public string CuentaHaber { get; set; }


        // Lista que almacena los apuntes generados
        public static List<Facturas> ApuntesDiario { get; set; }


        // Diccionario con el numero de columna y el nombre del contenido para luego asignarlo a la clase
        public static Dictionary<int, string> MapeoColumnasDiario;

        public static void MapeoDiario()
        {
            //Asigna a cada columna la propiedad que le corresponde para luego generar la salida (es el defecto que tendra la salida en csv
            MapeoColumnasDiario = new Dictionary<int, string>
            {
                { 1, "Cuenta" },
                { 2, "Descripcion" },
                { 3, "ImporteDebe" },
                { 4, "ImoprteHaber" },
                { 5, "Signo" },
            };
        }


        // Devuelve la lista de asientos una vez procesados
        public static List<Facturas> ObtenerDiario()
        {
            return ApuntesDiario;
        }


    }


    // Clase que representa la estructura de datos que finalmente se generara en el fichero de salida para dar de alta las cuentas en la importacion de balances para generar un asiento. 
    public class Cuentas
    {
        public string Cuenta { get; set; }

        public string Descripcion { get; set; }


        // Lista que almacena la relacion de cuentas
        public static List<Cuentas> RelacionCuentas { get; set; }

        // Devuelve la lista de cuentas una vez procesadas
        public static List<Cuentas> ObtenerCuentas()
        {
            return RelacionCuentas;
        }
    }
}

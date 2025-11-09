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
        public string Cuenta { get; set; } = string.Empty;
        public string Descripcion { get; set; } = string.Empty;
        public char Signo { get; set; }
        public decimal Importe { get; set; }
        public string CuentaDebe {  get; set; } = string.Empty;
        public string CuentaHaber { get; set; } = string.Empty;
        public decimal ImporteDebe { get; set; }
        public decimal ImporteHaber { get; set; }


        // Lista que almacena los apuntes generados
        public static List<Diario> ApuntesDiario { get; set; }


        // Diccionario con el numero de columna y el nombre del contenido para luego asignarlo a la clase
        public static Dictionary<int, string> MapeoColumnasDiario;

        public static List<string> ColumnasAexportar { get; set; }

        public static void MapeoDiario()
        {
            //Asigna a cada columna la propiedad que le corresponde para luego generar la salida (es el defecto que tendra la salida en csv
            MapeoColumnasDiario = new Dictionary<int, string>
            {
                { 1, "Cuenta" },
                { 2, "Descripcion" },
                { 3, "ImporteDebe" },
                { 4, "ImporteHaber" },
                { 5, "Signo" },
            };
        }


        // Devuelve la lista de asientos una vez procesados
        public static List<Diario> ObtenerDiario()
        {
            return ApuntesDiario;
        }


    }
}

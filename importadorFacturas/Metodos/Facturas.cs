using System;
using System.Collections.Generic;
using System.Linq;

namespace importadorFacturas
{
    //Clase que representa la estructura de datos que finalmente se generara en el fichero de salida. El atributo indica el orden en el que se incluira en el fichero de salida, y van de 10 en 10 para permitir insertar nuevos atributos si fuera necesario en clases especiales como la de Alcasal (tipo E01)
    public class Facturas
    {
        //Campos obligatorios
        [OrdenCsv(10)]
        public int contador { get; set; }

        [OrdenCsv(20)]
        public string fechaFactura { get; set; }

        [OrdenCsv(30)]
        public string fechaOperacion { get; set; }

        [OrdenCsv(40)]
        public string periodoFactura { get; set; }

        [OrdenCsv(50)]
        public string serieFactura { get; set; }

        [OrdenCsv(60)]
        public string numeroFactura { get; set; }

        [OrdenCsv(70)]
        public string referenciaFactura { get; set; }

        [OrdenCsv(80)]
        public string nifFactura { get; set; }

        [OrdenCsv(90)]
        public string nombreFactura { get; set; }

        [OrdenCsv(100)]
        public string direccionFactura { get; set; }

        [OrdenCsv(110)]
        public string codPostalFactura { get; set; }

        [OrdenCsv(120)]
        public string paisFactura { get; set; }

        [OrdenCsv(130)]
        public string cuentaContable { get; set; }

        [OrdenCsv(140)]
        public string cuentaContrapartida { get; set; }

        [OrdenCsv(150)]
        public string codigoConcepto { get; set; }

        [OrdenCsv(160)]
        public char facturaDeducible { get; set; }

        [OrdenCsv(170)]
        public decimal baseFactura1 { get; set; } //Base al 21%

        [OrdenCsv(180)]
        public float porcentajeIva1 { get; set; } //Porcentaje fijo del 21% (en el constructor)

        [OrdenCsv(190)]
        public decimal cuotaIva1 { get; set; }

        [OrdenCsv(200)]
        public float porcentajeRecargo1 { get; set; } //Porcentaje fijo del 5,2% (en el constructor)

        [OrdenCsv(210)]
        public decimal cuotaRecargo1 { get; set; }

        [OrdenCsv(220)]
        public decimal baseFactura2 { get; set; } //Base al 10%

        [OrdenCsv(230)]
        public float porcentajeIva2 { get; set; }//Porcentaje fijo del 10% (en el constructor)

        [OrdenCsv(240)]
        public decimal cuotaIva2 { get; set; }

        [OrdenCsv(250)]
        public float porcentajeRecargo2 { get; set; }//Porcentaje fijo del 1,4% (en el constructor)

        [OrdenCsv(260)]
        public decimal cuotaRecargo2 { get; set; }

        [OrdenCsv(270)]
        public decimal baseFactura3 { get; set; } //Base al 4%

        [OrdenCsv(280)]
        public float porcentajeIva3 { get; set; }//Porcentaje fijo del 4% (en el constructor)

        [OrdenCsv(290)]
        public decimal cuotaIva3 { get; set; }

        [OrdenCsv(300)]
        public float porcentajeRecargo3 { get; set; }//Porcentaje fijo del 0,5% (en el constructor)

        [OrdenCsv(310)]
        public decimal cuotaRecargo3 { get; set; }

        [OrdenCsv(320)]
        public decimal baseFactura4 { get; set; }//Base exenta

        [OrdenCsv(330)]
        public float porcentajeIva4 { get; set; }//Porcentaje fijo del 0% (en el constructor)

        [OrdenCsv(340)]
        public decimal cuotaIva4 { get; set; } //Se fija a cero por ser base exenta (en el constructor)

        [OrdenCsv(350)]
        public float porcentajeRecargo4 { get; set; }//Porcentaje fijo del 0% (en el constructor)

        [OrdenCsv(360)]
        public decimal cuotaRecargo4 { get; set; }//Se fija a cero por ser base exenta (en el constructor)

        [OrdenCsv(370)]
        public decimal baseFactura5 { get; set; }

        [OrdenCsv(380)]
        public float porcentajeIva5 { get; set; }

        [OrdenCsv(390)]
        public decimal cuotaIva5 { get; set; }

        [OrdenCsv(400)]
        public float porcentajeRecargo5 { get; set; }

        [OrdenCsv(410)]
        public decimal cuotaRecargo5 { get; set; }

        [OrdenCsv(420)]
        public decimal baseFactura6 { get; set; }

        [OrdenCsv(430)]
        public float porcentajeIva6 { get; set; }

        [OrdenCsv(440)]
        public decimal cuotaIva6 { get; set; }

        [OrdenCsv(450)]
        public float porcentajeRecargo6 { get; set; }

        [OrdenCsv(460)]
        public decimal cuotaRecargo6 { get; set; }

        [OrdenCsv(470)]
        public decimal baseFactura7 { get; set; }

        [OrdenCsv(480)]
        public float porcentajeIva7 { get; set; }

        [OrdenCsv(490)]
        public decimal cuotaIva7 { get; set; }

        [OrdenCsv(500)]
        public float porcentajeRecargo7 { get; set; }

        [OrdenCsv(510)]
        public decimal cuotaRecargo7 { get; set; }

        [OrdenCsv(520)]
        public decimal baseFactura8 { get; set; }

        [OrdenCsv(530)]
        public float porcentajeIva8 { get; set; }

        [OrdenCsv(540)]
        public decimal cuotaIva8 { get; set; }

        [OrdenCsv(550)]
        public float porcentajeRecargo8 { get; set; }

        [OrdenCsv(560)]
        public decimal cuotaRecargo8 { get; set; }

        [OrdenCsv(570)]
        public decimal baseFactura9 { get; set; }

        [OrdenCsv(580)]
        public float porcentajeIva9 { get; set; }

        [OrdenCsv(590)]
        public decimal cuotaIva9 { get; set; }

        [OrdenCsv(600)]
        public float porcentajeRecargo9 { get; set; }

        [OrdenCsv(610)]
        public decimal cuotaRecargo9 { get; set; }

        [OrdenCsv(620)]
        public decimal baseIrpf { get; set; }

        [OrdenCsv(630)]
        public float porcentajeIrpf { get; set; }

        [OrdenCsv(640)]
        public decimal cuotaIrpf { get; set; }

        [OrdenCsv(650)]
        public decimal totalFactura { get; set; }

        public static List<Facturas> ListaFacturas { get; set; }

        public static Dictionary<int, string> mapeoColumnas;

        public static string[] ColumnasAexportar { get; set; }

        
        //Constructor de la clase con los defectos de los campos
        public Facturas()
        {
            porcentajeIva1 = 21.0f;
            porcentajeIva2 = 10.0f;
            porcentajeIva3 = 4.0f;
            porcentajeIva4 = 0.0f;
            cuotaIva4 = 0.0M;
            porcentajeRecargo1 = 5.20f;
            porcentajeRecargo2 = 1.40f;
            porcentajeRecargo3 = 0.50f;
            porcentajeRecargo4 = 0.0f;
            cuotaRecargo4 = 0.0M;
            facturaDeducible = 'S';
            paisFactura = "ES";
        }


        //Metodo para obtener la lista de facturas procesadas
        public static List<Facturas> ObtenerFacturas()
        {
            return ListaFacturas;
        }

        //Metodo para mapear las columnas con sus nombres de propiedad
        public static void MapeoFacturas()
        {
            //Asigna a cada columna la propiedad que le corresponde (estan en el mismo orden que la plantilla de Excel con los campos)
            mapeoColumnas = new Dictionary<int, string>
            {
                {1, "fechaFactura" },
                {2, "fechaOperacion" },
                {3, "periodoFactura" },
                {4, "serieFactura" },
                {5, "numeroFactura" },
                {6, "referenciaFactura" },
                {7, "nifFactura" },
                {8, "nombreFactura" },
                {9, "direccionFactura" },
                {10, "codPostalFactura" },
                {11, "paisFactura" },
                {12, "cuentaContable" },
                {13, "cuentaContrapartida" },
                {14, "codigoConcepto" },
                {15, "facturaDeducible" },
                {16, "baseFactura1" },
                {17, "porcentajeIva1" },
                {18, "cuotaIva1" },
                {19, "porcentajeRecargo1" },
                {20, "cuotaRecargo1" },
                {21, "baseFactura2" },
                {22, "porcentajeIva2" },
                {23, "cuotaIva2" },
                {24, "porcentajeRecargo2" },
                {25, "cuotaRecargo2" },
                {26, "baseFactura3" },
                {27, "porcentajeIva3" },
                {28, "cuotaIva3" },
                {29, "porcentajeRecargo3" },
                {30, "cuotaRecargo3" },
                {31, "baseFactura4" },
                {32, "porcentajeIva4" },
                {33, "cuotaIva4" },
                {34, "porcentajeRecargo4" },
                {35, "cuotaRecargo4" },
                {36, "baseFactura5" },
                {37, "porcentajeIva5" },
                {38, "cuotaIva5" },
                {39, "porcentajeRecargo5" },
                {40, "cuotaRecargo5" },
                {41, "baseFactura6" },
                {42, "porcentajeIva6"},
                {43, "cuotaIva6" },
                {44, "porcentajeRecargo6" },
                {45, "cuotaRecargo6" },
                {46, "baseFactura7" },
                {47, "porcentajeIva7" },
                {48, "cuotaIva7" },
                {49, "porcentajeRecargo7" },
                {50, "cuotaRecargo7" },
                {51, "baseFactura8" },
                {52, "porcentajeIva8" },
                {53, "cuotaIva8" },
                {54, "porcentajeRecargo8" },
                {55, "cuotaRecargo8" },
                {56, "baseFactura9" },
                {57, "porcentajeIva9" },
                {58, "cuotaIva9" },
                {59, "porcentajeRecargo9" },
                {60, "cuotaRecargo9" },
                {61, "baseIrpf" },
                {62, "porcentajeIrpf" },
                {63, "cuotaIrpf" },
                {64, "totalFactura" }
            };

            //Se añade el campo 'contador' para que se incluya en el fichero de salida
            ColumnasAexportar = new List<string> { "contador" }.Concat(mapeoColumnas.Values).ToArray();
        }


        //Personalizacion de los atributos de las propiedades para poner el numero de orden en el que luego incluirlos en el fichero de salida
        [AttributeUsage(AttributeTargets.Property)]
        public class OrdenCsvAttribute : Attribute
        {
            public int Orden { get; }

            public OrdenCsvAttribute(int orden)
            {
                Orden = orden;
            }
        }

    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace importadorFacturas
{
    public class Facturas
    {
        //Clase que representa la estructura de datos que finalmente se generara en el fichero de salida. El atributo indica el orden en el que se incluira en el fichero de salida, y van de 10 en 10 para permitir insertar nuevos atributos si fuera necesario en clases especiales como la de Alcasal (tipo E01)

        [OrdenCsv(10)]
        public int contador { get; set; }

        [OrdenCsv(20)]
        public string fechaFactura { get; set; }

        [OrdenCsv(30)]
        public string serieFactura { get; set; }

        [OrdenCsv(40)]
        public string numeroFactura { get; set; }

        [OrdenCsv(50)]
        public int lineaFactura { get; set; }

        [OrdenCsv(60)]
        public string referenciaFactura { get; set; }

        [OrdenCsv(70)]
        public decimal baseFactura1 { get; set; }

        [OrdenCsv(80)]
        public float porcentajeIva1 { get; set; }

        [OrdenCsv(90)]
        public decimal cuotaIva1 { get; set; }

        [OrdenCsv(100)]
        public float porcentajeRecargo1 { get; set; }

        [OrdenCsv(110)]
        public decimal cuotaRecargo1 { get; set; }

        [OrdenCsv(120)]
        public decimal baseFactura2 { get; set; }

        [OrdenCsv(130)]
        public float porcentajeIva2 { get; set; }

        [OrdenCsv(140)]
        public decimal cuotaIva2 { get; set; }

        [OrdenCsv(150)]
        public float porcentajeRecargo2 { get; set; }

        [OrdenCsv(160)]
        public decimal cuotaRecargo2 { get; set; }

        [OrdenCsv(170)]
        public decimal baseFactura3 { get; set; }

        [OrdenCsv(180)]
        public float porcentajeIva3 { get; set; }

        [OrdenCsv(190)]
        public decimal cuotaIva3 { get; set; }

        [OrdenCsv(200)]
        public float porcentajeRecargo3 { get; set; }

        [OrdenCsv(210)]
        public decimal cuotaRecargo3 { get; set; }

        [OrdenCsv(220)]
        public decimal baseFactura4 { get; set; }

        [OrdenCsv(230)]
        public float porcentajeIva4 { get; set; }

        [OrdenCsv(240)]
        public decimal cuotaIva4 { get; set; }

        [OrdenCsv(250)]
        public float porcentajeRecargo4 { get; set; }

        [OrdenCsv(260)]
        public decimal cuotaRecargo4 { get; set; }

        [OrdenCsv(270)]
        public decimal baseIrpf { get; set; }

        [OrdenCsv(280)]
        public float porcentajeIrpf { get; set; }

        [OrdenCsv(290)]
        public decimal cuotaIrpf { get; set; }

        [OrdenCsv(300)]
        public decimal totalFactura { get; set; }

        [OrdenCsv(310)]
        public string nifFactura { get; set; }

        [OrdenCsv(320)]
        public string apellidoFactura { get; set; }

        [OrdenCsv(330)]
        public string nombreFactura { get; set; }


        //Campos opcionales
        [OrdenCsv(340)]
        public int periodoFactura { get; set; }

        [OrdenCsv(350)]
        public string cuentaContable { get; set; }

        [OrdenCsv(360)]
        public string cuentaContrapartida { get; set; }

        [OrdenCsv(370)]
        public string codigoConcepto { get; set; }

        [OrdenCsv(380)]
        public string fechaOperacion { get; set; }

        [OrdenCsv(390)]
        public char facturaDeducible { get; set; }

        [OrdenCsv(400)]
        public string direccionFactura { get; set; }

        [OrdenCsv(410)]
        public string codPostalFactura { get; set; }

        [OrdenCsv(420)]
        public string paisFactura { get; set; }

        [OrdenCsv(430)]
        public decimal baseFactura5 { get; set; }

        [OrdenCsv(440)]
        public float porcentajeIva5 { get; set; }

        [OrdenCsv(450)]
        public decimal cuotaIva5 { get; set; }

        [OrdenCsv(460)]
        public float porcentajeRecargo5 { get; set; }

        [OrdenCsv(470)]
        public decimal cuotaRecargo5 { get; set; }

        [OrdenCsv(480)]
        public decimal baseFactura6 { get; set; }

        [OrdenCsv(490)]
        public float porcentajeIva6 { get; set; }

        [OrdenCsv(500)]
        public decimal cuotaIva6 { get; set; }

        [OrdenCsv(510)]
        public float porcentajeRecargo6 { get; set; }

        [OrdenCsv(520)]
        public decimal cuotaRecargo6 { get; set; }

        [OrdenCsv(530)]
        public decimal baseFactura7 { get; set; }

        [OrdenCsv(540)]
        public float porcentajeIva7 { get; set; }

        [OrdenCsv(550)]
        public decimal cuotaIva7 { get; set; }

        [OrdenCsv(560)]
        public float porcentajeRecargo7 { get; set; }

        [OrdenCsv(570)]
        public decimal cuotaRecargo7 { get; set; }

        [OrdenCsv(580)]
        public decimal baseFactura8 { get; set; }

        [OrdenCsv(590)]
        public float porcentajeIva8 { get; set; }

        [OrdenCsv(600)]
        public decimal cuotaIva8 { get; set; }

        [OrdenCsv(610)]
        public float porcentajeRecargo8 { get; set; }

        [OrdenCsv(620)]
        public decimal cuotaRecargo8 { get; set; }

        [OrdenCsv(630)]
        public decimal baseFactura9 { get; set; }

        [OrdenCsv(640)]
        public float porcentajeIva9 { get; set; }

        [OrdenCsv(650)]
        public decimal cuotaIva9 { get; set; }

        [OrdenCsv(660)]
        public float porcentajeRecargo9 { get; set; }

        [OrdenCsv(670)]
        public decimal cuotaRecargo9 { get; set; }

        public static List<Facturas> ListaFacturas { get; set; }

        public static Dictionary<int, string> mapeoColumnas;

        public Facturas()
        {
            //Constructor de la clase que crea una nueva lista de facturas y asigna los defectos de varios campos
            ListaFacturas = new List<Facturas>();

            //Defectos de los campos
            facturaDeducible = 'S';
            paisFactura = "ES";

            //Asigna a cada columna la propiedad que le corresponde (estan en el mismo orden que la plantilla de Excel con los campos)
            mapeoColumnas = new Dictionary<int, string>
            {
                {1, "fechaFactura" },
                {2, "serieFactura" },
                {3, "numeroFactura" },
                {4, "referenciaFactura" },
                {5, "baseFactura1" },
                {6, "porcentajeIva1" },
                {7, "cuotaIva1" },
                {8, "porcentajeRecargo1" },
                {9, "cuotaRecargo1" },
                {10, "baseFactura2" },
                {11, "porcentajeIva2" },
                {12, "cuotaIva2" },
                {13, "porcentajeRecargo2" },
                {14, "cuotaRecargo2" },
                {15, "baseFactura3" },
                {16, "porcentajeIva3" },
                {17, "cuotaIva3" },
                {18, "porcentajeRecargo3" },
                {19, "cuotaRecargo3" },
                {20, "baseFactura4" },
                {21, "porcentajeIva4" },
                {22, "cuotaIva4" },
                {23, "porcentajeRecargo4" },
                {24, "cuotaRecargo4" },
                {25, "baseIrpf" },
                {26, "porcentajeIrpf" },
                {27, "cuotaIrpf" },
                {28, "totalFactura" },
                {29, "nifFactura" },
                {30, "apellidoFactura" },
                {31, "nombreFactura" },
                {32, "periodoFactura" },
                {33, "cuentaContable" },
                {34, "cuentaContrapartida" },
                {35, "codigoConcepto" },
                {36, "fechaOperacion" },
                {37, "facturaDeducible" },
                {38, "direccionFactura" },
                {39, "codPostalFactura" },
                {40, "paisFactura" },
                {41, "baseFactura5" },
                {42, "porcentajeIva5" },
                {43, "cuotaIva5" },
                {44, "porcentajeRecargo5" },
                {45, "cuotaRecargo5" },
                {46, "baseFactura6" },
                {47, "porcentajeIva6"},
                {48, "cuotaIva6" },
                {49, "porcentajeRecargo6" },
                {50, "cuotaRecargo6" },
                {51, "baseFactura7" },
                {52, "porcentajeIva7" },
                {53, "cuotaIva7" },
                {54, "porcentajeRecargo7" },
                {55, "cuotaRecargo7" },
                {56, "baseFactura8" },
                {57, "porcentajeIva8" },
                {58, "cuotaIva8" },
                {59, "porcentajeRecargo8" },
                {60, "cuotaRecargo8" },
                {61, "baseFactura9" },
                {62, "porcentajeIva9" },
                {63, "cuotaIva9" },
                {64, "porcentajeRecargo9" },
                {65, "cuotaRecargo9" }
            };
        }


        public static List<Facturas> obtenerDatos()
        {
            return ListaFacturas;
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

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using UtilidadesDiagram;

namespace importadorFacturas
{
    public class EmitidasE01 : Facturas //Hereda de 'Facturas' para tener todos los campos necesarios
    {
        //Campos especificos para la importacion de Alcasal. El atributo 'OrdenCsv' sirve para colocar esos campos en el orden que tiene esa exportacion a csv (empieza por 1000 para dejar esos numeros reservados para futuros campos de la clase base. Se generan nuevas propiedades para ocultar la de la clase base y poder modificar el orden

        [OrdenCsv(1010)]
        public string primerNumero { get; set; } //No existe en la clase base por lo que no necesita el 'new'

        [OrdenCsv(1020)]
        public string ultimoNumero { get; set; }//No existe en la clase base por lo que no necesita el 'new'

        [OrdenCsv(1030)]
        public int contadorFacturas { get; set; }//No existe en la clase base por lo que no necesita el 'new'

        [OrdenCsv(1040)]
        public new string nifFactura { get; set; }//Para sustituir a la propiedad de la clase base se crea una nueva propiedad con el 'new'

        [OrdenCsv(1050)]
        public string apellidoFactura { get; set; }//No existe en la clase base por lo que no necesita el 'new'

        [OrdenCsv(1060)]
        public new string nombreFactura { get; set; }//Para sustituir a la propiedad de la clase base se crea una nueva propiedad con el 'new'

        [OrdenCsv(1070)]
        public new string direccionFactura { get; set; }//Para sustituir a la propiedad de la clase base se crea una nueva propiedad con el 'new'

        [OrdenCsv(1080)]
        public new string codPostalFactura { get; set; }//Para sustituir a la propiedad de la clase base se crea una nueva propiedad con el 'new'


        //Lista que recoge todas las facturas que luego se exportaran
        public static List<EmitidasE01> ListaIngresosE01 { get; set; } = new List<EmitidasE01>();

        public EmitidasE01()
        {
            //Constructor de la clase que asigna los nombres de las propiedades que se van a incluir en el fichero de salida. Nota: lo dejo por si fuera necesario algun dia inicializar alguna propiedad aunque ahora no es necesario.

        }

        public static List<EmitidasE01> ObtenerFacturasE01()
        {
            return ListaIngresosE01;
        }
    }

    public class ProcesoAlcasal
    {
        //Metodo para procesar los datos del cliente Alcalsal (Raiña Asesores) - tiquet 5863-37
        public StringBuilder EmitidasAlcasar()
        {
            //Almacena los errores si se producen
            StringBuilder resultado = new StringBuilder();

            //Instanciacion de las clases para las facturas agrupadas (tipos T y TR)
            facturasAgrupadas agrupacionT = new facturasAgrupadas();
            facturasAgrupadas agrupacionTR = new facturasAgrupadas();

            //Carga las columnas a procesar y a exportar
            MapeoColumnas();

            //Carga los datos del fichero excel
            var datosExcel = Program.proceso.LeerExcel();

            var numFila = 0; //Permite controlar la fila en la que se ha podido producir un error
            var numColumna = 0;//Permite controla la columna en la que se ha podido producir un error

            try
            {
                //Procesado de las filas
                foreach(var fila in datosExcel)
                {
                    numFila++; //Se incrementa en uno para empezar por el numero 1

                    //Se crea una nueva factura para cada linea
                    var factura = new EmitidasE01();

                    //Se ponen las agrupaciones de facturas en false antes de procesar cada linea y poder sumarlas si se corresponde con la serie T o TR
                    agrupacionT.agrupar = false;
                    agrupacionTR.agrupar = false;

                    //Procesado de las columnas de cada fila
                    foreach(var columna in fila)
                    {
                        numColumna = columna.Key; //Se asigna el numero de columna

                        //Asignacion de valores a propiedades segun el numero de columna
                        switch(numColumna)
                        {
                            //Fecha factura
                            case 2:
                                //Convierte el valor a formato fecha para poder compararla
                                DateTime fechaFra = Convert.ToDateTime(columna.Value).Date;

                                //Controla si llega una fecha posterior a la que pueda tener la agrupacion para grabar un registro con lo que haya acumulado hasta esa fecha
                                if(!string.IsNullOrEmpty(agrupacionT.fechaFraAgrupada) && fechaFra > Convert.ToDateTime(agrupacionT.fechaFraAgrupada).Date)
                                {
                                    GrabarRegistroAgrupado(agrupacionT);
                                }

                                if(!string.IsNullOrEmpty(agrupacionTR.fechaFraAgrupada) && fechaFra > Convert.ToDateTime(agrupacionTR.fechaFraAgrupada).Date)
                                {
                                    GrabarRegistroAgrupado(agrupacionTR);
                                }

                                //Convierte el valor de la columna a un formato de fecha con tipo de cadena
                                if(DateTime.TryParse(columna.Value, out DateTime fecha))
                                {
                                    factura.fechaFactura = fecha.ToString("dd.MM.yyyy");
                                }
                                else
                                {
                                    factura.fechaFactura = columna.Value.Substring(0, 10);
                                }

                                break;

                            //Numero factura
                            case 3:
                                string numFactura = columna.Value;

                                if(numFactura.StartsWith("F") && numFactura.Substring(0, 2) != "FR")
                                {
                                    factura.serieFactura = numFactura.Substring(0, 3);
                                    factura.numeroFactura = columna.Value.Substring(columna.Value.Length - 6);
                                }
                                if(numFactura.StartsWith("FR"))
                                {
                                    factura.serieFactura = numFactura.Substring(0, 4);
                                    factura.numeroFactura = columna.Value.Substring(columna.Value.Length - 6);
                                }
                                if(numFactura.StartsWith("T") && numFactura.Substring(0, 2) != "TR")
                                {
                                    //Se manda al metodo para controlar la primera y ultima factura de la agrupacion
                                    agrupacionT.AgregarFactura(numFactura, factura);
                                }

                                if(numFactura.StartsWith("TR"))
                                {
                                    //Se manda al metodo para controlar la primera y ultima factura de la agrupacion
                                    agrupacionTR.AgregarFactura(numFactura, factura);
                                }

                                //Referencia factura. Como no viene en el Excel, se pone segun el numero de factura
                                factura.referenciaFactura = numFactura;

                                break;

                            //Nif factura
                            case 4:
                                if(columna.Value != "N/D") factura.nifFactura = columna.Value.ToUpper().Replace(" ", "").Replace("-", "");
                                break;

                            //Nombre factura
                            case 5:
                                if(columna.Value != "N/D") factura.nombreFactura = Utilidades.QuitaRaros(columna.Value.ToUpper());
                                break;

                            //Apellidos factura
                            case 6:
                                if(columna.Value != "N/D") factura.apellidoFactura = Utilidades.QuitaRaros(columna.Value.ToUpper());
                                break;

                            //Direccion factura
                            case 7:
                                if(columna.Value != "N/D") factura.direccionFactura = Utilidades.QuitaRaros(columna.Value.ToUpper().Replace(";", ","));
                                break;

                            //Codigo postal factura
                            case 8:
                                string cp = columna.Value;
                                if(cp != "N/D")
                                {
                                    //Añade ceros a la izquierda si el codigo postal tiene menos de 5 digitos
                                    if(cp.Length < 5)
                                    {
                                        factura.codPostalFactura = cp.PadLeft(5, '0');
                                    }
                                    else
                                    {
                                        factura.codPostalFactura = cp;
                                    }
                                }
                                break;

                            //Base factura
                            case 11:
                                decimal valorBase = decimal.Parse(columna.Value);

                                //Se redondea a dos decimales por si en el origen hubieran mas
                                valorBase = Math.Round(valorBase, 2);

                                //Suma la base segun si es no agrupada o de las agrupaciones T o TR
                                if(agrupacionT.agrupar)
                                {
                                    agrupacionT.baseAgrupada += valorBase;
                                }
                                else if(agrupacionTR.agrupar)
                                {
                                    agrupacionTR.baseAgrupada += valorBase;
                                }
                                else
                                {
                                    factura.baseFactura2 = valorBase;
                                }
                                break;

                            //Porcentaje IVA
                            case 12:
                                float valorPorcentaje = float.Parse(columna.Value);

                                //Toma el primer porcentaje en el caso de ser la agrupacion T o TR
                                if(agrupacionT.agrupar)
                                {
                                    if(agrupacionT.porcentajeAgrupado == 0.0f) agrupacionT.porcentajeAgrupado = valorPorcentaje;
                                }
                                else if(agrupacionTR.agrupar)
                                {
                                    if(agrupacionTR.porcentajeAgrupado == 0.0f) agrupacionTR.porcentajeAgrupado = valorPorcentaje;
                                }
                                else
                                {
                                    factura.porcentajeIva2 = valorPorcentaje;
                                }
                                break;

                            //Cuota IVA
                            case 13:
                                decimal valorCuota = decimal.Parse(columna.Value);

                                //Redondea la cuota a dos decimales por si en el origen hay mas
                                valorCuota = Math.Round(valorCuota, 2);

                                //Suma la cuota segun si es no agrupada o de las agrupaciones T o TR
                                if(agrupacionT.agrupar)
                                {
                                    agrupacionT.cuotaAgrupada += valorCuota;
                                }
                                else if(agrupacionTR.agrupar)
                                {
                                    agrupacionTR.cuotaAgrupada += valorCuota;
                                }
                                else
                                {
                                    factura.cuotaIva2 = valorCuota;
                                }
                                break;

                            //Total factura
                            case 14:
                                decimal valorTotal = decimal.Parse(columna.Value);

                                //Redondea el total a dos decimales por si en el origen hay mas
                                valorTotal = Math.Round(valorTotal, 2);

                                //Suma el total segun si es no agrupado o de las agrupaciones T o TR
                                if(agrupacionT.agrupar)
                                {
                                    agrupacionT.totalAgrupada += valorTotal;
                                }
                                else if(agrupacionTR.agrupar)
                                {
                                    agrupacionTR.totalAgrupada += valorTotal;
                                }
                                else
                                {
                                    factura.totalFactura = valorTotal;
                                }
                                break;
                        }
                    }

                    //Se añade el registro solo si no es un registro agrupado
                    if(!agrupacionT.agrupar && !agrupacionTR.agrupar) EmitidasE01.ListaIngresosE01.Add(factura);

                    //Se añade el registro si las facturas agrupadas llegan a 9999
                    if(agrupacionT.cantidadFacturas == 9999) GrabarRegistroAgrupado(agrupacionT);
                    if(agrupacionTR.cantidadFacturas == 9999) GrabarRegistroAgrupado(agrupacionTR);
                }

                GrabarRegistroAgrupado(agrupacionT);
                GrabarRegistroAgrupado(agrupacionTR);
                return resultado;
            }
            catch(Exception ex)
            {
                resultado.AppendLine($"Error al procesar los datos en la fila {numFila} y columna {numColumna}. Revise la estructura");
                resultado.AppendLine($"{ex.Message}");
                return resultado;
            }
        }

        //Metodo para generar el registro en la clase cuando se agrupan facturas
        private void GrabarRegistroAgrupado(facturasAgrupadas agrupacion)
        {
            if(agrupacion.cantidadFacturas > 0)
            {
                EmitidasE01.ListaIngresosE01.Add(new EmitidasE01
                {
                    fechaFactura = agrupacion.fechaFraAgrupada,
                    serieFactura = agrupacion.serieFraAgrupada,
                    baseFactura2 = agrupacion.baseAgrupada,
                    porcentajeIva2 = agrupacion.porcentajeAgrupado,
                    cuotaIva2 = agrupacion.cuotaAgrupada,
                    totalFactura = agrupacion.totalAgrupada,
                    primerNumero = agrupacion.primerNumero,
                    ultimoNumero = agrupacion.ultimoNumero,
                    contadorFacturas = agrupacion.cantidadFacturas
                });

                agrupacion.Reiniciar();
            }
        }

        //Metodo para generar el mapeo de columnas que se usara para la generacion de la salida
        private void MapeoColumnas()
        {
            Facturas.MapeoColumnas = new Dictionary<int, string>
            {
                {1, "fechaFactura" },
                {2, "serieFactura" },
                {3, "numeroFactura" },
                {4, "referenciaFactura" },
                {5, "baseFactura2" },
                {6, "porcentajeIva2" },
                {7, "cuotaIva2" },
                {8, "porcentajeRecargo2" },
                {9,"cuotaRecargo2" },
                {10,"baseIrpf" },
                {11,"porcentajeIrpf" },
                {12,"cuotaIrpf" },
                {13,"totalFactura" },
                {14,"primerNumero" },
                {15,"ultimoNumero" },
                {16,"contadorFacturas" },
                {17,"nifFactura" },
                {18,"apellidoFactura" },
                {19,"nombreFactura" },
                {20,"direccionFactura" },
                {21,"codPostalFactura" }
            };

            Facturas.ColumnasAexportar = new List<string>(Facturas.MapeoColumnas.Values).ToList();
        }
    }

    //Clase que representa las propiedades de las facturas agrupadas que acumulan los importes.
    public class facturasAgrupadas
    {
        public bool agrupar;
        public string fechaFraAgrupada;
        public string serieFraAgrupada;
        public decimal baseAgrupada;
        public float porcentajeAgrupado;
        public decimal cuotaAgrupada;
        public decimal totalAgrupada;
        public string primerNumero;
        public string ultimoNumero;
        public int cantidadFacturas;
        public string tipoFactura;


        //Constructor de la clase que inicializa las propiedades
        public facturasAgrupadas()
        {
            agrupar = false;
            fechaFraAgrupada = string.Empty;
            serieFraAgrupada = string.Empty;
            baseAgrupada = 0;
            porcentajeAgrupado = 0.0f;
            cuotaAgrupada = 0;
            totalAgrupada = 0;
            primerNumero = string.Empty;
            ultimoNumero = string.Empty;
            cantidadFacturas = 0;
        }


        //Metodo para controlar la primera factura que aparezca para agrupar, asi como la ultima y el numero de facturas que se han agrupado
        public void AgregarFactura(string numFactura, EmitidasE01 ingreso)
        {
            agrupar = true;
            if(string.IsNullOrEmpty(serieFraAgrupada))
            {
                serieFraAgrupada = numFactura.Substring(0, numFactura.StartsWith("TR") ? 4 : 3).Replace("T", "L");
            }

            if(string.IsNullOrEmpty(fechaFraAgrupada))
            {
                fechaFraAgrupada = ingreso.fechaFactura;
            }

            if(string.IsNullOrEmpty(primerNumero))
            {
                primerNumero = numFactura;
            }

            ultimoNumero = numFactura;
            cantidadFacturas++;
        }

        //Metodo para reiniciar la clase cuando hay un cambio de fecha
        public void Reiniciar()
        {
            agrupar = false;
            fechaFraAgrupada = string.Empty;
            serieFraAgrupada = string.Empty;
            baseAgrupada = 0;
            porcentajeAgrupado = 0.0f;
            cuotaAgrupada = 0;
            totalAgrupada = 0;
            primerNumero = string.Empty;
            ultimoNumero = string.Empty;
            cantidadFacturas = 0;
        }
    }

}

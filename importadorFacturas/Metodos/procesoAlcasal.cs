using importadorFacturas.Metodos;
using System;
using System.Linq.Expressions;
using System.Text;

namespace importadorFacturas
{
    public class procesoAlcasal
    {
        Utilidades util = new Utilidades();
        public StringBuilder emitidasAlcasar(string ficheroEntrada)
        {
            //Metodo para procesar los datos del cliente Alcalsal (Raiña Asesores) - tiquet 5863-37

            Procesos proceso = new Procesos();
            int filaInicio = 1; //Hay que pasar la fila de la cabecera para contar las columnas
            int columnaInicio = 1; //Los datos empiezan en la columna 1
            int columnaFinal = 14; //Para no tener que procesar todas las columnas se lee hasta la 14 que tiene el total factura

            //Devuelve el resultado si hay algun error
            StringBuilder resultado = new StringBuilder();

            //Instanciacion de las clases para las facturas agrupadas (tipos T y TR)
            facturasAgrupadas agrupacionT = new facturasAgrupadas();
            facturasAgrupadas agrupacionTR = new facturasAgrupadas();

            var datosExcel = proceso.leerExcel(ficheroEntrada, filaInicio, columnaInicio, columnaFinal);

            var numFila = 0; //Controla la fila en la que se ha podido producir un error
            var numColumna = 0;//Controla la columna en la que se ha podido producir un error

            //Proceso de los datos leidos
            try
            {
                foreach (var fila in datosExcel)
                {
                    numFila++;
                    //Instanciacion de la clase para cada linea
                    var ingreso = new facturasEmitidas();

                    //Se ponen en false antes de procesar cada linea y poder sumarlas si se corresponde con la serie T o TR
                    agrupacionT.agrupar = false;
                    agrupacionTR.agrupar = false;

                    //Asignar valores a las propiedades
                    foreach (var columna in fila)
                    {
                        numColumna++;
                        //Procesado de las columnas
                        switch (columna.Key)
                        {
                            case 2:
                                //Fecha factura
                                DateTime fechaFra = Convert.ToDateTime(columna.Value).Date;

                                //Controla si llega una fecha posterior a la que pueda tener la agrupacion para grabar un registro con lo que haya acumulado hasta esa fecha
                                if (!string.IsNullOrEmpty(agrupacionT.fechaFraAgrupada) && fechaFra > Convert.ToDateTime(agrupacionT.fechaFraAgrupada).Date)
                                {
                                    grabarRegistroAgrupado(agrupacionT);
                                }
                                if (!string.IsNullOrEmpty(agrupacionTR.fechaFraAgrupada) && fechaFra > Convert.ToDateTime(agrupacionTR.fechaFraAgrupada).Date)
                                {
                                    grabarRegistroAgrupado(agrupacionTR);
                                }

                                ingreso.fechaFactura = columna.Value.Substring(0, 10);

                                break;

                            case 3:
                                //Numero factura
                                string numFactura = columna.Value;

                                if (numFactura.StartsWith("F") && numFactura.Substring(0, 2) != "FR")
                                {
                                    ingreso.serieFactura = numFactura.Substring(0, 3);
                                    ingreso.numeroFactura = columna.Value.Substring(columna.Value.Length - 6);
                                }
                                if (numFactura.StartsWith("FR"))
                                {
                                    ingreso.serieFactura = numFactura.Substring(0, 4);
                                    ingreso.numeroFactura = columna.Value.Substring(columna.Value.Length - 6);
                                }
                                if (numFactura.StartsWith("T") && numFactura.Substring(0, 2) != "TR")
                                {
                                    //Se manda al metodo para controlar la primera y ultima factura de la agrupacion
                                    agrupacionT.AgregarFactura(numFactura, ingreso);
                                }

                                if (numFactura.StartsWith("TR"))
                                {
                                    //Se manda al metodo para controlar la primera y ultima factura de la agrupacion
                                    agrupacionTR.AgregarFactura(numFactura, ingreso);
                                }

                                //Referencia factura. Como no viene en el Excel, se pone segun el numero de factura
                                ingreso.referenciaFactura = numFactura;

                                break;

                            case 4:
                                //Nif factura
                                if (columna.Value != "N/D") ingreso.nifFactura = columna.Value.ToUpper().Replace(" ", "").Replace("-","");
                                break;

                            case 5:
                                //Nombre factura
                                if (columna.Value != "N/D") ingreso.nombreFactura = util.quitaRaros(columna.Value.ToUpper());
                                break;

                            case 6:
                                //Apellidos factura
                                if (columna.Value != "N/D") ingreso.apellidoFactura = util.quitaRaros(columna.Value.ToUpper());
                                break;

                            case 7:
                                //Direccion factura
                                if (columna.Value != "N/D") ingreso.direccionFactura = util.quitaRaros(columna.Value.ToUpper().Replace(";", ","));
                                break;

                            case 8:
                                //Codigo postal factura
                                string cp = columna.Value;
                                if (cp != "N/D")
                                {
                                    if (cp.Length < 5)
                                    {
                                        ingreso.codPostalFactura = cp.PadLeft(5, '0');
                                    }
                                    else
                                    {
                                        ingreso.codPostalFactura = cp;
                                    }
                                }
                                break;

                            case 11:
                                //Base factura
                                decimal valorBase = decimal.Parse(columna.Value);
                                valorBase = Math.Round(valorBase, 2);
                                if (agrupacionT.agrupar)
                                {
                                    agrupacionT.baseAgrupada += valorBase;
                                }
                                else if (agrupacionTR.agrupar)
                                {
                                    agrupacionTR.baseAgrupada += valorBase;
                                }
                                else
                                {
                                    ingreso.baseFactura = valorBase;
                                }
                                break;

                            case 12:
                                //Porcentaje IVA
                                float valorPorcentaje = float.Parse(columna.Value);
                                if (agrupacionT.agrupar)
                                {
                                    if (agrupacionT.porcentajeAgrupado == 0.0f) agrupacionT.porcentajeAgrupado = valorPorcentaje;
                                }
                                else if (agrupacionTR.agrupar)
                                {
                                    if (agrupacionTR.porcentajeAgrupado == 0.0f) agrupacionTR.porcentajeAgrupado = valorPorcentaje;
                                }
                                else
                                {
                                    ingreso.porcentajeIva = valorPorcentaje;
                                }
                                break;

                            case 13:
                                //Cuota IVA
                                decimal valorCuota = decimal.Parse(columna.Value);
                                valorCuota = Math.Round(valorCuota, 2);
                                if (agrupacionT.agrupar)
                                {
                                    agrupacionT.cuotaAgrupada += valorCuota;
                                }
                                else if (agrupacionTR.agrupar)
                                {
                                    agrupacionTR.cuotaAgrupada += valorCuota;
                                }
                                else
                                {
                                    ingreso.cuotaIva = valorCuota;
                                }
                                break;

                            case 14:
                                //Total factura
                                decimal valorTotal = decimal.Parse(columna.Value);
                                valorTotal = Math.Round(valorTotal, 2);
                                if (agrupacionT.agrupar)
                                {
                                    agrupacionT.totalAgrupada += valorTotal;
                                }
                                else if (agrupacionTR.agrupar)
                                {
                                    agrupacionTR.totalAgrupada += valorTotal;
                                }
                                else
                                {
                                    ingreso.totalFactura = valorTotal;
                                }
                                break;
                        }

                    }

                    //Se añade el registro solo si no es un registro agrupado
                    if (!agrupacionT.agrupar && !agrupacionTR.agrupar) facturasEmitidas.ListaIngresos.Add(ingreso);

                    //Se añade el registro si las facturas agrupadas llegan a 9999
                    if (agrupacionT.cantidadFacturas == 9999) grabarRegistroAgrupado(agrupacionT);
                    if (agrupacionTR.cantidadFacturas == 9999) grabarRegistroAgrupado(agrupacionTR);
                }

                grabarRegistroAgrupado(agrupacionT);
                grabarRegistroAgrupado(agrupacionTR);
                return resultado;
            }
            catch (Exception ex)
            {
                resultado.AppendLine($"Error al procesar los datos en la fila {numFila} y columna {numColumna}. Revise la estructura");
                resultado.AppendLine($"{ex.Message}");
                return resultado;
            }
        }

        private void grabarRegistroAgrupado(facturasAgrupadas agrupacion)
        {
            //Metodo para generar el registro en la clase cuando se agrupan facturas
            if (agrupacion.cantidadFacturas > 0)
            {
                facturasEmitidas.ListaIngresos.Add(new facturasEmitidas
                {
                    fechaFactura = agrupacion.fechaFraAgrupada,
                    serieFactura = agrupacion.serieFraAgrupada,
                    baseFactura = agrupacion.baseAgrupada,
                    porcentajeIva = agrupacion.porcentajeAgrupado,
                    cuotaIva = agrupacion.cuotaAgrupada,
                    totalFactura = agrupacion.totalAgrupada,
                    primerNumero = agrupacion.primerNumero,
                    ultimoNumero = agrupacion.ultimoNumero,
                    contadorFacturas = agrupacion.cantidadFacturas
                });

                agrupacion.Reiniciar();
            }
        }
    }

    public class facturasAgrupadas
    {
        //Clase que representa las propiedades de las facturas agrupadas que acumlan los importes.
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

        public facturasAgrupadas()
        {
            //Constructor de la clase que inicializa las propiedades
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

        public void AgregarFactura(string numFactura, facturasEmitidas ingreso)
        {
            //Metodo para controlar la primera factura que aparezca para agrupar, asi como la ultima y el numero de facturas que se han agrupado
            agrupar = true;
            if (string.IsNullOrEmpty(serieFraAgrupada)) serieFraAgrupada = numFactura.Substring(0, numFactura.StartsWith("TR") ? 4 : 3).Replace("T", "L");
            if (string.IsNullOrEmpty(fechaFraAgrupada)) fechaFraAgrupada = ingreso.fechaFactura;
            if (string.IsNullOrEmpty(primerNumero)) primerNumero = numFactura;
            ultimoNumero = numFactura;
            cantidadFacturas++;
        }

        public void Reiniciar()
        {
            //Metodo para reiniciar la clase cuando hay un cambio de fecha
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

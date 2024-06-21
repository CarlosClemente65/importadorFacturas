using System;

namespace importadorFacturas
{
    public class procesoAlcasar
    {
        public void emitidasAlcasar(string ficheroEntrada)
        {
            Procesos proceso = new Procesos();
            int filaInicio = 1; //Hay que pasar la fila de la cabecera para contar las columnas
            int columnaInicio = 1; //Los datos empiezan en la columna 1

            //Variables para agrupar las facturas con serie T
            bool agruparT = false;
            string fechaFraAgrupadaT = string.Empty;
            string serieFraAgrupadaT = string.Empty;
            float baseAgrupadaT = 0.0f;
            float porcentajeAgrupadoT = 0.0f;
            float cuotaAgrupadaT = 0.0f;
            float totalAgrupadaT = 0.0f;
            string primerNumeroT = string.Empty;
            string ultimoNumeroT = string.Empty;
            int cantidadFacturasT = 0;

            //Variables para agrupar las facturas con serie TR
            bool agruparTR = false;
            string fechaFraAgrupadaTR = string.Empty;
            string serieFraAgrupadaTR = string.Empty;
            float baseAgrupadaTR = 0.0f;
            float porcentajeAgrupadoTR = 0.0f;
            float cuotaAgrupadaTR = 0.0f;
            float totalAgrupadaTR = 0.0f;
            string primerNumeroTR = string.Empty;
            string ultimoNumeroTR = string.Empty;
            int cantidadFacturasTR = 0;

            var datosExcel = proceso.leerExcel(ficheroEntrada, filaInicio, columnaInicio);

            //Proceso de los datos leidos
            foreach (var fila in datosExcel)
            {
                var ingreso = new ingresosAlcasar();
                agruparT = false;
                agruparTR = false;

                //Asignar valores a las propiedades
                foreach (var columna in fila)
                {
                    switch (columna.Key)
                    {
                        case 1:
                            //Tipo factura
                            ingreso.tipoFactura = columna.Value;
                            break;

                        case 2:
                            //Fecha factura
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
                                agruparT = true;
                                if (string.IsNullOrEmpty(serieFraAgrupadaT)) serieFraAgrupadaT = numFactura.Substring(0, 3).Replace("T", "L");
                                if (string.IsNullOrEmpty(fechaFraAgrupadaT)) fechaFraAgrupadaT = ingreso.fechaFactura;
                                if (string.IsNullOrEmpty(primerNumeroT)) primerNumeroT = numFactura;
                                ultimoNumeroT = numFactura;

                                //Nota: revisar esta parte porque acumula tambien las TR
                                cantidadFacturasT++;
                            }

                            if (numFactura.StartsWith("TR"))
                            {
                                agruparTR = true;
                                if (string.IsNullOrEmpty(serieFraAgrupadaTR)) serieFraAgrupadaTR = numFactura.Substring(0, 4).Replace("T", "L");
                                if (string.IsNullOrEmpty(fechaFraAgrupadaTR)) fechaFraAgrupadaTR = ingreso.fechaFactura;
                                if (string.IsNullOrEmpty(primerNumeroTR)) primerNumeroTR = numFactura;
                                ultimoNumeroTR = numFactura;
                                cantidadFacturasTR++;
                            }
                            break;

                        case 4:
                            //Nif factura
                            if (columna.Value != "N/D") ingreso.nifFactura = columna.Value.ToUpper();
                            break;

                        case 5:
                            //Nombre factura
                            if (columna.Value != "N/D") ingreso.nombreFactura = columna.Value.ToUpper();
                            break;

                        case 6:
                            //Apellidos factura
                            if (columna.Value != "N/D") ingreso.apellidoFactura = columna.Value.ToUpper();
                            break;

                        case 7:
                            //Direccion factura
                            if (columna.Value != "N/D") ingreso.direccionFactura = columna.Value.ToUpper();
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

                        case 9:
                            //Base factura
                            float valorBase = float.Parse(columna.Value);
                            if (agruparT)
                            {
                                baseAgrupadaT += valorBase;
                            }
                            else if (agruparTR)
                            {
                                baseAgrupadaTR += valorBase;
                            }
                            else
                            {
                                ingreso.baseFactura = valorBase;
                            }
                            break;

                        case 10:
                            //Porcentaje IVA
                            float valorPorcentaje = float.Parse(columna.Value);
                            if (agruparT)
                            {
                                if (porcentajeAgrupadoT != 0.0f) porcentajeAgrupadoT = valorPorcentaje;
                            }
                            else if (agruparTR)
                            {
                                if (porcentajeAgrupadoTR != 0.0f) porcentajeAgrupadoTR = valorPorcentaje;
                            }
                            else
                            {
                                ingreso.porcentajeIva = valorPorcentaje;
                            }
                            break;

                        case 11:
                            //Cuota IVA
                            float valorCuota = float.Parse(columna.Value);
                            if (agruparT)
                            {
                                cuotaAgrupadaT += valorCuota;

                            }
                            else if (agruparTR)
                            {
                                cuotaAgrupadaTR += valorCuota;

                            }
                            else
                            {
                                ingreso.cuotaIva = valorCuota;
                            }
                            break;

                        case 12:
                            //Total factura
                            float valorTotal = float.Parse(columna.Value);
                            if (agruparT)
                            {
                                totalAgrupadaT += valorTotal;
                            }
                            else if (agruparTR)
                            {
                                totalAgrupadaTR += valorTotal;
                            }
                            else
                            {
                                ingreso.totalFactura = valorTotal;
                            }
                            break;
                    }
                }

                //Se añade el registro solo si no es un registro agrupado
                if (!agruparT && !agruparTR) ingresosAlcasar.ListaIngresos.Add(ingreso);
            }

            if (cantidadFacturasT > 0)
            {
                ingresosAlcasar.ListaIngresos.Add(new ingresosAlcasar
                {
                    tipoFactura = "F4",
                    fechaFactura = fechaFraAgrupadaT,
                    serieFactura = serieFraAgrupadaT,
                    baseFactura = (float)Math.Round(baseAgrupadaT, 2),
                    porcentajeIva = porcentajeAgrupadoT,
                    cuotaIva = (float)Math.Round(cuotaAgrupadaT, 2),
                    totalFactura = (float)Math.Round(totalAgrupadaT, 2),
                    fechaFraAgrupada = fechaFraAgrupadaT,
                    primerNumero = primerNumeroT,
                    ultimoNumero = ultimoNumeroT,
                    contadorFacturas = cantidadFacturasT
                });
            }

            if (cantidadFacturasTR > 0)
            {
                ingresosAlcasar.ListaIngresos.Add(new ingresosAlcasar
                {
                    tipoFactura = "R5",
                    fechaFactura = fechaFraAgrupadaTR,
                    serieFactura = serieFraAgrupadaTR,
                    baseFactura = (float)Math.Round(baseAgrupadaTR, 2),
                    porcentajeIva = porcentajeAgrupadoTR,
                    cuotaIva = (float)Math.Round(cuotaAgrupadaTR, 2),
                    totalFactura = (float)Math.Round(totalAgrupadaTR, 2),
                    fechaFraAgrupada = fechaFraAgrupadaTR,
                    primerNumero = primerNumeroTR,
                    ultimoNumero = ultimoNumeroTR,
                    contadorFacturas = cantidadFacturasTR
                });
            }
        }
    }
}

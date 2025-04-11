using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
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
            porcentajeRecargo2 = 0.0f; //Se inicializa a 0 porque en la clase 'Facturas esta inicializada con 1.4 y aqui no hay recargo.

        }

        public static List<EmitidasE01> ObtenerFacturasE01()
        {
            return ListaIngresosE01;
        }
    }

    public class RecibidasR01 : Facturas //Hereda de 'Facturas' para tener todos los campos necesarios
    {
        //Campos especificos para la importacion de Alcasal. El atributo 'OrdenCsv' sirve para colocar esos campos en el orden que tiene esa exportacion a csv (empieza por 1000 para dejar esos numeros reservados para futuros campos de la clase base. Se generan nuevas propiedades para ocultar la de la clase base y poder modificar el orden//Hereda de 'Facturas' para tener todos los campos necesarios



        public static List<RecibidasR01> ListaRecibidasR01 { get; set; } = new List<RecibidasR01>();

        public RecibidasR01()
        {
            //Constructor de la clase que asigna los nombres de las propiedades que se van a incluir en el fichero de salida. Nota: lo dejo por si fuera necesario algun dia inicializar alguna propiedad aunque ahora no es necesario.

        }

        public static List<RecibidasR01> ObtenerFacturasR01()
        {
            return ListaRecibidasR01;
        }

    }


    public class ProcesoAlcasal
    {
        //Metodo para procesar las facturas emitidas del cliente Alcalsal (Raiña Asesores) - tiquet 5863-37
        public StringBuilder EmitidasAlcasal()
        {
            //Almacena los errores si se producen
            StringBuilder resultado = new StringBuilder();

            //Carga las columnas a procesar y a exportar
            MapeoColumnasE01();

            //Carga los datos del fichero excel
            var datosExcel = Program.proceso.LeerExcel();

            //Se ordenan los datos por la fecha y numero de factura para evitar errores en la agrupacion
            var datosExcelOrdenados = datosExcel
                .Select(fila => new
                {
                    Fila = fila,
                    Fecha = DateTime.Parse(fila[2]), // Convierte la fecha de string a DateTime
                    NumeroFactura = int.Parse(fila[3]) // Convierte el número de factura de string a entero
                })
                .OrderBy(x => x.Fecha) // Primero ordenamos por fecha
                .ThenBy(x => x.NumeroFactura) // Luego por número de factura
                .Select(x => x.Fila) // Seleccionamos solo la fila original
                .ToList();


            var numFila = 0; //Permite controlar la fila en la que se ha podido producir un error
            var numColumna = 0;//Permite controlar la columna en la que se ha podido producir un error

            try
            {
                //Se crea un diccionario para almacenar las agrupaciones de las facturas. Si posteriormente se necesitan mas agrupaciones, se pueden añadir al diccionario
                //Nota: se instancian las agrupaciones y cuando se tengan que usar se hace accediendo al diccionario del siguiente modo: agrupaciones["T"] o el que corresponda (es lo mismo que instanciar con los nombre agrupacionT, agrupacionTR, etc)
                Dictionary<string, facturasAgrupadas> agrupaciones = new Dictionary<string, facturasAgrupadas>
                    {
                        { "T", new facturasAgrupadas() }, //Tiquet anterior 
                        { "TR", new facturasAgrupadas() }, //Tiquet rectificativo anterior
                        { "ESTA", new facturasAgrupadas()}, //Tiquet automatico España
                        { "ESTR", new facturasAgrupadas() }, //Tiquet rectificativo España
                        { "DETA", new facturasAgrupadas() }, //Tiquet automatico Alemania
                        { "DETR", new facturasAgrupadas() } //Tiquet rectificativo Alemania
                    };

                //Procesado de las filas
                foreach(var fila in datosExcelOrdenados)
                {
                    numFila++; //Se incrementa en uno para empezar por el numero 1

                    //Se asigna el numero de factura completo para poder obtener la serie y el numero de la factura
                    string numFacturaCompleto = fila[3];
                    string serieFactura = ObtenerSerieNumero(numFacturaCompleto).serie; //Se obtiene la serie de la factura
                    string numeroFactura = ObtenerSerieNumero(numFacturaCompleto).numero; //Se obtiene el numero de la factura

                    ////Si la serie es TE (serie de pruebas), no se procesa esta factura y salta a la siguiente
                    if(serieFactura == "TE")
                    {
                        continue;
                    }

                    //Se crea una nueva factura para cada linea
                    var factura = new EmitidasE01();

                    // Resetear todas las agrupaciones y asignar el tipo de agrupacion en una nueva instancia
                    foreach(var tipo in agrupaciones.Keys)
                    {
                        agrupaciones[tipo].agrupar = false; //Reinicia la propiedad 'agrupar' para controlar cuando llegue la primera factura de ese tipo
                    }

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
                                foreach(var agrupacion in agrupaciones.Values)
                                {
                                    if(!string.IsNullOrEmpty(agrupacion.fechaFraAgrupada) && fechaFra > Convert.ToDateTime(agrupacion.fechaFraAgrupada).Date)
                                    {
                                        GrabarRegistroAgrupado(agrupacion); // Grabar el registro y reiniciar
                                    }
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
                                //Referencia factura. Como no viene en el Excel, se pone segun el numero de factura
                                factura.referenciaFactura = numFacturaCompleto;

                                switch(serieFactura)
                                {
                                    //Las facturas agrupadas se mandan al metodo para controlar la primera y ultima factura de la agrupacion

                                    case "T":
                                        //Serie Tiquets anterior
                                        agrupaciones["T"].AgregarFactura("L", numeroFactura, factura); //Cuando la serie es una T se convierte a L para la agrupacion
                                        break;

                                    case "TR":
                                        //Serie tiquets rectificativos anterior
                                        agrupaciones["TR"].AgregarFactura("LR", numeroFactura, factura); //Cuando la serie es una TR se convierte a LR para la agrupacion
                                        break;

                                    case "ESTA":
                                        //Serie tiquet automatico España
                                        agrupaciones["ESTA"].AgregarFactura("LESTA", numeroFactura, factura); //Cuando la serie es una ESTA se convierte a LESTA para la agrupacion
                                        break;

                                    case "ESTR":
                                        //Serie tiquet rectificativo España
                                        agrupaciones["ESTR"].AgregarFactura("LESTR", numeroFactura, factura); //Cuando la serie es una ESTR se convierte a LESTR para la agrupacion
                                        break;

                                    case "DETA":
                                        //Serie tiquet automatico Alemania
                                        agrupaciones["DETA"].AgregarFactura("LDETA", numeroFactura, factura); //Cuando la serie es una DETA se convierte a LDETA para la agrupacion
                                        break;

                                    case "DETR":
                                        //Serie tiquet rectificativo Alemania
                                        agrupaciones["DETR"].AgregarFactura("LDETR", numeroFactura, factura); //Cuando la serie es una DETR se convierte a LDETR para la agrupacion
                                        break;

                                    default:
                                        //Resto de facturas se tratan igual
                                        factura.serieFactura = serieFactura;
                                        factura.numeroFactura = numeroFactura;
                                        break;
                                }

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

                                //Suma la base a la base agrupada de las series agrupadas
                                if(agrupaciones.ContainsKey(serieFactura) && agrupaciones[serieFactura].agrupar)
                                {
                                    agrupaciones[serieFactura].baseAgrupada += valorBase;
                                }
                                else
                                {
                                    factura.baseFactura2 = valorBase;
                                }
                                break;


                            //Porcentaje IVA
                            case 12:
                                float valorPorcentaje = float.Parse(columna.Value);

                                //Toma el primer porcentaje en el caso de ser facturas agrupadas
                                if(agrupaciones[serieFactura].agrupar)
                                {
                                    if(agrupaciones[serieFactura].porcentajeAgrupado == 0.0f) agrupaciones[serieFactura].porcentajeAgrupado = valorPorcentaje;
                                }
                                else
                                {
                                    factura.porcentajeIva2 = valorPorcentaje * 100; //Se multiplica por 100 porque en el origen viene como decimal
                                }

                                break;

                            //Cuota IVA
                            case 13:
                                decimal valorCuota = decimal.Parse(columna.Value);

                                //Redondea la cuota a dos decimales por si en el origen hay mas
                                valorCuota = Math.Round(valorCuota, 2);

                                //Suma la cuota a la cuota agrupada de las series agrupadas
                                if(agrupaciones.ContainsKey(serieFactura) && agrupaciones[serieFactura].agrupar)
                                {
                                    agrupaciones[serieFactura].cuotaAgrupada += valorCuota;
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

                                //Suma el total al total agrupado de las series agrupadas
                                if(agrupaciones.ContainsKey(serieFactura) && agrupaciones[serieFactura].agrupar)
                                {
                                    agrupaciones[serieFactura].totalAgrupada += valorTotal;
                                }
                                else
                                {
                                    factura.totalFactura = valorTotal;
                                }

                                break;
                        }
                    }

                    //Se añade el registro solo si no es un registro agrupado
                    if(agrupaciones.ContainsKey(serieFactura) && !agrupaciones[serieFactura].agrupar)
                    {
                        EmitidasE01.ListaIngresosE01.Add(factura);
                    }


                    //Se añade el registro si las facturas agrupadas llegan a 9999
                    foreach(var kvp in agrupaciones)
                    {
                        if(kvp.Value.cantidadFacturas == 9999)
                        {
                            GrabarRegistroAgrupado(kvp.Value);
                        }
                    }
                }

                //Una vez procesadas todas las filas, se añade el registro de todas las facturas agrupadas
                foreach(var kvp in agrupaciones)
                {
                    GrabarRegistroAgrupado(kvp.Value);
                }

                return resultado;
            }
            catch(Exception ex)
            {
                resultado.AppendLine($"Error al procesar los datos en la fila {numFila} y columna {numColumna}. Revise la estructura");
                resultado.AppendLine($"{ex.Message}");
                return resultado;
            }
        }

        //Metodo para procesar las facturas recibidas del cliente Alcasal (Raiña Asesores) - tiquet 5863-59
        public StringBuilder RecibidasAlcasal()
        {
            //Almacena los errores si se producen
            StringBuilder resultado = new StringBuilder();

            //Genera la lista de las columnas a exportar segun el defecto de facturas
            Facturas.MapeoFacturas();
            Facturas.ColumnasAexportar = new List<string> { "contador" }; //Se añade una columna para poner el contador de facturas.
            Facturas.ColumnasAexportar.AddRange(Facturas.MapeoColumnas.Values);

            //Carga los datos del fichero de entrada
            List<string[]> datosEntrada = Program.proceso.LeerCsv(Configuracion.FicheroEntrada);

            int numFila = 0; //Permite controlar la fila en la que se ha podido producir un error
            int numColumna = 0; //Permite controlar la columna en la que se ha podido producir un error
            try
            {
                foreach(var fila in datosEntrada.Skip(1)) //Salta la primera linea ya que es la cabecera de datos
                {
                    numFila++;
                    //Instancia de una nueva factura para procesar
                    var factura = new RecibidasR01();

                    //Añade el contador de facturas
                    factura.contador = numFila;

                    //Procesado de todas las filas 
                    for(int columna = 0; columna < fila.Length; columna++)
                    {
                        string valor = fila[columna];
                        numColumna++;
                        //if(valor != null)
                        //{
                        switch(columna + 1) //Se suma 1 porque el array empieza por 0
                        {
                            //Nombre proveedor
                            case 2:
                                factura.nombreFactura = Utilidades.QuitaRaros(fila[columna].ToUpper());
                                break;

                            //CIF proveedor
                            case 3:
                                factura.nifFactura = fila[columna].ToUpper();
                                break;

                            //Nº factura
                            case 6:
                                //Se almacena como 'referenciaFactura' porque son compras; poner automaticamente el numero de factura que corresponda al importar
                                factura.referenciaFactura = fila[columna].ToUpper();
                                break;


                            //Fecha factura
                            case 7:
                                //Convierte el valor de la columna a un formato de fecha con tipo de cadena
                                if(DateTime.TryParse(fila[columna], out DateTime fecha))
                                {
                                    factura.fechaFactura = fecha.ToString("dd.MM.yyyy");
                                }
                                else
                                {
                                    factura.fechaFactura = fila[columna].Substring(0, 10);
                                }

                                break;

                            //Base al 0%
                            case 9:
                                if(!string.IsNullOrEmpty(valor))
                                {
                                    factura.baseFactura4 = decimal.Parse(fila[columna]);
                                }
                                break;

                            //Cuota al 0%
                            case 10:
                                if(!string.IsNullOrEmpty(valor))
                                {
                                    factura.cuotaIva4 = decimal.Parse(fila[columna]);
                                    if(factura.cuotaIva4 != 0)
                                    {
                                        resultado.AppendLine($"La factura de la linea {numFila} del proveedor {factura.nombreFactura} y fecha {factura.fechaFactura} tiene un importe en la cuota de IVA del 0%. Revisela");
                                    }
                                }
                                break;

                            //Base al 2%
                            case 11:
                                if(!string.IsNullOrEmpty(valor))
                                {
                                    factura.baseFactura5 = decimal.Parse(fila[columna]);
                                    factura.porcentajeIva5 = 2.0f;
                                }
                                break;

                            //Cuota al 2%
                            case 12:
                                if(!string.IsNullOrEmpty(valor))
                                {
                                    factura.cuotaIva5 = decimal.Parse(fila[columna]);
                                }
                                break;

                            //Base al 4%
                            case 13:
                                if(!string.IsNullOrEmpty(valor))
                                {
                                    factura.baseFactura3 = decimal.Parse(fila[columna]);
                                }
                                break;

                            //Cuota al 4%
                            case 14:
                                if(!string.IsNullOrEmpty(valor))
                                {
                                    factura.cuotaIva3 = decimal.Parse(fila[columna]);
                                }
                                break;

                            //Base al 5%
                            case 15:
                                if(!string.IsNullOrEmpty(valor))
                                {
                                    factura.baseFactura6 = decimal.Parse(fila[columna]);
                                    factura.porcentajeIva6 = 5.0f;
                                }
                                break;

                            //Cuota al 5%
                            case 16:
                                if(!string.IsNullOrEmpty(valor))
                                {
                                    factura.cuotaIva6 = decimal.Parse(fila[columna]);
                                }
                                break;

                            //Base al 7,5%
                            case 17:
                                if(!string.IsNullOrEmpty(valor))
                                {
                                    factura.baseFactura7 = decimal.Parse(fila[columna]);
                                    factura.porcentajeIva7 = 7.5f;
                                }
                                break;

                            //Cuota al 7,5%
                            case 18:
                                if(!string.IsNullOrEmpty(valor))
                                {
                                    factura.cuotaIva7 = decimal.Parse(fila[columna]);
                                }
                                break;

                            //Base al 10%
                            case 19:
                                if(!string.IsNullOrEmpty(valor))
                                {
                                    factura.baseFactura2 = decimal.Parse(fila[columna]);
                                }
                                break;

                            //Cuota al 10%
                            case 20:
                                if(!string.IsNullOrEmpty(valor))
                                {
                                    factura.cuotaIva2 = decimal.Parse(fila[columna]);
                                }
                                break;

                            //Base al 21%
                            case 21:
                                if(!string.IsNullOrEmpty(valor))
                                {
                                    factura.baseFactura1 = decimal.Parse(fila[columna]);
                                }
                                break;

                            //Cuota al 21%
                            case 22:
                                if(!string.IsNullOrEmpty(valor))
                                {
                                    factura.cuotaIva1 = decimal.Parse(fila[columna]);
                                }
                                break;

                            //Base IRPF
                            case 24:
                                if(!string.IsNullOrEmpty(valor))
                                {
                                    factura.baseIrpf = decimal.Parse(fila[columna]);
                                }
                                break;

                            //Cuota IRPF
                            case 25:
                                if(!string.IsNullOrEmpty(valor))
                                {
                                    factura.cuotaIrpf = decimal.Parse(fila[columna]);
                                }
                                break;

                            //Porcentaje IRPF
                            case 26:
                                if(!string.IsNullOrEmpty(valor))
                                {
                                    factura.porcentajeIrpf = float.Parse(fila[columna]);
                                }
                                break;

                            //Total factura
                            case 27:
                                if(!string.IsNullOrEmpty(valor))
                                {
                                    factura.totalFactura = decimal.Parse(fila[columna]);
                                }
                                break;

                        }
                        //}
                    }

                    RecibidasR01.ListaRecibidasR01.Add(factura);

                }
            }

            catch(Exception ex)
            {
                resultado.AppendLine($"Error al procesar los datos en la fila {numFila} y columna {numColumna}. Revise la estructura");
                resultado.AppendLine($"{ex.Message}");
                return resultado;
            }

            return resultado;
        }

        //Permite chequear si hay alguna cuota de IVA que no esta bien calculada
        private void ChequeoCuotaIva<T>(T factura, Func<T, decimal> obtenerBase, Func<T, float> obtenerPorentaje, Func<T, decimal> obtenerCuota, string tipoIva, StringBuilder resultado, int numFila) where T : Facturas
        {
            //Almacena los valores de las propiedades segun la base, porcentaje y cuota pasada al metodo
            decimal baseFactura = obtenerBase(factura);
            float porcentajeIva = obtenerPorentaje(factura);
            decimal cuotaIva = obtenerCuota(factura);

            decimal cuotaCalculada = Math.Round(baseFactura * (decimal)porcentajeIva / 100, 2);

            //Si la cuota calculada difiere en mas o menos 5 centimos, genera el error en la factura.
            if(Math.Abs(cuotaIva - cuotaCalculada) > 0.05m)
            {
                resultado.AppendLine($"Descuadre en la cuota de IVA del {tipoIva} en la factura de la linea {numFila} del proveedor {factura.nombreFactura} y fecha {factura.fechaFactura}. Revisela");
            }
        }


        //Metodo para generar el registro en la clase cuando se agrupan facturas emitidas
        private void GrabarRegistroAgrupado(facturasAgrupadas agrupacion)
        {
            if(agrupacion.cantidadFacturas > 0)
            {
                EmitidasE01.ListaIngresosE01.Add(new EmitidasE01
                {
                    fechaFactura = agrupacion.fechaFraAgrupada,
                    serieFactura = agrupacion.serieFraAgrupada,
                    baseFactura2 = agrupacion.baseAgrupada,
                    porcentajeIva2 = agrupacion.porcentajeAgrupado * 100, //Se multiplica por 100 porque en el origen viene como decimal
                    cuotaIva2 = agrupacion.cuotaAgrupada,
                    totalFactura = agrupacion.totalAgrupada,
                    primerNumero = agrupacion.primerNumero,
                    ultimoNumero = agrupacion.ultimoNumero,
                    contadorFacturas = agrupacion.cantidadFacturas
                });

                agrupacion.Reiniciar();
            }
        }

        //Metodo para generar el mapeo de columnas que se usara para la generacion de la salida en facturas emitidas
        private void MapeoColumnasE01()
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

        ////Metodo para generar el mapeo de columnas que se usara para la generacion de la salida en facturas recibidas
        //private void MapeoColumnasR01()
        //{
        //    Facturas.MapeoColumnas = new Dictionary<int, string>
        //    {
        //        {1,"nombreFactura" },
        //        {2,"nifFactura" },
        //        {5, "numeroFactura" },
        //        {5, "referenciaFactura" },
        //        {6, "fechaFactura" },
        //        {8,"totalFactura" },
        //        {9, "baseFactura4" },
        //        {10, "porcentajeIva2" },
        //        {7, "cuotaIva2" },
        //        {8, "porcentajeRecargo2" },
        //        {9,"cuotaRecargo2" },
        //        {10,"baseIrpf" },
        //        {11,"porcentajeIrpf" },
        //        {12,"cuotaIrpf" },

        //    };

        //    Facturas.ColumnasAexportar = new List<string>(Facturas.MapeoColumnas.Values).ToList();
        //}


        //Metodo para obtener la serie y el numero de la factura a partir del numero de factura
        private (string serie, string numero) ObtenerSerieNumero(string numFactura)
        {
            var regex = new Regex(@"([a-zA-Z]+)(\d+)");
            var match = regex.Match(numFactura);

            if(match.Success)
            {
                // Retorna las partes como una tupla
                return (match.Groups[1].Value, match.Groups[2].Value);
            }
            else
            {
                // Si no hay coincidencia, retorna valores vacíos
                return (string.Empty, string.Empty);
            }
        }
    }

    public interface IAgrupacion
    {
        string fechaFraAgrupada { get; }
    }

    //Clase que representa las propiedades de las facturas agrupadas que acumulan los importes.
    public class facturasAgrupadas : IAgrupacion
    {
        public bool agrupar { get; set; }
        public string fechaFraAgrupada { get; set; }
        public string serieFraAgrupada { get; set; }
        public decimal baseAgrupada { get; set; }
        public float porcentajeAgrupado { get; set; }
        public decimal cuotaAgrupada { get; set; }
        public decimal totalAgrupada { get; set; }
        public string primerNumero { get; set; }
        public string ultimoNumero { get; set; }
        public int cantidadFacturas { get; set; }
        public string tipoFactura { get; set; }


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


        //Metodo para controlar la primera factura que aparezca para agrupar, asi como la ultima y el numero de facturas que se han agrupado, pasando el numero de factura completo (metodo anterior)
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

        //Sobrecarga del metodo para controlar la primera factura que aparezca para agrupar, asi como la ultima y el numero de facturas que se han agrupado, pasando la serie y numero por separado
        public void AgregarFactura(string serie, string numero, EmitidasE01 ingreso)
        {
            agrupar = true;
            string numeroCompleto = serie + numero;
            if(string.IsNullOrEmpty(serieFraAgrupada))
            {
                serieFraAgrupada = serie;
            }
            if(string.IsNullOrEmpty(fechaFraAgrupada))
            {
                fechaFraAgrupada = ingreso.fechaFactura;
            }
            if(string.IsNullOrEmpty(primerNumero))
            {
                primerNumero = numeroCompleto;
            }
            ultimoNumero = numeroCompleto;
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

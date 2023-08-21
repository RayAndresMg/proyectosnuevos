using System;
using Npgsql;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OracleClient;
using System.Data;
using System.Configuration;
using System.Net.Mail;
using System.Net;
using System.Net.Mime;
using ClosedXML.Excel;
using System.IO;
using System.Data.Sql;
using System.Data.SqlClient;


namespace ConsoleCorreo
{
    class Program
    {
        //string conexionstring = ConfigurationManager.ConnectionStrings["conexion"].ConnectionString;

        static string nombreDoc = "";
        static int cantidaddias = 0;
        static int dia = 0;

        static void Main(string[] args)
        {
            #region excel

            var valorIngresado = "1";
            Console.Clear();
            Console.WriteLine("\n");
            Console.WriteLine("::::::::::::::::GENERADOR DE REPORTES VULNERACIONES MOVISTAR AUTOACTIVADO - INSOLUTIONS::::::::::::::::");
            Console.WriteLine("\n");
            Console.ForegroundColor = ConsoleColor.Red;

            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine("\n");
            Console.WriteLine(":::::::::::::::::::::::::::::::MODO AUTOMÁTICO:::::::::::::::::::::::::::::::");
            Console.WriteLine("\n");


            if (!String.IsNullOrEmpty(Convert.ToString(valorIngresado)))
            {

                if (Char.IsNumber(Convert.ToChar(valorIngresado)))
                {
                    cantidaddias = Convert.ToInt32(valorIngresado);


                    cantidaddias = cantidaddias * -dia;

                    Console.WriteLine("Generar reporte Biometría para Movistar AA");



                    DateTime thisDay = DateTime.Today.AddDays(-1);
                    var fechaAyer = thisDay.ToString("yyyy/MM/dd");

                    DateTime thisDay2 = DateTime.Today.AddDays(0);

                    var fechahoy= thisDay2.ToString("yyyy/MM/dd");

                    Console.WriteLine(fechahoy);

                    Console.WriteLine("Incio de generacion de reporte...");

                    NpgsqlConnection conexionpgs = new NpgsqlConnection();

                    string servidor = "rds-pg-adp-is-main.postgres.database.azure.com";
                    string bd = "IS.Movistar.AutoActivaChip_DB";
                    string usuario = "ispgadmin@rds-pg-adp-is-main";
                    string password = "cf5wM@H53yV1iLtR!*9t9mkg";
                    string puerto = "5432";

                    String cadenaConexion = "server=" + servidor + ";" + "port=" + puerto + ";" + "user id=" + usuario + ";" + "password=" + password + ";" + "database=" + bd + ";";


                    conexionpgs.ConnectionString = cadenaConexion;
                    conexionpgs.Open();

                    Console.Write("Ingrese los Id de vulneración: \n");
                    var vulneraciones = Console.ReadLine();
                    
                        var wb = new XLWorkbook();
                        var ws = wb.Worksheets.Add("Hoja1");
                        int num = 2;

                        string query11 = "select rep.\"Id\", rep.\"Dni\",rep.\"NombreCompleto\", rep.\"TipoCliente\", rep.\"FechaLogin\", rep.\"InicioProceso\", rep.\"DniVerificado\", rep.\"AceptacionTerminos\",rep.\"MejorHuella\", rep.\"ValidacionBiometrica\", rep.\"ActivacionEnviada\", rep.\"EstadoProceso\", rep.\"TipoOperacion\", rep.\"ChipNumber\", rep.\"PhoneNumber\",rep.\"Plan\", rep.\"FlagProteccionDatos\", rep.\"OrdenAmdocs\", rep.\"FechaEmision\", rep.\"FlagDNIValidation\",rep.\"Rta1\", rep.\"Rta2\", rep.\"Rta3\", rep.\"Rta4\", rep.\"Rta5\",rep.\"UniqueDeviceId\", rep.\"DeviceBrand\", rep.\"DeviceModel\", rep.\"AppVersion\", tk.\"Os\", o2.\"PromotionsAndNewsAccepted\", (case when oc.\"Code\" is null then 'AC' else oc.\"Code\" end ) \"Flujo\"";

                    //string query13 = "select rep.\"Id\", rep.\"Dni\", rep.\"NombreCompleto\", '' \"TipoCliente\", rep.\"FechaLogin\", rep.\"InicioProceso\", rep.\"DniVerificado\", rep.\"AceptacionTerminos\",rep.\"MejorHuella\", rep.\"ValidacionBiometrica\",'-' \"ActivacionEnviada\", rep.\"EstadoProceso\",'-' \"TipoOperacion\", '-' \"ChipNumber\" , '-' \"PhoneNumber\",'-' \"Plan\", rep.\"FlagProteccionDatos\",'-' \"OrdenAmdocs\", rep.\"FechaEmision\", rep.\"FlagDNIValidation\", rep.\"Rta1\", rep.\"Rta2\", rep.\"Rta3\", rep.\"Rta4\", rep.\"Rta5\",rep.\"UniqueDeviceId\", rep.\"DeviceBrand\", rep.\"DeviceModel\", rep.\"AppVersion\", tk.\"Os\", o2.\"PromotionsAndNewsAccepted\", (case when oc.\"Code\" is null then 'AC' else oc.\"Code\" end ) \"Flujo\"";
                        string query14 = "  from \"IS\".\"VW_RepOperation\" rep inner join \"IS\".\"TokenApi\" tk on (rep.\"Id\" = tk.\"Id\") inner join \"IS\".\"Operation\" o2 on (rep.\"Id\" = o2.\"Id\") left outer join \"IS\".\"OperationChannel\" oc on(o2.\"CanalId\" = oc.\"Id\"  ) where  rep.\"Id\" in ("+ vulneraciones+")";

                        string queryTransacciones2 = query11 + query14;

                        var cmd7 = new NpgsqlCommand(queryTransacciones2, conexionpgs);
                        DataTable dataTablepgtrans2 = new DataTable();
                        dataTablepgtrans2.Load(cmd7.ExecuteReader());

                        foreach (DataRow dataRow in dataTablepgtrans2.Rows)
                        {
                            string Id = dataRow["Id"].ToString();
                            Console.WriteLine("Registro numero => " + num);

                            //string Id = dataRow["Id"].ToString();
                            ws.Cell("A" + num).Value = dataRow["Id"].ToString();
                            ws.Cell("A" + num).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                            ws.Cell("A" + num).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                            ws.Cell("B" + num).Value = dataRow["Dni"].ToString();
                            ws.Cell("B" + num).DataType = XLCellValues.Text;
                            ws.Cell("B" + num).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                            ws.Cell("B" + num).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                            ws.Cell("C" + num).Value = dataRow["NombreCompleto"].ToString();
                            ws.Cell("C" + num).DataType = XLCellValues.Text;
                            ws.Cell("C" + num).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                            ws.Cell("C" + num).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                            ws.Cell("D" + num).Value = dataRow["TipoCliente"].ToString();
                            ws.Cell("D" + num).DataType = XLCellValues.Text;
                            ws.Cell("D" + num).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                            ws.Cell("D" + num).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                            var dateTm1p = dataRow["FechaLogin"].ToString();
                            if (dateTm1p != "")
                            {
                                DateTime dateTmp1 = Convert.ToDateTime(dateTm1p);

                                //DateTime fecha1 = dateTmp.AddHours(-5);

                                ws.Cell("E" + num).Value = dateTmp1.ToString("dd/MM/yyyy HH:mm:ss");
                                //ws.Cell("F" + num).SetDataType(XLCellValues.DateTime);
                                ws.Cell("E" + num).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                                ws.Cell("E" + num).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            }
                            else
                            {
                                ws.Cell("E" + num).Value = "";
                                //ws.Cell("F" + num).SetDataType(XLCellValues.DateTime);
                                ws.Cell("E" + num).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                                ws.Cell("E" + num).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            }

                            var dateTm2p = dataRow["InicioProceso"].ToString();
                            if (dateTm2p != "")
                            {
                                DateTime dateTmp2 = Convert.ToDateTime(dateTm2p);

                                //DateTime fecha1 = dateTmp.AddHours(-5);

                                ws.Cell("F" + num).Value = dateTmp2.ToString("dd/MM/yyyy HH:mm:ss");
                                //ws.Cell("F" + num).SetDataType(XLCellValues.DateTime);
                                ws.Cell("F" + num).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                                ws.Cell("F" + num).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            }
                            else
                            {
                                ws.Cell("F" + num).Value = "";
                                //ws.Cell("F" + num).SetDataType(XLCellValues.DateTime);
                                ws.Cell("F" + num).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                                ws.Cell("F" + num).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            }

                            var dateTm3p = dataRow["DniVerificado"].ToString();
                            if (dateTm3p != "")
                            {
                                DateTime dateTmp3 = Convert.ToDateTime(dateTm3p);

                                //DateTime fecha1 = dateTmp.AddHours(-5);

                                ws.Cell("G" + num).Value = dateTmp3.ToString("dd/MM/yyyy HH:mm:ss");
                                //ws.Cell("F" + num).SetDataType(XLCellValues.DateTime);
                                ws.Cell("G" + num).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                                ws.Cell("G" + num).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            }
                            else
                            {
                                ws.Cell("G" + num).Value = "";
                                //ws.Cell("F" + num).SetDataType(XLCellValues.DateTime);
                                ws.Cell("G" + num).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                                ws.Cell("G" + num).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            }

                            var dateTm4p = dataRow["AceptacionTerminos"].ToString();
                            if (dateTm4p != "")
                            {
                                DateTime dateTmp4 = Convert.ToDateTime(dateTm4p);

                                //DateTime fecha1 = dateTmp.AddHours(-5);

                                ws.Cell("H" + num).Value = dateTmp4.ToString("dd/MM/yyyy HH:mm:ss");
                                //ws.Cell("F" + num).SetDataType(XLCellValues.DateTime);
                                ws.Cell("H" + num).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                                ws.Cell("H" + num).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            }
                            else
                            {
                                ws.Cell("H" + num).Value = "";
                                //ws.Cell("F" + num).SetDataType(XLCellValues.DateTime);
                                ws.Cell("H" + num).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                                ws.Cell("H" + num).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            }

                            var dateTm5p = dataRow["MejorHuella"].ToString();
                            if (dateTm5p != "")
                            {
                                DateTime dateTmp5 = Convert.ToDateTime(dateTm5p);

                                //DateTime fecha1 = dateTmp.AddHours(-5);

                                ws.Cell("I" + num).Value = dateTmp5.ToString("dd/MM/yyyy HH:mm:ss");
                                //ws.Cell("F" + num).SetDataType(XLCellValues.DateTime);
                                ws.Cell("I" + num).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                                ws.Cell("I" + num).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            }
                            else
                            {
                                ws.Cell("I" + num).Value = "";
                                //ws.Cell("F" + num).SetDataType(XLCellValues.DateTime);
                                ws.Cell("I" + num).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                                ws.Cell("I" + num).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            }

                            var dateTm6p = dataRow["ValidacionBiometrica"].ToString();
                            if (dateTm6p != "")
                            {
                                DateTime dateTmp6 = Convert.ToDateTime(dateTm6p);

                                //DateTime fecha1 = dateTmp.AddHours(-5);

                                ws.Cell("J" + num).Value = dateTmp6.ToString("dd/MM/yyyy HH:mm:ss");
                                //ws.Cell("F" + num).SetDataType(XLCellValues.DateTime);
                                ws.Cell("J" + num).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                                ws.Cell("J" + num).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            }
                            else
                            {
                                ws.Cell("J" + num).Value = "";
                                //ws.Cell("F" + num).SetDataType(XLCellValues.DateTime);
                                ws.Cell("J" + num).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                                ws.Cell("J" + num).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            }

                            ws.Cell("K" + num).Value = dataRow["ActivacionEnviada"].ToString();
                            ws.Cell("K" + num).DataType = XLCellValues.Text;
                            ws.Cell("K" + num).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                            ws.Cell("K" + num).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;


                            ws.Cell("L" + num).Value = dataRow["EstadoProceso"].ToString();
                            ws.Cell("L" + num).DataType = XLCellValues.Text;
                            ws.Cell("L" + num).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                            ws.Cell("L" + num).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                            ws.Cell("M" + num).Value = dataRow["TipoOperacion"].ToString();
                            ws.Cell("M" + num).DataType = XLCellValues.Text;
                            ws.Cell("M" + num).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                            ws.Cell("M" + num).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                            ws.Cell("N" + num).Value = dataRow["ChipNumber"].ToString();
                            ws.Cell("N" + num).DataType = XLCellValues.Text;
                            ws.Cell("N" + num).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                            ws.Cell("N" + num).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                            ws.Cell("O" + num).Value = dataRow["PhoneNumber"].ToString();
                            ws.Cell("O" + num).DataType = XLCellValues.Text;
                            ws.Cell("O" + num).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                            ws.Cell("O" + num).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                            ws.Cell("P" + num).Value = dataRow["Plan"].ToString();
                            ws.Cell("P" + num).DataType = XLCellValues.Text;
                            ws.Cell("P" + num).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                            ws.Cell("P" + num).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                            ws.Cell("Q" + num).Value = dataRow["FlagProteccionDatos"].ToString();
                            ws.Cell("Q" + num).DataType = XLCellValues.Text;
                            ws.Cell("Q" + num).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                            ws.Cell("Q" + num).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                            ws.Cell("R" + num).Value = dataRow["OrdenAmdocs"].ToString();
                            ws.Cell("R" + num).DataType = XLCellValues.Text;
                            ws.Cell("R" + num).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                            ws.Cell("R" + num).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                            var dateTm8p = dataRow["FechaEmision"].ToString();

                            if (dateTm8p != "")
                            {
                                DateTime dateTmp8 = Convert.ToDateTime(dateTm8p);

                                ws.Cell("S" + num).Value = dateTmp8.ToString("dd/MM/yyyy");
                                ws.Cell("S" + num).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                                ws.Cell("S" + num).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;


                            }
                            else
                            {
                                ws.Cell("S" + num).Value = "";
                                //ws.Cell("F" + num).SetDataType(XLCellValues.DateTime);
                                ws.Cell("S" + num).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                                ws.Cell("S" + num).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            }

                            ws.Cell("T" + num).Value = dataRow["FlagDNIValidation"].ToString();
                            ws.Cell("T" + num).DataType = XLCellValues.Text;
                            ws.Cell("T" + num).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                            ws.Cell("T" + num).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;


                            var dateTm9p = dataRow["Rta1"].ToString();

                            if (dateTm9p != "")
                            {
                                DateTime dateTmp9 = Convert.ToDateTime(dateTm9p);

                                ws.Cell("U" + num).Value = dateTmp9.ToString("dd/MM/yyyy");
                                ws.Cell("U" + num).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                                ws.Cell("U" + num).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;


                            }
                            else
                            {
                                ws.Cell("U" + num).Value = "";
                                //ws.Cell("F" + num).SetDataType(XLCellValues.DateTime);
                                ws.Cell("U" + num).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                                ws.Cell("U" + num).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            }

                            var dateTm10p = dataRow["Rta2"].ToString();

                            if (dateTm10p != "")
                            {
                                DateTime dateTmp10 = Convert.ToDateTime(dateTm10p);

                                ws.Cell("V" + num).Value = dateTmp10.ToString("dd/MM/yyyy");
                                ws.Cell("V" + num).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                                ws.Cell("V" + num).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;


                            }
                            else
                            {
                                ws.Cell("V" + num).Value = "";
                                //ws.Cell("F" + num).SetDataType(XLCellValues.DateTime);
                                ws.Cell("V" + num).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                                ws.Cell("V" + num).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            }

                            var dateTm11p = dataRow["Rta3"].ToString();

                            if (dateTm11p != "")
                            {
                                DateTime dateTmp11 = Convert.ToDateTime(dateTm11p);

                                ws.Cell("W" + num).Value = dateTmp11.ToString("dd/MM/yyyy");
                                ws.Cell("W" + num).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                                ws.Cell("W" + num).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;


                            }
                            else
                            {
                                ws.Cell("W" + num).Value = "";
                                //ws.Cell("F" + num).SetDataType(XLCellValues.DateTime);
                                ws.Cell("W" + num).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                                ws.Cell("W" + num).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            }

                            ///
                            var dateTm12p = dataRow["Rta4"].ToString();
                            if (dateTm12p != "")
                            {
                                DateTime dateTmp12 = Convert.ToDateTime(dateTm12p);

                                ws.Cell("X" + num).Value = dateTmp12.ToString("dd/MM/yyyy");
                                ws.Cell("X" + num).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                                ws.Cell("X" + num).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;


                            }
                            else
                            {
                                ws.Cell("X" + num).Value = "";
                                //ws.Cell("F" + num).SetDataType(XLCellValues.DateTime);
                                ws.Cell("X" + num).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                                ws.Cell("X" + num).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            }
                            ///
                            var dateTm13p = dataRow["Rta5"].ToString();
                            if (dateTm13p != "")
                            {
                                DateTime dateTmp13 = Convert.ToDateTime(dateTm13p);

                                ws.Cell("Y" + num).Value = dateTmp13.ToString("dd/MM/yyyy");
                                ws.Cell("Y" + num).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                                ws.Cell("Y" + num).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;


                            }
                            else
                            {
                                ws.Cell("Y" + num).Value = "";
                                //ws.Cell("F" + num).SetDataType(XLCellValues.DateTime);
                                ws.Cell("Y" + num).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                                ws.Cell("Y" + num).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            }

                            ws.Cell("Z" + num).Value = dataRow["UniqueDeviceId"].ToString();
                            ws.Cell("Z" + num).DataType = XLCellValues.Text;
                            ws.Cell("Z" + num).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                            ws.Cell("Z" + num).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                            ws.Cell("AA" + num).Value = dataRow["DeviceBrand"].ToString();
                            ws.Cell("AA" + num).DataType = XLCellValues.Text;
                            ws.Cell("AA" + num).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                            ws.Cell("AA" + num).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                            ws.Cell("AB" + num).Value = dataRow["DeviceModel"].ToString();
                            ws.Cell("AB" + num).DataType = XLCellValues.Text;
                            ws.Cell("AB" + num).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                            ws.Cell("AB" + num).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                            var variable = dataRow["AppVersion"].ToString();
                            ws.Cell("AC" + num).Value = "v " + variable;
                            //ws.Cell("AC" + num).DataType = XLCellValues.Text;
                            ws.Cell("AC" + num).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                            ws.Cell("AC" + num).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                            ws.Cell("AD" + num).Value = dataRow["Os"].ToString();
                            ws.Cell("AD" + num).DataType = XLCellValues.Text;
                            ws.Cell("AD" + num).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                            ws.Cell("AD" + num).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                            ws.Cell("AE" + num).Value = dataRow["PromotionsAndNewsAccepted"].ToString();
                            ws.Cell("AE" + num).DataType = XLCellValues.Text;
                            ws.Cell("AE" + num).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                            ws.Cell("AE" + num).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;


                            ws.Cell("AF" + num).Value = dataRow["Flujo"].ToString();
                            ws.Cell("AF" + num).DataType = XLCellValues.Text;
                            ws.Cell("AF" + num).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                            ws.Cell("AF" + num).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;




                            num = num + 1;

                        }//foreach (DataRow dataRow in dataTable2.Rows)


                        //CABECERA
                        ws.Cell("A1").Value = "Id";
                        ws.Cell("B1").Value = "Dni";
                        ws.Cell("C1").Value = "NombreCompleto";
                        ws.Cell("D1").Value = "TipoCliente";
                        ws.Cell("E1").Value = "FechaLogin";
                        ws.Cell("F1").Value = "InicioProceso";
                        ws.Cell("G1").Value = "DniVerificado";
                        ws.Cell("H1").Value = "AceptacionTerminos";
                        ws.Cell("I1").Value = "MejorHuella";
                        ws.Cell("J1").Value = "ValidacionBiometrica";
                        ws.Cell("K1").Value = "ActivacionEnviada";
                        ws.Cell("L1").Value = "EstadoProceso";
                        ws.Cell("M1").Value = "TipoOperacion";
                        ws.Cell("N1").Value = "ChipNumber";
                        ws.Cell("O1").Value = "PhoneNumber";
                        ws.Cell("P1").Value = "Plan";
                        ws.Cell("Q1").Value = "FlagProteccionDatos";
                        ws.Cell("R1").Value = "OrdenAmdocs";
                        ws.Cell("S1").Value = "FechaEmision";
                        ws.Cell("T1").Value = "FlagDNIValidation";
                        ws.Cell("U1").Value = "Rta1";
                        ws.Cell("V1").Value = "Rta2";
                        ws.Cell("W1").Value = "Rta3";
                        ws.Cell("X1").Value = "Rta4";
                        ws.Cell("Y1").Value = "Rta5";
                        ws.Cell("Z1").Value = "UniqueDeviceId";
                        ws.Cell("AA1").Value = "DeviceBrand";
                        ws.Cell("AB1").Value = "DeviceModel";
                        ws.Cell("AC1").Value = "AppVersion";
                        ws.Cell("AD1").Value = "Os";
                        ws.Cell("AE1").Value = "PromotionsAndNewsAccepted";
                        ws.Cell("AF1").Value = "Flujo";

                        var cabeceras = ws.Range("A1:AF1");
                        ws.Columns().AdjustToContents();

                        cabeceras.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                        cabeceras.Style.Font.SetBold(true);

                        cabeceras.Style.Font.FontColor = XLColor.White;
                        cabeceras.Style.Fill.BackgroundColor = XLColor.FromArgb(0, 112, 192);
                        cabeceras.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                        cabeceras.SetAutoFilter();


                        DateTime diadehoy = DateTime.Today;
                        DateTime hoy = DateTime.Now;
                        int hour = hoy.Hour;
                        //ReporteBiometrico_290723
                        if(hour < 15) { 
                        //nombreDoc = "C\\Users\\Rodrigo Quinteros\\Documents\\HuellaEntel\\MiHuellaEntel_Trazabilidad_SegundoReporte_" + thisDay.ToString("ddMMyy") + ".xlsx";

                        nombreDoc = "C:\\Movistar Activa Chip - Reportes\\Incidencias_" + diadehoy.ToString("ddMMyy") + "_1.xlsx";
                        //  nombreDoc = "C:\\Users\\insolutions\\Videos\\Reportes-Prueba-Entel\\hola_hola.xlsx";
                        wb.SaveAs(nombreDoc);
                        }
                        else {
                            nombreDoc = "C:\\Movistar Activa Chip - Reportes\\Incidencias_" + diadehoy.ToString("ddMMyy") + "_2.xlsx";
                            //  nombreDoc = "C:\\Users\\insolutions\\Videos\\Reportes-Prueba-Entel\\hola_hola.xlsx";
                            wb.SaveAs(nombreDoc);
                        }
                    //correoElectronico("soporte@insolutions.pe", oCorreosDestinatariosBanBifTo, oCorreosDestinatariosBanBifCC);





                    Console.WriteLine("\n");
                    Console.WriteLine("\n");
                    Console.WriteLine("Se generaron los correos");
                            
                       
                           

                }
                else
                {
                    Console.Write("El valor ingresado no es un numero, intente nuevamente...\n");
                    Console.ReadLine();
                }
            }
            else
            {
                Console.Write("Debe ingresar un valor, intente nuevamente...\n");
                Console.ReadLine();
            }
               

            


            //while (op.Key != ConsoleKey.Escape);
            #endregion
        }


    }
}



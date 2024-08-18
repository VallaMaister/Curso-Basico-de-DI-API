using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbobsCOM;

namespace Academia_SDK_DIAPI
{
    internal class Program
    {
        public static Company oCompany = null;
        static void Main(string[] args)
        {
            if (ConexionSAP())
            {
                //CrearSN();
                //ActualizarSN();
                //EliminarSN();
                //CrearArticulo();
                //ObtenerXMLSN();
                //CrearSNXML();
                //ObtenerXMLArticulo();
                //CrearPedidoVentas();
                //CrearEntrgaConLote();
                //CrearEntregaConSerie();
                //CrearEntregadePedido();

                //UtilizarRecordSet();

                //CrearEntregaconRecordSet();

                //DataBrowser();

                //SBObobObject();
                //TC();
                //TablaUsuario();
                //CrearCampoUsuario();
                //string[,] ValorValido = { { "SI", "SI" }, { "NO", "NO" } }; 
                //CrearCampoValoresValidos(ValorValido);
                //InsertarDatosTabla();
                //ActualizarinfoSociedad();
                //ActualizarLotes();
                //CrearPagoRecibido();
                //SolicitudesTraslado();
                //TrasladoStock();
                //EntradasMercanciaDirectas();
                //SalidasMercanciaDirectas();
                //CrearAlmacenes();
                //ActualizarAlmacenes();

                //CreanEntregaUbicaciones();
                TransferenciaUbicaciones();
            }


            if (oCompany.Connected)
            {
                oCompany.Disconnect();
                Console.WriteLine("Desconexion exitosa");
            }
            Console.ReadKey();
        }

        static bool ConexionSAP()
        {
            bool respuesta = false;

            try
            {
                oCompany = new Company();
                oCompany.Server = ConfigurationManager.AppSettings["Server"];
                oCompany.CompanyDB = ConfigurationManager.AppSettings["CompanyDB"];
                oCompany.UserName = ConfigurationManager.AppSettings["UserName"];
                oCompany.Password = ConfigurationManager.AppSettings["Password"];
                oCompany.DbUserName = ConfigurationManager.AppSettings["DbUserName"];
                oCompany.DbPassword = ConfigurationManager.AppSettings["DbPassword"];
                oCompany.language = BoSuppLangs.ln_Spanish_La;


                switch (ConfigurationManager.AppSettings["DbServerType"])
                {
                    case "SQL2019":
                        oCompany.DbServerType = BoDataServerTypes.dst_MSSQL2019;
                        break;
                    case "SQL2017":
                        oCompany.DbServerType = BoDataServerTypes.dst_MSSQL2017;
                        break;
                    case "SQL2016":
                        oCompany.DbServerType = BoDataServerTypes.dst_MSSQL2016;
                        break;
                    case "SQL2014":
                        oCompany.DbServerType = BoDataServerTypes.dst_MSSQL2014;
                        break;
                    case "HANA":
                        oCompany.DbServerType = BoDataServerTypes.dst_HANADB;
                        break;
                }

                int iRes = oCompany.Connect();

                if (iRes == 0)
                {
                    respuesta = true;
                    Console.WriteLine("Conexion Correcta con: " + oCompany.CompanyDB);
                }
                else
                {
                    Console.WriteLine("Error: " + oCompany.GetLastErrorDescription());
                }


            }
            catch (Exception e)
            {

            }

            return respuesta;
        }

        static void CrearSN()
        {
            BusinessPartners oBP = oCompany.GetBusinessObject(BoObjectTypes.oBusinessPartners);
            oBP.CardCode = "V08";
            oBP.CardName = "Oscar Valladares";
            oBP.FederalTaxID = "VACO931104K69";

            //Añadir persona de contacto

            oBP.ContactEmployees.Name = "Oscar Valladares";
            oBP.ContactEmployees.FirstName = "Oscar";
            oBP.ContactEmployees.LastName = "Valladares";
            oBP.ContactEmployees.Position = "Desarrollador";
            oBP.ContactEmployees.Add();

            oBP.ContactEmployees.Name = "Eduardo Calleros";
            oBP.ContactEmployees.FirstName = "Eduardo";
            oBP.ContactEmployees.LastName = "Calleros";
            oBP.ContactEmployees.Position = "Desarrollador";
            oBP.ContactEmployees.Add();

            oBP.Addresses.AddressName = "Entrega";
            oBP.Addresses.AddressType = BoAddressType.bo_ShipTo;
            oBP.Addresses.Street = "Ejemplo";
            oBP.Addresses.Add();

            oBP.Addresses.AddressName = "Fiscal";
            oBP.Addresses.AddressType = BoAddressType.bo_BillTo;
            oBP.Addresses.Street = "Ejemplo2";
            oBP.Addresses.Add();


            if (oBP.Add() == 0)
            {
                Console.WriteLine("Se creo el SN Correctamente: " + oCompany.GetNewObjectKey());
            }
            else
            {
                Console.WriteLine("Error: " + oCompany.GetLastErrorDescription());
            }
        }

        static void ActualizarSN()
        {
            BusinessPartners oBP = oCompany.GetBusinessObject(BoObjectTypes.oBusinessPartners);
            if (oBP.GetByKey("V08"))
            {

                Console.WriteLine("Contactos: " + oBP.ContactEmployees.Count);

                if (oBP.ContactEmployees.Count > 0)
                {
                    for (int i = 0; i < oBP.ContactEmployees.Count; i++)
                    {
                        oBP.ContactEmployees.SetCurrentLine(i);
                        if (oBP.ContactEmployees.FirstName == "Eduardo")
                            oBP.ContactEmployees.Position = "Desarrollador";

                    }
                    if (oBP.Update() == 0)
                    {
                        Console.WriteLine("Actualizacion exitosa de: " + oCompany.GetNewObjectKey());
                    }
                    else
                    {
                        Console.WriteLine("Error: " + oCompany.GetLastErrorDescription());
                    }
                }


            }
            else
            {
                Console.WriteLine("No existe el SN dado");
            }
        }

        static void EliminarSN()
        {
            BusinessPartners oBP = oCompany.GetBusinessObject(BoObjectTypes.oBusinessPartners);
            if (oBP.GetByKey("V08"))
            {
                if (oBP.Remove() == 0)
                {
                    Console.WriteLine("Se elimino el SN: " + oCompany.GetNewObjectKey());
                }
                else
                {
                    Console.WriteLine("Error:" + oCompany.GetLastErrorDescription());
                }
            }
            else
            {
                Console.WriteLine("No existe el SN dado");
            }
        }

        static void CrearArticulo()
        {
            Items oItem = oCompany.GetBusinessObject(BoObjectTypes.oItems);

            oItem.ItemCode = "SDK05";
            oItem.ItemName = "Articulo SDK 5";
            oItem.PurchaseItem = BoYesNoEnum.tNO;
            oItem.InventoryItem = BoYesNoEnum.tYES;
            oItem.SalesItem = BoYesNoEnum.tYES;

            oItem.DefaultWarehouse = "01";

            /*Asignacion de propiedades*/
            oItem.Properties[1] = BoYesNoEnum.tYES;
            oItem.Properties[3] = BoYesNoEnum.tYES;
            oItem.Properties[5] = BoYesNoEnum.tYES;
            oItem.Properties[7] = BoYesNoEnum.tYES;

            if (oItem.Add() == 0)
            {
                Console.WriteLine("Articulo creado con exito: " + oCompany.GetNewObjectKey());
            }
            else
            {
                Console.WriteLine("Error: " + oCompany.GetLastErrorCode() + " " + oCompany.GetLastErrorDescription());
            }
        }

        static void ObtenerXMLSN()
        {
            try
            {
                BusinessPartners oBP = oCompany.GetBusinessObject(BoObjectTypes.oBusinessPartners);
                oCompany.XmlExportType = BoXmlExportTypes.xet_ExportImportMode;
                if (oBP.GetByKey("V08"))
                {
                    oBP.SaveXML("C:\\Users\\AdministratorD\\Desktop\\Academia SDK DIAPI\\" + oBP.CardCode + ".xml");
                    Console.WriteLine("XML Exportado con Exito");
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
        }

        static void ObtenerXMLArticulo()
        {
            try
            {
                Items oItem = oCompany.GetBusinessObject(BoObjectTypes.oItems);
                oCompany.XmlExportType = BoXmlExportTypes.xet_ExportImportMode;
                if (oItem.GetByKey("SDK01"))
                {
                    oItem.SaveXML("C:\\Users\\AdministratorD\\Desktop\\Academia SDK DIAPI\\" + oItem.ItemCode + ".xml");
                    Console.WriteLine("XML Exportado con Exito");
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
        }
        static void CrearSNXML()
        {
            string sRuta = "C:\\Users\\AdministratorD\\Desktop\\Academia SDK DIAPI\\V08.xml";
            int count, ii;
            BusinessPartners oBP = null;

            count = oCompany.GetXMLelementCount(sRuta);
            for (ii = 0; ii < count; ii++)
            {
                if (oCompany.GetXMLobjectType(sRuta, ii) == BoObjectTypes.oBusinessPartners)
                {
                    oBP = oCompany.GetBusinessObjectFromXML(sRuta, ii);
                    if (oBP.Add() == 0)
                    {
                        Console.WriteLine("SN creado con exito: " + oCompany.GetNewObjectKey());
                    }
                    else
                    {
                        Console.WriteLine("Error:" + oCompany.GetLastErrorDescription());
                    }
                }
            }
        }

        static void CrearPedidoVentas()
        {
            Documents oOrder = oCompany.GetBusinessObject(BoObjectTypes.oOrders);

            //Informacion de Cabecera
            oOrder.CardCode = "V08";
            oOrder.DocDate = DateTime.Now;
            oOrder.DocDueDate = DateTime.Now;
            oOrder.Comments = "Orden de venta creada con DI API";
            oOrder.NumAtCard = "123456987";
            oOrder.UserFields.Fields.Item("U_TS_CopiaA").Value = "NO";
            int CantidadArticulos = 5;
            Random random = new Random();
            //Informacion de Detalle
            for (int i = 1; i <= CantidadArticulos; i++)
            {
                if (i != 1)
                    oOrder.Lines.Add();

                oOrder.Lines.ItemCode = "SDK0" + i;
                oOrder.Lines.Quantity = random.NextDouble();
                oOrder.Lines.UnitPrice = random.NextDouble();
                oOrder.Lines.WarehouseCode = "01";

            }

            if (oOrder.Add() == 0)
            {
                int DocEntry = Convert.ToInt32(oCompany.GetNewObjectKey());
                if (oOrder.GetByKey(DocEntry))
                    Console.WriteLine("Orden Creada con exito #" + oOrder.DocNum);
            }
            else
            {
                Console.WriteLine("Error: " + oCompany.GetLastErrorDescription());
            }
        }

        static void CrearEntrgaConLote()
        {
            Documents oDelivery = oCompany.GetBusinessObject(BoObjectTypes.oDeliveryNotes);

            //Encabezado
            oDelivery.CardCode = "V08";
            oDelivery.DocDate = DateTime.Now;
            oDelivery.DocDueDate = DateTime.Now;
            oDelivery.Comments = "ajsgdasj";
            oDelivery.NumAtCard = "123246";

            //detalle
            oDelivery.Lines.ItemCode = "SDK06";
            oDelivery.Lines.Quantity = 5;
            oDelivery.Lines.UnitPrice = 20;
            oDelivery.Lines.WarehouseCode = "01";

            //Detalle de lo lotes

            oDelivery.Lines.BatchNumbers.BatchNumber = "Prueba 01";
            oDelivery.Lines.BatchNumbers.Quantity = 3;
            oDelivery.Lines.BatchNumbers.Add();
            oDelivery.Lines.BatchNumbers.BatchNumber = "Prueba 02";
            oDelivery.Lines.BatchNumbers.Quantity = 2;
            oDelivery.Lines.BatchNumbers.Add();


            oDelivery.Lines.Add();

            if (oDelivery.Add() == 0)
            {
                int DocEntry = Convert.ToInt32(oCompany.GetNewObjectKey());
                if (oDelivery.GetByKey(DocEntry))
                    Console.WriteLine("Entrega #" + oDelivery.DocNum);
            }
            else
            {
                Console.WriteLine("Error: " + oCompany.GetLastErrorDescription());
            }
        }

        static void CrearEntregaConSerie()
        {

            Documents oDelivery = oCompany.GetBusinessObject(BoObjectTypes.oDeliveryNotes);

            //Encabezado
            oDelivery.CardCode = "V08";
            oDelivery.DocDate = DateTime.Now;
            oDelivery.DocDueDate = DateTime.Now;
            oDelivery.Comments = "ajsgdasj";
            oDelivery.NumAtCard = "123246";

            //detalle
            oDelivery.Lines.ItemCode = "SDK07";
            oDelivery.Lines.Quantity = 10;
            oDelivery.Lines.UnitPrice = 20;
            oDelivery.Lines.WarehouseCode = "01";

            //Detalle de las series
            for (int i = 1; i <= oDelivery.Lines.Quantity; i++)
            {
                if (i == 10)
                    oDelivery.Lines.SerialNumbers.InternalSerialNumber = "Serie-" + i;
                else
                    oDelivery.Lines.SerialNumbers.InternalSerialNumber = "Serie-0" + i;

                oDelivery.Lines.SerialNumbers.Add();
            }

            oDelivery.Lines.Add();

            if (oDelivery.Add() == 0)
            {
                int DocEntry = Convert.ToInt32(oCompany.GetNewObjectKey());
                if (oDelivery.GetByKey(DocEntry))
                    Console.WriteLine("Entrega #" + oDelivery.DocNum);
            }
            else
            {
                Console.WriteLine("Error: " + oCompany.GetLastErrorDescription());
            }

        }

        static void CrearEntregadePedido()
        {
            Documents oOrder = oCompany.GetBusinessObject(BoObjectTypes.oOrders);
            Documents oDelivery = oCompany.GetBusinessObject(BoObjectTypes.oDeliveryNotes);

            if (oOrder.GetByKey(663))
            {
                oDelivery.Comments = "Entrega vinculada de una orden de venta #" + oOrder.DocNum;

                for (int i = 0; i < oOrder.Lines.Count; i++)
                {
                    oOrder.Lines.SetCurrentLine(i);

                    if (oOrder.Lines.RemainingOpenQuantity > 0)
                    {
                        oDelivery.Lines.BaseEntry = oOrder.DocEntry;
                        oDelivery.Lines.BaseType = Convert.ToInt32(BoObjectTypes.oOrders);
                        oDelivery.Lines.BaseLine = oOrder.Lines.LineNum;
                        oDelivery.Lines.Add();
                    }

                }

                if (oDelivery.Add() == 0)
                {
                    int DocEntry = Convert.ToInt32(oCompany.GetNewObjectKey());
                    if (oDelivery.GetByKey(DocEntry))
                        Console.WriteLine("Entrega #" + oDelivery.DocNum);
                }
                else
                {
                    Console.WriteLine("Error: " + oCompany.GetLastErrorDescription());
                }
            }
        }

        static void UtilizarRecordSet()
        {
            try
            {
                Recordset oRet = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

                string sQuery = "EXEC EjemploRecordSet";

                oRet.DoQuery(sQuery);

                if (oRet.RecordCount > 0)
                {
                    while (!oRet.EoF)
                    {
                        Console.WriteLine("ItemCode: " + oRet.Fields.Item("Codigo").Value);
                        Console.WriteLine("ItemName: " + oRet.Fields.Item("Nombre").Value);
                        Console.WriteLine("SellItem: " + oRet.Fields.Item("SellItem").Value);

                        Console.WriteLine("-----------------------------------------------------------------------");

                        oRet.MoveNext();
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        static void CrearEntregaconRecordSet()
        {
            SAPbobsCOM.Documents oDelivery = null;
            SAPbobsCOM.Recordset oRet = null;
            try
            {
                oRet = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                string sQuery = "EXEC TS_OrdenesVenta";
                oRet.DoQuery(sQuery);
                int i = 0;
                if (oRet.RecordCount > 0)
                {
                    oDelivery = oCompany.GetBusinessObject(BoObjectTypes.oDeliveryNotes);

                    oDelivery.Comments = "Entrega creada con RecordSet";
                    //Esta linea se debe de agregar cuando se genera una entrega con dos o mas ordenes de venta
                    oDelivery.CardCode = oRet.Fields.Item("CardCode").Value;


                    while (!oRet.EoF)
                    {
                        Console.WriteLine($"Entrega del cliente: {oRet.Fields.Item("CardCode").Value}");
                        Console.WriteLine($"DocEntry: {oRet.Fields.Item("DocEntry").Value} - LineNum: {oRet.Fields.Item("LineNum").Value}");

                        oDelivery.Lines.BaseEntry = Convert.ToInt32(oRet.Fields.Item("DocEntry").Value);
                        oDelivery.Lines.BaseType = Convert.ToInt32(BoObjectTypes.oOrders);
                        oDelivery.Lines.BaseLine = Convert.ToInt32(oRet.Fields.Item("LineNum").Value);
                        oDelivery.Lines.Add();
                        oRet.MoveNext();
                    }

                    if (oDelivery.Add() == 0)
                    {
                        int DocEntry = Convert.ToInt32(oCompany.GetNewObjectKey());
                        if (oDelivery.GetByKey(DocEntry))
                            Console.WriteLine("Entrega #" + oDelivery.DocNum);
                    }
                    else
                    {
                        oDelivery = null;
                        Console.WriteLine("Error: " + oCompany.GetLastErrorDescription());
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                if (oRet != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRet);
                    oRet = null;
                }

                if (oDelivery != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDelivery);
                    oDelivery = null;
                }

                GC.Collect();
            }
        }

        static void DataBrowser()
        {
            try
            {
                Recordset oRet = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                Items oItem = oCompany.GetBusinessObject(BoObjectTypes.oItems);
                string sQuery = "Select * FROM OITM ";

                oRet.DoQuery(sQuery);

                oItem.Browser.Recordset = oRet;
                oItem.Browser.MoveFirst();
                Console.WriteLine(oItem.ItemCode);
                oItem.Browser.MoveNext();
                Console.WriteLine(oItem.ItemCode);
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
        }

        static void SBObobObject()
        {
            SBObob oSBObob = oCompany.GetBusinessObject(BoObjectTypes.BoBridge);
            Recordset oRet = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

            oRet = oSBObob.GetContactEmployees("V08");
            if (oRet.RecordCount > 0)
            {
                while (!oRet.EoF)
                {
                    Console.WriteLine("Persona de contacto:" + oRet.Fields.Item("Name").Value);
                    oRet.MoveNext();
                }
            }

            oRet = oSBObob.GetCurrencyRate("USD", DateTime.Now);
            Console.WriteLine(oRet.Fields.Item("CurrencyRate").Value);

        }

        static void TC()
        {
            SBObob oSBObob = oCompany.GetBusinessObject(BoObjectTypes.BoBridge);

            oSBObob.SetCurrencyRate("USD", DateTime.Now, 23, true);
            oSBObob.SetCurrencyRate("EUR", DateTime.Now, 25, true);
            oSBObob.SetCurrencyRate("CAN", DateTime.Now, 18, true);
            Console.WriteLine("Tipo de cambio actualizado con exito");
        }

        static void TablaUsuario()
        {
            UserTablesMD oTable = oCompany.GetBusinessObject(BoObjectTypes.oUserTables);
            try
            {
                if (oTable.GetByKey("TS_ACADEMIA"))
                {
                    Console.WriteLine("Tabla TS_ACADEMIA ya existe en la base de datos");
                }
                else
                {
                    Console.WriteLine("Iniciamos con la creacion de la tabla TS_ACADEMIA");

                    oTable.TableName = "TS_ACADEMIA";
                    oTable.TableDescription = "Prueba Academia";
                    oTable.TableType = BoUTBTableType.bott_NoObject;

                    if (oTable.Add() == 0)
                    {
                        Console.WriteLine("Tabla creada con exito: " + oCompany.GetNewObjectKey());
                    }
                    else
                    {
                        Console.WriteLine("Error: " + oCompany.GetLastErrorDescription());
                    }
                }
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); };
        }

        static void CrearCampoUsuario()
        {
            UserFieldsMD oUF = oCompany.GetBusinessObject(BoObjectTypes.oUserFields);


            if (ValidarCampos("@TS_ACADEMIA", "TS_Folio2"))
            {
                oUF.TableName = "TS_ACADEMIA";
                oUF.Name = "TS_Folio2";
                oUF.Description = "Folio2";
                oUF.Type = BoFieldTypes.db_Alpha;
                oUF.EditSize = 50;

                if (oUF.Add() == 0)
                {
                    Console.WriteLine("Se creo el campo de manera exitosa: U_TS_Folio2");
                }
                else
                {
                    Console.WriteLine("Error: " + oCompany.GetLastErrorDescription());
                }
            }

            //Importante siempre limpiar la memoria
            Console.WriteLine("Liberacon de memoria");
            while (System.Runtime.InteropServices.Marshal.ReleaseComObject(oUF) > 0) { Console.WriteLine(System.Runtime.InteropServices.Marshal.ReleaseComObject(oUF)); }
            oUF = null;
            GC.Collect();

        }

        static void CrearCampoValoresValidos(string[,] sValue)
        {
            UserFieldsMD oUF = oCompany.GetBusinessObject(BoObjectTypes.oUserFields);

            oUF.TableName = "TS_ACADEMIA";
            oUF.Name = "TS_Activo";
            oUF.Description = "Activo";
            oUF.Type = BoFieldTypes.db_Alpha;
            oUF.EditSize = 50;

            //oUF.ValidValues.Value = "SI";
            //oUF.ValidValues.Description = "SI";
            //oUF.ValidValues.Add();
            //oUF.ValidValues.Value = "NO";
            //oUF.ValidValues.Description = "NO";
            //oUF.ValidValues.Add();

            for (int i = 0; i < sValue.Length / 2; i++)
            {
                oUF.ValidValues.Value = sValue[i, 0];
                oUF.ValidValues.Description = sValue[i, 1];
                oUF.ValidValues.Add();
            }

            oUF.DefaultValue = "SI";

            if (oUF.Add() == 0)
            {
                Console.WriteLine("Se creo el campo de manera exitosa: U_TS_Activo");
            }
            else
            {
                Console.WriteLine("Error: " + oCompany.GetLastErrorDescription());
            }

            //Importante siempre limpiar la memoria
            Console.WriteLine("Liberacon de memoria");
            while (System.Runtime.InteropServices.Marshal.ReleaseComObject(oUF) > 0) { Console.WriteLine(System.Runtime.InteropServices.Marshal.ReleaseComObject(oUF)); }
            oUF = null;
            GC.Collect();
        }

        static bool ValidarCampos(string sTabla, string sCampo)
        {
            bool Validacion = false;
            try
            {
                Recordset oRet = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

                string sQuery = $"Select U_{sCampo} From \"{sTabla}\"";

                oRet.DoQuery(sQuery);

                if (oRet.RecordCount > 0)
                {
                    Console.WriteLine($"El campo U_{sCampo} ya existe en la tabla {sTabla}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Este campo U_{sCampo} no existe en la tabla {sTabla}");
                Validacion = true;
            }

            return Validacion;
        }

        static void InsertarDatosTabla()
        {
            try
            {
                UserTable oTable = oCompany.UserTables.Item("TS_ACADEMIA");

                oTable.Code = "4";
                oTable.Name = "4";
                oTable.UserFields.Fields.Item("U_TS_Folio").Value = "4";
                oTable.UserFields.Fields.Item("U_TS_Activo").Value = "NO";

                if(oTable.Add() == 0)
                {
                    Console.WriteLine("Informacion insertada con exito: " + oCompany.GetNewObjectKey());
                }
                else
                {
                    Console.WriteLine("Error: " + oCompany.GetLastErrorDescription());
                }

            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
        }

        //Servicios
        static void ActualizarinfoSociedad()
        {
            CompanyService oCompanyService;
            AdminInfo oCompanyAdminInfo;

            oCompanyService = oCompany.GetCompanyService();
            oCompanyAdminInfo = oCompanyService.GetAdminInfo();
            Console.WriteLine("Direccion: " + oCompanyAdminInfo.Address);

            oCompanyAdminInfo.CompanyName = "Tesselar Soluciones";
            oCompanyService.UpdateAdminInfo(oCompanyAdminInfo);
            Console.WriteLine("Se actualizo de manera exitosa");
        }

        static void ActualizarLotes()
        {
            CompanyService oCompanyService = oCompany.GetCompanyService();
            BatchNumberDetailsService bnumService = oCompanyService.GetBusinessService(ServiceTypes.BatchNumberDetailsService);
            BatchNumberDetailParams bNumparams = bnumService.GetDataInterface(BatchNumberDetailsServiceDataInterfaces.bndsBatchNumberDetailParams);

            Recordset oRet = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            string sQuery = "Select AbsEntry,DistNumber from OBTN Where Status=2";
            string sNombre = "";
            oRet.DoQuery(sQuery);

            if(oRet.RecordCount > 0)
            {
                while (!oRet.EoF)
                {
                    sNombre = oRet.Fields.Item("DistNumber").Value;
                    bNumparams.DocEntry = Convert.ToInt32(oRet.Fields.Item("AbsEntry").Value);//Es el AbsEntry
                    BatchNumberDetail bNumDetail = bnumService.Get(bNumparams);
                    bNumDetail.Status = BoDefaultBatchStatus.dbs_Released;
                    bNumDetail.Details = "Lote liberado por DIAPI";
                    bnumService.Update(bNumDetail);
                    Console.WriteLine($"Lote {sNombre} actualizado con exito");
                    oRet.MoveNext();
                }
            }

            
        }

        //Pagos

        static void CrearPagoRecibido()
        {
            try
            {
                Payments oPay = oCompany.GetBusinessObject(BoObjectTypes.oIncomingPayments);

                //Cabecera
                oPay.CardCode = "C1111";
                oPay.DocType = BoRcptTypes.rCustomer;
                oPay.DocDate = DateTime.Now;
                oPay.TransferSum = 0;
                oPay.TransferDate = DateTime.Now;
                oPay.TransferAccount = "11105000";

                //Detalle de la facturas a pagar
                oPay.Invoices.InvoiceType = BoRcptInvTypes.it_Invoice;
                oPay.Invoices.DocEntry = 532;
                oPay.Invoices.SumApplied = 66.40;
                oPay.Invoices.Add();

                oPay.Invoices.InvoiceType = BoRcptInvTypes.it_JournalEntry;
                oPay.Invoices.DocEntry = 3549;
                oPay.Invoices.SumApplied = -16.40;
                oPay.Invoices.Add();

                oPay.Invoices.InvoiceType = BoRcptInvTypes.it_JournalEntry;
                oPay.Invoices.DocEntry = 3550;
                oPay.Invoices.SumApplied = -50;
                oPay.Invoices.Add();


                if (oPay.Add() == 0)
                {
                    int DocEntry = Convert.ToInt32(oCompany.GetNewObjectKey());
                    if(oPay.GetByKey(DocEntry))
                        Console.WriteLine("Pago Creado con exito #" + oPay.DocNum);
                }
                else
                {
                    Console.WriteLine("Error: " + oCompany.GetLastErrorDescription());
                }


            }
            catch (Exception ex)
            {
                Console.WriteLine (ex.Message);
            }
        }


        //Transacciones de Inventario

        static void SolicitudesTraslado()
        {
            try
            {
                StockTransfer oSolitud = oCompany.GetBusinessObject(BoObjectTypes.oInventoryTransferRequest);
                oSolitud.CardCode = "";
                oSolitud.DocDate = DateTime.Now;
                oSolitud.FromWarehouse = "01";
                oSolitud.ToWarehouse = "02";

                //Detalle
                oSolitud.Lines.ItemCode = "SDK01";
                oSolitud.Lines.Quantity = 10;
                oSolitud.Lines.FromWarehouseCode = "04";
                oSolitud.Lines.WarehouseCode = "05";


                if(oSolitud.Add() == 0)
                {
                    Console.WriteLine("Exito");
                }
                else
                {
                    Console.WriteLine("Error: " + oCompany.GetLastErrorDescription());
                }

            }catch (Exception ex)
            {
                Console.WriteLine (ex.Message);
            }
        }

        static void TrasladoStock()
        {
            try
            {
                StockTransfer oTransfer = oCompany.GetBusinessObject (BoObjectTypes.oStockTransfer);
                oTransfer.CardCode = "";
                oTransfer.DocDate = DateTime.Now;
                oTransfer.FromWarehouse = "01";
                oTransfer.ToWarehouse = "02";

                //Detalle
                oTransfer.Lines.ItemCode = "SDK06";
                oTransfer.Lines.Quantity = 10;
                oTransfer.Lines.FromWarehouseCode = "01";
                oTransfer.Lines.WarehouseCode = "05";
                oTransfer.Lines.BatchNumbers.BatchNumber = "Prueba 02";
                oTransfer.Lines.BatchNumbers.Quantity = 10;


                if (oTransfer.Add() == 0)
                {
                    Console.WriteLine("Exito");
                }
                else
                {
                    Console.WriteLine("Error: " + oCompany.GetLastErrorDescription());
                }


            }
            catch (Exception ex)
            {
                Console.WriteLine (ex.Message);
            }
        }

        static void EntradasMercanciaDirectas()
        {
            try
            {
                Documents oDoc = oCompany.GetBusinessObject(BoObjectTypes.oInventoryGenEntry);

                oDoc.DocDate = DateTime.Now;
                oDoc.Comments = "Creado con DI API";

                oDoc.Lines.ItemCode = "SDK01";
                oDoc.Lines.Quantity = 10;
                oDoc.Lines.UnitPrice = 10;
                oDoc.Lines.WarehouseCode = "05";



                if(oDoc.Add() == 0)
                {
                    Console.WriteLine("Exito");
                }
                else
                {
                    Console.WriteLine("Error: " + oCompany.GetLastErrorDescription());
                }

                
            }catch(Exception ex)
            {
                Console.WriteLine (ex.Message);
            }
        }

        static void SalidasMercanciaDirectas()
        {
            try
            {
                Documents oDoc = oCompany.GetBusinessObject(BoObjectTypes.oInventoryGenExit);

                oDoc.DocDate = DateTime.Now;
                oDoc.Comments = "Creado con DI API";

                oDoc.Lines.ItemCode = "SDK06";
                oDoc.Lines.Quantity = 11;
                oDoc.Lines.UnitPrice = 10;
                oDoc.Lines.WarehouseCode = "04";
                oDoc.Lines.BatchNumbers.BatchNumber = "Prueba 01";
                oDoc.Lines.BatchNumbers.Quantity = 11;


                if (oDoc.Add() == 0)
                {
                    Console.WriteLine("Exito");
                }
                else
                {
                    Console.WriteLine("Error: " + oCompany.GetLastErrorDescription());
                }


            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        //Almacenes

        static void CrearAlmacenes()
        {
            try
            {

                Warehouses oWH = oCompany.GetBusinessObject(BoObjectTypes.oWarehouses);

                oWH.WarehouseCode = "06";
                oWH.WarehouseName = "Creado DI API";
                
                if(oWH.Add() == 0) { Console.WriteLine("Exito"); } else { Console.WriteLine("Error: " + oCompany.GetLastErrorDescription()); }


            }catch(Exception ex) 
            {
                Console.WriteLine(ex.Message);
            }
        }

        static void ActualizarAlmacenes()
        {
            try
            {

                Warehouses oWH = oCompany.GetBusinessObject(BoObjectTypes.oWarehouses);

                if (oWH.GetByKey("06"))
                {
                    oWH.EnableBinLocations = BoYesNoEnum.tYES;
                    if (oWH.Update() == 0) { Console.WriteLine("Exito"); } else { Console.WriteLine("Error: " + oCompany.GetLastErrorDescription()); }
                }

                


            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        static void CreanEntregaUbicaciones()
        {
            Documents oDelivery = oCompany.GetBusinessObject(BoObjectTypes.oDeliveryNotes);
            //Encabezado
            oDelivery.CardCode = "V08";
            oDelivery.DocDate = DateTime.Now;
            oDelivery.DocDueDate = DateTime.Now;
            oDelivery.Comments = "ajsgdasj";
            oDelivery.NumAtCard = "123246";

            //detalle
            oDelivery.Lines.ItemCode = "SDK06";
            oDelivery.Lines.Quantity = 5;
            oDelivery.Lines.UnitPrice = 20;
            oDelivery.Lines.WarehouseCode = "05";

            //Detalle de lo lotes

            oDelivery.Lines.BatchNumbers.BatchNumber = "Prueba 01";
            oDelivery.Lines.BatchNumbers.Quantity = 5;
            oDelivery.Lines.BatchNumbers.Add();

            //Asignacion de ubicacion en la linea
            oDelivery.Lines.BinAllocations.BinAbsEntry = 1;
            oDelivery.Lines.BinAllocations.Quantity = 5;
            oDelivery.Lines.BinAllocations.SerialAndBatchNumbersBaseLine = 0;
            oDelivery.Lines.BinAllocations.Add();

            oDelivery.Lines.Add();

            if (oDelivery.Add() == 0)
            {
                int DocEntry = Convert.ToInt32(oCompany.GetNewObjectKey());
                if (oDelivery.GetByKey(DocEntry))
                    Console.WriteLine("Entrega #" + oDelivery.DocNum);
            }
            else
            {
                Console.WriteLine("Error: " + oCompany.GetLastErrorDescription());
            }
        }

        static void TransferenciaUbicaciones()
        {
            // Crear la transferencia de stock
            SAPbobsCOM.StockTransfer oTransfer = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer);
            oTransfer.DocDate = DateTime.Now;


            // Añadir la primera línea de transferencia con lote
            oTransfer.Lines.ItemCode = "SDK06"; // Código del artículo
            oTransfer.Lines.Quantity = 10;       // Cantidad a transferir
            oTransfer.Lines.FromWarehouseCode = "01"; // Almacén de origen
            oTransfer.Lines.WarehouseCode = "05"; // Almacén de destino

            // Asignar lote en el almacén de origen
            oTransfer.Lines.BatchNumbers.BatchNumber = "Prueba 02";
            oTransfer.Lines.BatchNumbers.Quantity = 10;
            oTransfer.Lines.BatchNumbers.Add();


            // Asignar ubicación en el almacén de destino
            int binAbsEntry = 7; // ID de la ubicación en el almacén de destino

            // Asegurarse de que la línea de lote es correcta antes de asignar la ubicación
            oTransfer.Lines.BinAllocations.BinAbsEntry = binAbsEntry;
            oTransfer.Lines.BinAllocations.BinActionType = BinActionTypeEnum.batToWarehouse;
            oTransfer.Lines.BinAllocations.Quantity = 10;
            oTransfer.Lines.BinAllocations.SerialAndBatchNumbersBaseLine = 0;
            oTransfer.Lines.BinAllocations.Add();

            // Añadir la línea al documento
            oTransfer.Lines.Add();

            // Añadir el documento de transferencia a SAP
            int lRetCode = oTransfer.Add();
            if (lRetCode == 0)
            {
                Console.WriteLine("Éxito se creó la transferencia. DocEntry: " + oCompany.GetNewObjectKey());
            }
            else
            {
                Console.WriteLine("Error al crear la transferencia: " + oCompany.GetLastErrorDescription());
            }

        }
    }
}

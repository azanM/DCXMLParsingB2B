using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using System.Data.SqlClient;

namespace ParsingPurchaseOrder
{

    public class ProductDetail
    {
        public string PurcDoc { set; get; }
        public string ItemDoc { set; get; }
        public string DateCreated { set; get; }
        public string DocType { set; get; }
        public string PayTerm { set; get; }
        public string Vendor { set; get; }
        public string DocDate { set; get; }
        public string ChangeDate { set; get; }
        public string Material { set; get; }
        public string MatGroup { set; get; }
        public string DocCond { set; get; }
        public string NetPrice { set; get; }
        public string Curr { set; get; }
        public string POQty { set; get; }
        public string UoM { set; get; }
        public string InfoRecord { set; get; }
        public string PRNumber { set; get; }
        public string CompCode { set; get; }
        public string CompCodeText { set; get; }
        public string Plant { set; get; }
        public string Text { set; get; }
        public string FlagDelete { set; get; }
        public string Name1 { set; get; }
        public string City { set; get; }
        public string DiscPrice { set; get; }
        public string CurrDisc { set; get; }
        public string BBNPrice { set; get; }
        public string CurrBBN { set; get; }
        public string DPPPrice { set; get; }
        public string CurrDPP { set; get; }
        public string DLCPrice { set; get; }
        public string CurrDLC { set; get; }
        public string PPNPrice { set; get; }
        public string CurrPPN { set; get; }
        public string OPTPrice { set; get; }
        public string CurrOPT { set; get; }
        public string Totalpayment { set; get; }
        public string Localprice { set; get; }
        public string DocCurr { set; get; }
        public string ItemText { set; get; }
        public string TextLine { set; get; }
        public string TextLine2 { set; get; }
        public string TextLine3 { set; get; }
        public string TextLine4 { set; get; }
        public string TextLine5 { set; get; }
        public string TextLine6 { set; get; }
        public string TextLine7 { set; get; }
        public string TextLine8 { set; get; }
        public string TextLine9 { set; get; }
        public string TextLine10 { set; get; }
        public string TextLine11 { set; get; }
        public string PRRelDate { set; get; }
        public string Name { set; get; }
        public string ItemDelvDate { set; get; }
        public string ItemDelvDate2 { set; get; }
        public string NetprPurcInfoRec { set; get; }
        public string CurrNetprInfRec { set; get; }
        public string CarDesc { set; get; }
        public string CarModel { set; get; }
        public string CarType { set; get; }
        public string CarBrand { set; get; }
        public string CarTransmisi { set; get; }
        public string CarSeries { set; get; }
        public string CarYear { set; get; }
        public string MatDoc { set; get; }
        public string MatDocYear { set; get; }
        public string MatDocItem { set; get; }
        public string PostDateDoc { set; get; }
        public string PostDateDocBPKB { set; get; }
        public string AccDocNumber { set; get; }
        public string FiscalYear { set; get; }
        public string ClearingDocNumber { set; get; }
        public string ClearingDate { set; get; }
        public string MatDocGI { set; get; }
        public string EquipmentNumb { set; get; }
        public string BatchNumber { set; get; }
        public string SerialNumber { set; get; }
        public string ManSerialNumber { set; get; }
        public string ModelNumber { set; get; }
        public string DateRecordCreated { set; get; }
        public string AssetNumber { set; get; }
        public string CarSTNK { set; get; }
        public string CarRBentuk { set; get; }
        public string DateCarRBentuk { set; get; }
        public string CarFaktur { set; get; }
        public string DateCarFaktur { set; get; }
        public string CarFormA { set; get; }
        public string DateCarFormA { set; get; }
        public string CarSertif { set; get; }
        public string DateCarSertif { set; get; }
        public string CarRegUji { set; get; }
        public string CarBPKB { set; get; }
        public string StatusCarBPKB { set; get; }
        public string RefDocNo { set; get; }
        public string RefKey { set; get; }
        public string PRNo { set; get; }
        public string TGLPRSAP { set; get; }
        public string PRDeliveryDate { set; get; }
        public string RequesterName { set; get; }
        public string PRStatus { set; get; }
        public string PRKaroseri { set; get; }
        public string PRAccessories { set; get; }
        public string ProcessVKaroseri { set; get; }
        public string ProcessVAccs { set; get; }
        public string Customer { set; get; }
        public string OntheRoadPrice { set; get; }
        public string PromiseDeiveryDate { set; get; }
        public string PeriodePO { set; get; }
        public string OfficerName { set; get; }
        public string UnitDeliveryAddress { set; get; }
        public string POStatus { set; get; }
        public string SchedItem { set; get; }
        public string SchedDelvDate { set; get; }
        public string BBN { set; get; }
        public string Color { set; get; }
        public string Year { set; get; }
        public string Gardan { set; get; }
        public string salescontractNo { set; get; }
        public string Salescontractdate { set; get; }
        public string Customername { set; get; }
        public string TextLine12 { set; get; }

        public static List<ProductDetail> GetListProduct(string XMLPathFile)
        {
            var product = new List<ProductDetail>();

            using (var xmlReader = new StreamReader(XMLPathFile))
            {
                var doc = XDocument.Load(xmlReader);
                XNamespace nonamespace = XNamespace.None;
                var result = (from table in doc.Descendants(nonamespace + "Proc_Detail")
                              .Where(item => (DateTime.Parse(item.Element("DateCreated").Value.ToString())).Year 
                              > (DateTime.Now.AddYears(-2)).Year) 
                              select new ProductDetail
                              {
                                  PurcDoc = table.Element("PurcDoc").Value,
                                  ItemDoc = table.Element("ItemDoc").Value,
                                  DateCreated = table.Element("DateCreated").Value,
                                  DocType = table.Element("DocType").Value,
                                  PayTerm = table.Element("PayTerm").Value,
                                  Vendor = table.Element("Vendor").Value,
                                  DocDate = table.Element("DocDate").Value,
                                  ChangeDate = table.Element("ChangeDate").Value,
                                  Material = table.Element("Material").Value,
                                  MatGroup = table.Element("MatGroup").Value,
                                  DocCond = table.Element("DocCond").Value,
                                  NetPrice = table.Element("NetPrice").Value,
                                  Curr = table.Element("Curr").Value,
                                  POQty = table.Element("POQty").Value,
                                  UoM = table.Element("UoM").Value,
                                  InfoRecord = table.Element("InfoRecord").Value,
                                  PRNumber = table.Element("PRNumber").Value,
                                  CompCode = table.Element("CompCode").Value,
                                  CompCodeText = table.Element("CompCodeText").Value,
                                  Plant = table.Element("Plant").Value,
                                  Text = table.Element("Text").Value,
                                  FlagDelete = table.Element("FlagDelete").Value,
                                  Name1 = table.Element("Name1").Value,
                                  City = table.Element("City").Value,
                                  DiscPrice = table.Element("DiscPrice").Value,
                                  CurrDisc = table.Element("CurrDisc").Value,
                                  BBNPrice = table.Element("BBNPrice").Value,
                                  CurrBBN = table.Element("CurrBBN").Value,
                                  DPPPrice = table.Element("DPPPrice").Value,
                                  CurrDPP = table.Element("CurrDPP").Value,
                                  DLCPrice = table.Element("DLCPrice").Value,
                                  CurrDLC = table.Element("CurrDLC").Value,
                                  PPNPrice = table.Element("PPNPrice").Value,
                                  CurrPPN = table.Element("CurrPPN").Value,
                                  OPTPrice = table.Element("OPTPrice").Value,
                                  CurrOPT = table.Element("CurrOPT").Value,
                                  Totalpayment = table.Element("DocPrice").Value,
                                  Localprice = table.Element("LocalPrice").Value,
                                  DocCurr = table.Element("DocCurr").Value,
                                  ItemText = table.Element("ItemText").Value,
                                  TextLine = table.Element("TextLine").Value,
                                  TextLine2 = table.Element("TextLine2").Value,
                                  TextLine3 = table.Element("TextLine3").Value,
                                  TextLine4 = table.Element("TextLine4").Value,
                                  TextLine5 = table.Element("TextLine5").Value,
                                  TextLine6 = table.Element("TextLine6").Value,
                                  TextLine7 = table.Element("TextLine7").Value,
                                  TextLine8 = table.Element("TextLine8").Value,
                                  TextLine9 = table.Element("TextLine9").Value,
                                  TextLine10 = table.Element("TextLine10").Value,
                                  TextLine11 = table.Element("TextLine11").Value,
                                  PRRelDate = table.Element("PRRelDate").Value,
                                  Name = table.Element("Name").Value,
                                  ItemDelvDate = table.Element("ItemDelvDate").Value,
                                  ItemDelvDate2 = table.Element("ItemDelvDate2").Value,
                                  NetprPurcInfoRec = table.Element("NetprPurcInfoRec").Value,
                                  CurrNetprInfRec = table.Element("CurrNetprInfRec").Value,
                                  CarDesc = table.Element("CarDesc").Value,
                                  CarModel = table.Element("CarModel").Value,
                                  CarType = table.Element("CarType").Value,
                                  CarBrand = table.Element("CarBrand").Value,
                                  CarTransmisi = table.Element("CarTransmisi").Value,
                                  CarSeries = table.Element("CarSeries").Value,
                                  CarYear = table.Element("CarYear").Value,
                                  MatDoc = table.Element("MatDoc").Value,
                                  MatDocYear = table.Element("MatDocYear").Value,
                                  MatDocItem = table.Element("MatDocItem").Value,
                                  PostDateDoc = table.Element("PostDateDoc").Value,
                                  PostDateDocBPKB = table.Element("PostDateDocBPKB").Value,
                                  AccDocNumber = table.Element("AccDocNumber").Value,
                                  FiscalYear = table.Element("FiscalYear").Value,
                                  ClearingDocNumber = table.Element("ClearingDocNumber").Value,
                                  ClearingDate = table.Element("ClearingDate").Value,
                                  MatDocGI = table.Element("MatDocGI").Value,
                                  EquipmentNumb = table.Element("EquipmentNumb").Value,
                                  BatchNumber = table.Element("BatchNumber").Value,
                                  SerialNumber = table.Element("SerialNumber").Value,
                                  ManSerialNumber = table.Element("ManSerialNumber").Value,
                                  ModelNumber = table.Element("ModelNumber").Value,
                                  DateRecordCreated = table.Element("DateRecordCreated").Value,
                                  AssetNumber = table.Element("AssetNumber").Value,
                                  CarSTNK = table.Element("CarSTNK").Value,
                                  CarRBentuk = table.Element("CarRBentuk").Value,
                                  DateCarRBentuk = table.Element("DateCarRBentuk").Value,
                                  CarFaktur = table.Element("CarFaktur").Value,
                                  DateCarFaktur = table.Element("DateCarFaktur").Value,
                                  CarFormA = table.Element("CarFormA").Value,
                                  DateCarFormA = table.Element("DateCarFormA").Value,
                                  CarSertif = table.Element("CarSertif").Value,
                                  DateCarSertif = table.Element("DateCarSertif").Value,
                                  CarRegUji = table.Element("CarRegUji").Value,
                                  CarBPKB = table.Element("CarBPKB").Value,
                                  StatusCarBPKB = table.Element("StatusCarBPKB").Value,
                                  RefDocNo = table.Element("RefDocNo").Value,
                                  RefKey = table.Element("RefKey").Value,
                                  PRNo = table.Element("PRNo").Value,
                                  TGLPRSAP = table.Element("PRDate").Value,
                                  PRDeliveryDate = table.Element("PRDeliveryDate").Value,
                                  RequesterName = table.Element("RequesterName").Value,
                                  PRStatus = table.Element("PRStatus").Value,
                                  PRKaroseri = table.Element("PRKaroseri").Value,
                                  PRAccessories = table.Element("PRAccessories").Value,
                                  ProcessVKaroseri = table.Element("ProcessVKaroseri").Value,
                                  ProcessVAccs = table.Element("ProcessVAccs").Value,
                                  Customer = table.Element("Customer").Value,
                                  OntheRoadPrice = table.Element("OntheRoadPrice").Value,
                                  PromiseDeiveryDate = table.Element("PromiseDeiveryDate").Value,
                                  PeriodePO = table.Element("PeriodePO").Value,
                                  OfficerName = table.Element("OfficerName").Value,
                                  UnitDeliveryAddress = table.Element("UnitDeliveryAddress").Value,
                                  POStatus = table.Element("POStatus").Value,
                                  SchedItem = table.Element("SchedItem").Value,
                                  SchedDelvDate = table.Element("SchedDelvDate").Value,
                                  BBN = table.Element("BBN").Value,
                                  Color = table.Element("Color").Value,
                                  Year = table.Element("Year").Value,
                                  Gardan = table.Element("Gardan").Value,
                                  salescontractNo = table.Element("SalesContract").Value,
                                  Salescontractdate = table.Element("SalesContactDate").Value,
                                  Customername = table.Element("CustomerName").Value,
                                  TextLine12 = table.Element("TextLine12").Value,
                              });
                product = result.ToList();
            }
            return product;
        }        
    }
}

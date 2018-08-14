using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Windows.Forms;
using System.Configuration;
using System.Net.Mail;
using System.Data.SqlClient;
using System.Xml;
using System.Data;
using System.Security.Principal;
using System.Text;
using System.Threading.Tasks;
using System.Threading;

namespace ParsingPurchaseOrder
{

    public partial class ParsingPO : Form
    {
        string _User = Encrypt.Decrypt(ConfigurationManager.AppSettings["Username"]);
        string _Pass = Encrypt.Decrypt(ConfigurationManager.AppSettings["Password"]);
        string _Domain = Encrypt.Decrypt(ConfigurationManager.AppSettings["Domain"]);
        string _Path = ConfigurationManager.AppSettings["pathParsing"];
        int f = 0;
        public ParsingPO()
        {
            InitializeComponent();
            try
            {
                List<string> listString = new List<string>();
                do
                {
                    //cek data xml di folder
                    NetworkCredential theNetworkCredential = new NetworkCredential(_User, _Pass, _Domain);
                    CredentialCache theNetcache = new CredentialCache();
                    theNetcache.Add(new Uri(@"\\sera03018"), "Basic", theNetworkCredential);

                    listString = Directory.GetFiles(_Path).ToList();
                    if (listString.Count > 0)
                    {
                        richTextBox1.Text = richTextBox1.Text + "Start read file" + "\n";
                        ReadFilesXML();
                        listString = null;
                        listString = Directory.GetFiles(_Path).ToList();
                    }
                    listString = new List<string>();
                }
                while (listString.Count > 0);

                foreach (var process in Process.GetProcessesByName("ParsingPurchaseOrder"))
                {
                    process.Kill();
                }
            }
            catch (Exception ex)
            {
                WriteLogCatch("" + ex.ToString());
                foreach (var process in Process.GetProcessesByName("ParsingPurchaseOrder"))
                {
                    process.Kill();
                }
            }
        }

        public static void WriteLogSIS(string xml, string message)
        {
            if (ConfigurationManager.AppSettings["AppLogger"].ToLower() != "on") return;

            //nama directory
            string logDateFolder = string.Format("{0}\\{1}-{2}-{3}", ConfigurationManager.AppSettings["LogFolderPathSIS"],
                DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day);

            if (Directory.Exists(logDateFolder) == false) Directory.CreateDirectory(logDateFolder);

            string filename = string.Format("{0}\\{1}.log   ", logDateFolder, xml);

            using (StreamWriter sw = File.Exists(filename) == true ? File.AppendText(filename) : File.CreateText(filename))
            {
                sw.WriteLine(string.Format("{0} {1}", DateTime.Now, message));
                sw.Close();
            }
        }

        public static void WriteLogCatch(string message)
        {
            if (ConfigurationManager.AppSettings["AppLogger"].ToLower() != "on") return;

            //nama directory
            string logDateFolder = string.Format("{0}\\{1}-{2}-{3}", ConfigurationManager.AppSettings["LogFolderPathSIS"],
                DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day);

            if (Directory.Exists(logDateFolder) == false) Directory.CreateDirectory(logDateFolder);

            string filename = string.Format("{0}\\{1}.log", logDateFolder, "catch");

            using (StreamWriter sw = File.Exists(filename) == true ? File.AppendText(filename) : File.CreateText(filename))
            {
                sw.WriteLine(string.Format("{0} {1}", DateTime.Now, message));
                sw.Close();
            }
        }

        public static void WriteLogError(string xml, string message)
        {
            if (ConfigurationManager.AppSettings["AppLogger"].ToLower() != "on") return;
            string logErr = string.Format("{0}\\Error.log", ConfigurationManager.AppSettings["LogErrorFolder"]);

            StreamWriter tx = new StreamWriter(logErr, true, Encoding.UTF8);
            tx.WriteLine(string.Format("{0} {1}", DateTime.Now, message));
            tx.Close();
            GC.Collect();
        }

        public DateTime? convert_date(DateTime? a)
        {
            Convert.ToDateTime(a);
            return a;
        }

        private void ReadFilesXML()
        {
            try
            {
                string a = WindowsIdentity.GetCurrent().Name;
                string[] array1 = Directory.GetFiles(_Path);

                richTextBox1.Text = richTextBox1.Text + "Windows Identity : " + a + "\n";
                List<DateTime> listdate = new List<DateTime>();
                List<String> liststring = new List<String>();

                foreach (string name in array1)
                {
                    int leng = name.Length;
                    string getName = name.Substring(leng - 31, 31);
                    string AllDate = getName.Substring(8, 15);
                    string dates = AllDate.Substring(0, 8);
                    string times = AllDate.Substring(9, 6);

                    DateTime dt = Convert.ToDateTime(dates.Substring(0, 4) + "/" + dates.Substring(4, 2) + "/" + dates.Substring(6, 2));
                    DateTime datetime = dt.AddHours(Convert.ToDouble(times.Substring(0, 2))).AddMinutes(Convert.ToDouble(times.Substring(2, 2))).AddSeconds(Convert.ToDouble(times.Substring(4, 2)));

                    listdate.Add(datetime);
                    liststring.Add(name + "|" + datetime);

                    richTextBox1.Text = richTextBox1.Text + "Filename : " + name + "\n";
                    Console.WriteLine(name);
                    // CR                    
                    String file = name;
                    int index = file.Split('\\').Length - 1;
                    String fileName = file.Split('\\')[index];
                    string sourceFile = file;
                    string destinationBackup = ConfigurationManager.AppSettings["DestinationBackup"] + fileName;
                    richTextBox1.Text = richTextBox1.Text + "Backup Path: " + destinationBackup + "\n";
                    string fileXml = _Path + fileName;

                    string split = fileName.Split('_')[1];
                    string namesXml = split;

                    if (namesXml == "PURCHASEORDER")
                    {
                        richTextBox1.Text = richTextBox1.Text + " XML Filename : " + fileXml + "\n";
                        ParsingXMPFiles(fileXml);
                    }
                    // CR
                }
                /*
                if (array1.Length > 0)
                {
                    DateTime newDate = listdate.Min();
                    String file = liststring.Find(f => f.Split('|')[1] == newDate.ToString()).Split('|')[0];

                    int index = file.Split('\\').Length - 1;
                    String fileName = file.Split('\\')[index];
                    string sourceFile = file;
                    string destinationBackup = @"\\sera03018\DEALERCONNECTOR\archive\apps\PURCHASEORDER\" + fileName;
                    //string destinationBackup = @"\\D\PURCHASEORDER\BACKUP\" + fileName;
                    richTextBox1.Text = richTextBox1.Text + "Backup Path: " + destinationBackup + "\n";
                    string fileXml = _Path + fileName;

                    string split = fileName.Split('_')[1];
                    string namesXml = split;

                    if (namesXml == "PURCHASEORDER")
                    {
                        richTextBox1.Text = richTextBox1.Text + " XML Filename : " + fileXml + "\n";
                        ParsingXMPFiles(fileXml);
                    }
                }
                */
            }
            catch (Exception ex)
            {
                WriteLogCatch(ex.ToString());
                throw ex;
            }
        }

        private void ParsingXMPFiles(String fileName)
        {
            try
            {
                richTextBox1.Text = richTextBox1.Text + "Start parsing file" + "\n";
                //SqlConnection con = null;
                //SqlCommand cmd = null;
                //SqlCommand cmd2 = null;
                //con = new SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings["proc"].ConnectionString);

                //XmlDocument xmldoc = new XmlDocument();
                //xmldoc.Load(fileName);
                string logErr = string.Format("{0}\\Error.Log", ConfigurationManager.AppSettings["LogErrorFolder"]);
                File.WriteAllText(logErr, string.Empty);
                //XmlNodeList nodeList = xmldoc.GetElementsByTagName("Proc_Detail");

                List<DateTime> listdate = new List<DateTime>();
                List<String> liststring = new List<String>();

                //int leng = fileName.Length;
                //string getXMLName = fileName.Substring(leng - 31, 31);
                //string getName = getXMLName.Substring(0, getXMLName.Length - 4);

                int index = fileName.Split('\\').Length - 1;
                string getName = fileName.Split('\\')[index];

                //read xml file using linq 2018-07-20
                List<ProductDetail> ListProduct = new List<ProductDetail>();
                var ListProduct_Split1 = new List<ProductDetail>();
                var ListProduct_Split2 = new List<ProductDetail>();

                //get list product from xml file
                ListProduct = ProductDetail.GetListProduct(fileName);
                ListProduct_Split1 = ListProduct.Where(x=>x.Material != "BPKB").ToList(); //material non bpkb production
                ListProduct_Split2 = ListProduct.Where(x => x.Material == "BPKB").ToList(); //material bpkb production

                /*
                con.Open();
                cmd = new SqlCommand("[SP_PARSINGXML]", con);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd2 = new SqlCommand("[SP_PARSINGXML]", con);
                cmd2.CommandType = CommandType.StoredProcedure;
                */

                /*
                Thread Firstthread = null;
                Firstthread = new Thread(() => ProcessIntoDatabase(ListProduct_Split1, getName));
                Thread Secondthread2 = new Thread(() => ProcessIntoDatabase(ListProduct_Split2, getName));
                Firstthread.Start();
                Secondthread2.Start();
                Firstthread.Join();
                Secondthread2.Join();
                
                Parallel.Invoke(
                    () => ProcessIntoDatabase(ListProduct_Split1, getName),
                    () => ProcessIntoDatabase(ListProduct_Split2, getName));
                */

                Task task1 = Task.Factory.StartNew(() => ProcessIntoDatabase(ListProduct_Split1, getName));
                Task task2 = Task.Factory.StartNew(() => ProcessIntoDatabase(ListProduct_Split2, getName));
                Task.WaitAll(task1, task2);
                string namesxml = string.Empty;
                richTextBox1.Text = richTextBox1.Text + "Start copy file" + "\n";
                SendEmail(logErr, fileName, namesxml);
            }
            catch (Exception ex)
            {
                WriteLogCatch(ex.ToString());
                throw ex;
            }
        }

        private void SendEmail(string logInfo, string fileNames, string subject)
        {
            try
            {
                richTextBox1.Text = richTextBox1.Text + "Start copy file service" + "\n";
                //Add send email log error
                FileInfo info = new FileInfo(logInfo);
                string namaFile = @"" + fileNames + "";

                string nama;

                if (subject == "ReleaseP")
                {
                    var r = namaFile.Substring(namaFile.IndexOf(@"R"));
                    nama = r.Substring(0, r.Length);
                }
                else
                {
                    var r = namaFile.Substring(namaFile.IndexOf(@"P"));
                    nama = r.Substring(8, r.Length - 8);
                }

                var streams = info.OpenText();
                string ErrMessage = streams.ReadToEnd();
                streams.Close();

                richTextBox1.Text = richTextBox1.Text + ErrMessage + "\n";
                if (info.Length > 0)
                {
                    MailMessage message = new MailMessage();
                    SmtpClient smtpC = new SmtpClient();
                    message.From = new MailAddress("no-reply@sera.astra.co.id");
                    message.To.Add("farika.maharani@sera.astra.co.id");
                    message.To.Add("siska.dwi@sera.astra.co.id");
                    message.To.Add("henriy.firman@gmail.com");

                    message.Subject = "Error Log : " + nama;
                    message.Body = ErrMessage;

                    smtpC.Port = 25;
                    smtpC.Host = "webmail.sera.astra.co.id";

                    smtpC.UseDefaultCredentials = true;
                    smtpC.DeliveryMethod = SmtpDeliveryMethod.Network;

                    GC.Collect();

                }
                else
                {
                    int index = fileNames.Split('\\').Length - 1;
                    String fileName = fileNames.Split('\\')[index];

                    string destinationBackup = ConfigurationManager.AppSettings["DestinationBackup"] + fileName;
                    richTextBox1.Text = richTextBox1.Text + "Start copy file in service" + "\n";
                    System.IO.File.Move(fileNames, destinationBackup);
                    richTextBox1.Text = richTextBox1.Text + "End copy file in service" + "\n";

                }
            }

            catch (Exception ex)
            {
                WriteLogCatch(ex.ToString());
                throw ex;
            }
        }

        public async void ProcessIntoDatabase(List<ProductDetail> ListProduct, string getName)
        {
            string colmn = string.Empty;
            //foreach(var product in ListProduct)
            Parallel.ForEach(ListProduct, product =>
            {
                using (var connection = new SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings["proc"].ConnectionString))
                {
                    try
                    {
                        connection.Open();
                        using (var cmd = new SqlCommand("[SP_PARSINGXML]", connection))
                        {
                            try
                            {
                                cmd.CommandType = CommandType.StoredProcedure;
                                colmn = "PurcDoc"; cmd.Parameters.AddWithValue("@PurcDoc", product.PurcDoc); //1
                                colmn = "ItemDoc"; cmd.Parameters.AddWithValue("@ItemDoc", product.ItemDoc); //2
                                colmn = "DateCreated"; cmd.Parameters.AddWithValue("@DateCreated", product.DateCreated == "0000-00-00" ? null : product.DateCreated == "" ? null : convert_date(Convert.ToDateTime(product.DateCreated))); //3
                                colmn = "DocType"; cmd.Parameters.AddWithValue("@DocType", product.DocType); //4
                                colmn = "PayTerm"; cmd.Parameters.AddWithValue("@PayTerm", product.PayTerm == "" ? null : product.PayTerm.ToString().Substring(2, 2)); //5
                                colmn = "Vendor"; cmd.Parameters.AddWithValue("@Vendor", product.Vendor); //6
                                colmn = "DocDate"; cmd.Parameters.AddWithValue("@DocDate", product.DocDate == "0000-00-00" ? null : product.DocDate == "" ? null : convert_date(Convert.ToDateTime(product.DocDate))); //7
                                colmn = "ChangeDate"; cmd.Parameters.AddWithValue("@ChangeDate", product.ChangeDate == "0000-00-00" ? null : product.ChangeDate == "" ? null : convert_date(Convert.ToDateTime(product.ChangeDate))); //8
                                colmn = "Material"; cmd.Parameters.AddWithValue("@Material", product.Material); //9
                                colmn = "MatGroup"; cmd.Parameters.AddWithValue("@MatGroup", product.MatGroup); //10
                                colmn = "DocCond"; cmd.Parameters.AddWithValue("@DocCond", product.DocCond); //11
                                colmn = "NetPrice"; cmd.Parameters.AddWithValue("@NetPrice", product.NetPrice == ".01" ? 0.01 : Convert.ToDouble(product.NetPrice)); //12
                                colmn = "Curr"; cmd.Parameters.AddWithValue("@Curr", product.Curr); //13
                                colmn = "POQty"; cmd.Parameters.AddWithValue("@POQty", Convert.ToDouble(product.POQty)); //14
                                colmn = "UoM"; cmd.Parameters.AddWithValue("@UoM", product.UoM); //15
                                colmn = "InfoRecord"; cmd.Parameters.AddWithValue("@InfoRecord", product.InfoRecord); //16                            
                                colmn = "PRNumber"; cmd.Parameters.AddWithValue("@PRNumber", product.PRNumber == "" ? "kosong" : product.PRNumber); //17
                                colmn = "CompCode"; cmd.Parameters.AddWithValue("@CompCode", product.CompCode); //18
                                colmn = "CompName"; cmd.Parameters.AddWithValue("@CompName", product.CompCodeText); //19
                                colmn = "Plant"; cmd.Parameters.AddWithValue("@Plant", product.Plant); //20
                                colmn = "Text"; cmd.Parameters.AddWithValue("@Text", product.Text); //21
                                colmn = "FlagDelete"; cmd.Parameters.AddWithValue("@FlagDelete", product.FlagDelete); //22
                                colmn = "Name1"; cmd.Parameters.AddWithValue("@Name1", product.Name1); //23
                                colmn = "City"; cmd.Parameters.AddWithValue("@City", product.City); //24
                                colmn = "DiscPrice"; cmd.Parameters.AddWithValue("@DiscPrice", product.DiscPrice == "" ? null : product.DiscPrice); //25
                                colmn = "CurrDisc"; cmd.Parameters.AddWithValue("@CurrDisc", product.CurrDisc); //26
                                colmn = "BBNPrice"; cmd.Parameters.AddWithValue("@BBNPrice", product.BBNPrice); //27
                                colmn = "CurrBBN"; cmd.Parameters.AddWithValue("@CurrBBN", product.CurrBBN); //28
                                colmn = "DPPPrice"; cmd.Parameters.AddWithValue("@DPPPrice", product.DPPPrice); //29
                                colmn = "CurrDPP"; cmd.Parameters.AddWithValue("@CurrDPP", product.CurrDPP); //30
                                colmn = "DLCPrice"; cmd.Parameters.AddWithValue("@DLCPrice", product.DLCPrice); //31
                                colmn = "CurrDLC"; cmd.Parameters.AddWithValue("@CurrDLC", product.CurrDLC); //32
                                colmn = "PPNPrice"; cmd.Parameters.AddWithValue("@PPNPrice", product.PPNPrice); //33
                                colmn = "CurrPPN"; cmd.Parameters.AddWithValue("@CurrPPN", product.CurrPPN); //34
                                colmn = "OPTPrice"; cmd.Parameters.AddWithValue("@OPTPrice", product.OPTPrice); //35
                                colmn = "CurrOPT"; cmd.Parameters.AddWithValue("@CurrOPT", product.CurrOPT); //36
                                colmn = "Totalpayment"; cmd.Parameters.AddWithValue("@Totalpayment", product.Totalpayment); //37
                                colmn = "Localprice"; cmd.Parameters.AddWithValue("@Localprice", product.Localprice); //38
                                colmn = "Currency"; cmd.Parameters.AddWithValue("@Currency", product.DocCurr); //39
                                colmn = "ItemText"; cmd.Parameters.AddWithValue("@ItemText", product.ItemText); //40
                                colmn = "TextLine"; cmd.Parameters.AddWithValue("@TextLine", product.TextLine); //41
                                colmn = "TextLine2"; cmd.Parameters.AddWithValue("@TextLine2", product.TextLine2); //42
                                colmn = "TextLine3"; cmd.Parameters.AddWithValue("@TextLine3", product.TextLine3); //43
                                colmn = "TextLine4"; cmd.Parameters.AddWithValue("@TextLine4", product.TextLine4); //44
                                colmn = "TextLine5"; cmd.Parameters.AddWithValue("@TextLine5", product.TextLine5); //45
                                colmn = "TextLine6"; cmd.Parameters.AddWithValue("@TextLine6", product.TextLine6); //46
                                colmn = "TextLine7"; cmd.Parameters.AddWithValue("@TextLine7", product.TextLine7); //47
                                colmn = "TextLine8"; cmd.Parameters.AddWithValue("@TextLine8", product.TextLine8); //48
                                colmn = "TextLine9"; cmd.Parameters.AddWithValue("@TextLine9", product.TextLine9); //49
                                colmn = "TextLine10"; cmd.Parameters.AddWithValue("@TextLine10", product.TextLine10); //50
                                colmn = "TextLine11"; cmd.Parameters.AddWithValue("@TextLine11", product.TextLine11); //51
                                colmn = "PRRelDate"; cmd.Parameters.AddWithValue("@PRRelDate", product.PRRelDate == "0000-00-00" ? null : product.PRRelDate == "" ? null : convert_date(Convert.ToDateTime(product.PRRelDate))); //52
                                colmn = "Name"; cmd.Parameters.AddWithValue("@Name", product.Name); //53
                                colmn = "ItemDelvDate"; cmd.Parameters.AddWithValue("@ItemDelvDate", product.ItemDelvDate == "0000-00-00" ? null : product.ItemDelvDate == "" ? null : convert_date(Convert.ToDateTime(product.ItemDelvDate))); //54
                                colmn = "ItemDelvDate2"; cmd.Parameters.AddWithValue("@ItemDelvDate2", product.ItemDelvDate2 == "0000-00-00" ? null : product.ItemDelvDate2 == "" ? null : convert_date(Convert.ToDateTime(product.ItemDelvDate2))); //55
                                colmn = "NetprPurcInfoRec"; cmd.Parameters.AddWithValue("@NetprPurcInfoRec", product.NetprPurcInfoRec); //56
                                colmn = "CurrNetprInfRec"; cmd.Parameters.AddWithValue("@CurrNetprInfRec", product.CurrNetprInfRec); //57
                                colmn = "CarDesc"; cmd.Parameters.AddWithValue("@CarDesc", product.CarDesc); //58
                                colmn = "CarModel"; cmd.Parameters.AddWithValue("@CarModel", product.CarModel); //59
                                colmn = "CarType"; cmd.Parameters.AddWithValue("@CarType", product.CarType); //60
                                colmn = "CarBrand"; cmd.Parameters.AddWithValue("@CarBrand", product.CarBrand); //61
                                colmn = "CarTransmisi"; cmd.Parameters.AddWithValue("@CarTransmisi", product.CarTransmisi); //62
                                colmn = "CarSeries"; cmd.Parameters.AddWithValue("@CarSeries", product.CarSeries); //63
                                colmn = "CarYear"; cmd.Parameters.AddWithValue("@CarYear", product.CarYear); //64
                                colmn = "MatDoc"; cmd.Parameters.AddWithValue("@MatDoc", product.MatDoc); //65
                                colmn = "MatDocYear"; cmd.Parameters.AddWithValue("@MatDocYear", product.MatDocYear); //66
                                colmn = "MatDocItem"; cmd.Parameters.AddWithValue("@MatDocItem", product.MatDocItem == "" ? null : product.MatDocItem); //67
                                colmn = "PostDateDoc"; cmd.Parameters.AddWithValue("@PostDateDoc", product.PostDateDoc == "0000-00-00" ? null : product.PostDateDoc == "" ? null : convert_date(Convert.ToDateTime(product.PostDateDoc))); //68
                                colmn = "PostDateDocBPKB"; cmd.Parameters.AddWithValue("@PostDateDocBPKB", product.PostDateDocBPKB == "0000-00-00" ? null : product.PostDateDocBPKB == "" ? null : convert_date(Convert.ToDateTime(product.PostDateDocBPKB))); //69
                                colmn = "AccDocNumber"; cmd.Parameters.AddWithValue("@AccDocNumber", product.AccDocNumber); //70
                                colmn = "FiscalYear"; cmd.Parameters.AddWithValue("@FiscalYear", product.FiscalYear); //71
                                colmn = "ClearingDocNumber"; cmd.Parameters.AddWithValue("@ClearingDocNumber", product.ClearingDocNumber); //72
                                colmn = "ClearingDate"; cmd.Parameters.AddWithValue("@ClearingDate", product.ClearingDate == "0000-00-00" ? null : product.ClearingDate == "" ? null : convert_date(Convert.ToDateTime(product.ClearingDate))); //73
                                colmn = "MatDocGI"; cmd.Parameters.AddWithValue("@MatDocGI", product.MatDocGI); //74
                                colmn = "EquipmentNumb"; cmd.Parameters.AddWithValue("@EquipmentNumb", product.EquipmentNumb); //75
                                colmn = "BatchNumber"; cmd.Parameters.AddWithValue("@BatchNumber", product.BatchNumber); //76
                                colmn = "SerialNumber"; cmd.Parameters.AddWithValue("@SerialNumber", product.SerialNumber); //77
                                colmn = "ManSerialNumber"; cmd.Parameters.AddWithValue("@ManSerialNumber", product.ManSerialNumber); //78
                                colmn = "ModelNumber"; cmd.Parameters.AddWithValue("@ModelNumber", product.ModelNumber); //79
                                colmn = "DateRecordCreated"; cmd.Parameters.AddWithValue("@DateRecordCreated", product.DateRecordCreated == "0000-00-00" ? null : product.DateRecordCreated == "" ? null : convert_date(Convert.ToDateTime(product.DateRecordCreated))); //80
                                colmn = "AssetNumber"; cmd.Parameters.AddWithValue("@AssetNumber", product.AssetNumber); //81
                                colmn = "CarSTNK"; cmd.Parameters.AddWithValue("@CarSTNK", product.CarSTNK == "0000-00-00" ? null : product.CarSTNK == "0018-02-11" ? convert_date(Convert.ToDateTime("2018-02-11")) : product.CarSTNK == "" ? null : convert_date(Convert.ToDateTime(product.CarSTNK))); //82
                                colmn = "CarRBentuk"; cmd.Parameters.AddWithValue("@CarRBentuk", product.CarRBentuk); //83
                                colmn = "DateCarRBentuk"; cmd.Parameters.AddWithValue("@DateCarRBentuk", product.DateCarRBentuk == "0000-00-00" ? null : product.DateCarRBentuk == "" ? null : convert_date(Convert.ToDateTime(product.DateCarRBentuk))); //84
                                colmn = "CarFaktur"; cmd.Parameters.AddWithValue("@CarFaktur", product.CarFaktur); //85
                                colmn = "DateCarFaktur"; cmd.Parameters.AddWithValue("@DateCarFaktur", product.DateCarFaktur == "0000-00-00" ? null : product.DateCarFaktur == "" ? null : convert_date(Convert.ToDateTime(product.DateCarFaktur))); //86
                                colmn = "CarFormA"; cmd.Parameters.AddWithValue("@CarFormA", product.CarFormA); //87
                                colmn = "DateCarFormA"; cmd.Parameters.AddWithValue("@DateCarFormA", product.DateCarFormA == "0000-00-00" ? null : product.DateCarFormA == "" ? null : convert_date(Convert.ToDateTime(product.DateCarFormA))); //88
                                colmn = "CarSertif"; cmd.Parameters.AddWithValue("@CarSertif", product.CarSertif); //89
                                colmn = "DateCarSertif"; cmd.Parameters.AddWithValue("@DateCarSertif", product.DateCarSertif == "0000-00-00" ? null : product.DateCarSertif == "" ? null : convert_date(Convert.ToDateTime(product.DateCarSertif))); //90
                                colmn = "CarRegUji"; cmd.Parameters.AddWithValue("@CarRegUji", product.CarRegUji); //91
                                colmn = "CarBPKB"; cmd.Parameters.AddWithValue("@CarBPKB", product.CarBPKB); //92
                                colmn = "StatusCarBPKB"; cmd.Parameters.AddWithValue("@StatusCarBPKB", product.StatusCarBPKB); //93
                                colmn = "RefDocNo"; cmd.Parameters.AddWithValue("@RefDocNo", product.RefDocNo); //94
                                colmn = "RefKey"; cmd.Parameters.AddWithValue("@RefKey", product.RefKey); //95
                                colmn = "PRNo"; cmd.Parameters.AddWithValue("@PRSAP", product.PRNo); //96
                                colmn = "TGLPRSAP"; cmd.Parameters.AddWithValue("@PRDate", product.PRDeliveryDate == "0000-00-00" ? null : product.TGLPRSAP == "" ? null : convert_date(Convert.ToDateTime(product.TGLPRSAP))); //97

                                colmn = "PRDeliveryDate"; cmd.Parameters.AddWithValue("@PRDeliveryDate", product.PRDeliveryDate == "0000-00-00" ? null : product.PRDeliveryDate == "" ? null : convert_date(Convert.ToDateTime(product.PRDeliveryDate))); //98
                                colmn = "RequesterName"; cmd.Parameters.AddWithValue("@RequesterName", product.RequesterName); //99
                                colmn = "PRStatus"; cmd.Parameters.AddWithValue("@PRStatus", product.PRStatus); //100
                                colmn = "PRKaroseri"; cmd.Parameters.AddWithValue("@PRKaroseri", product.PRKaroseri); //101
                                colmn = "PRAccessories"; cmd.Parameters.AddWithValue("@PRAccessories", product.PRAccessories); //102
                                colmn = "ProcessVKaroseri"; cmd.Parameters.AddWithValue("@ProcessVKaroseri", product.ProcessVKaroseri); //103
                                colmn = "ProcessVAccs"; cmd.Parameters.AddWithValue("@ProcessVAccs", product.ProcessVAccs); //104
                                colmn = "Customer"; cmd.Parameters.AddWithValue("@Customer", product.Customer); //105
                                colmn = "OntheRoadPrice"; cmd.Parameters.AddWithValue("@OntheRoadPrice", product.OntheRoadPrice == ".01" ? 0.01 : Convert.ToDouble(product.OntheRoadPrice)); //106
                                colmn = "PromiseDeiveryDate"; cmd.Parameters.AddWithValue("@PromiseDeiveryDate", product.PromiseDeiveryDate == "0000-00-00" ? null : product.PromiseDeiveryDate == "" ? null : convert_date(Convert.ToDateTime(product.PromiseDeiveryDate))); //107
                                colmn = "PeriodePO"; cmd.Parameters.AddWithValue("@PeriodePO", product.PeriodePO); //108
                                colmn = "OfficerName"; cmd.Parameters.AddWithValue("@OfficerName", product.OfficerName); //109
                                colmn = "UnitDeliveryAddress"; cmd.Parameters.AddWithValue("@UnitDeliveryAddress", product.UnitDeliveryAddress); //110
                                colmn = "POStatus"; cmd.Parameters.AddWithValue("@POStatus", product.POStatus); //111
                                colmn = "SchedItem"; cmd.Parameters.AddWithValue("@SchedItem", product.SchedItem == "" ? null : product.SchedItem);  //112
                                colmn = "SchedDelvDate"; cmd.Parameters.AddWithValue("@SchedDelvDate", product.SchedDelvDate); //113
                                colmn = "BBN"; cmd.Parameters.AddWithValue("@BBN", product.BBN); //114
                                colmn = "Color"; cmd.Parameters.AddWithValue("@Color", product.Color); //115
                                colmn = "Year"; cmd.Parameters.AddWithValue("@Year", product.Year); //116
                                colmn = "Gardan"; cmd.Parameters.AddWithValue("@Gardan", product.Gardan); //117
                                colmn = "salescontractNo"; cmd.Parameters.AddWithValue("@salescontractNo", product.salescontractNo); //118
                                colmn = "Salescontractdate"; cmd.Parameters.AddWithValue("@Salescontractdate", product.Salescontractdate == "0000-00-00" ? null : product.Salescontractdate == "" ? null : convert_date(Convert.ToDateTime(product.Salescontractdate))); //119
                                colmn = "Customername"; cmd.Parameters.AddWithValue("@Customername", product.Customername); //120
                                colmn = "TextLine12"; cmd.Parameters.AddWithValue("@TextLine12", product.TextLine12); //121
                                colmn = "NamaFileXLMLastUpdate"; cmd.Parameters.AddWithValue("@NamaFileXLMLastUpdate", getName); //122
                                colmn = "StatusLog"; cmd.Parameters.AddWithValue("@StatusLog", "SUCCESS"); //123
                                cmd.ExecuteScalar();
                            }
                            catch (Exception e)
                            {

                                if (cmd.Parameters.Contains("@StatusLog"))
                                {
                                    cmd.Parameters.RemoveAt("@StatusLog");
                                }                               
                                cmd.Parameters.AddWithValue("@StatusLog", "FAIL");
                                cmd.Parameters.AddWithValue("@Message", colmn + ": " + e.Message);
                                if (connection.State != ConnectionState.Open)
                                {
                                    connection.Open();
                                }
                                cmd.ExecuteScalar();
                            }
                            finally
                            {
                                connection.Close();
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        WriteLogCatch(ex.Message.ToString());
                    }
                }
            });
        }
    }
}

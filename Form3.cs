using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using Newtonsoft.Json.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using HMC_Crypto.HMCCrypto;

namespace WindowsFormsApplication4
{
    public partial class Form3 : Form
    {
        HMC_Crypto.HMCCrypto.UseEncrypt crypt;
        private int CountLoop = 0;
        private SAPbobsCOM.Company oComp;
        private bool OpenAuto = false;

        IList<Post> searchResults = new List<Post>();
        public Form3()
        {
            InitializeComponent();
            this.groupBox1.Hide();
            splitContainer1.SplitterDistance = 0;
            comboBox1.SelectedIndex = 1;

        }

        private void button1_Click(object sender, EventArgs e)
        {

            textBox4.Text = crypt.EncodeSymmetric("HMCBOT", textBox4.Text);
            txtPxPass.Text = crypt.EncodeSymmetric("PxHMCBOT", txtPxPass.Text);


            System.Configuration.Configuration configSave = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            configSave.AppSettings.Settings["Server"].Value = textBox1.Text;
            configSave.AppSettings.Settings["DbServerType"].Value = comboBox1.Text;
            configSave.AppSettings.Settings["CompanyDB"].Value = textBox2.Text;
            configSave.AppSettings.Settings["UserName"].Value = textBox3.Text;
            configSave.AppSettings.Settings["Password"].Value = textBox4.Text;
            configSave.AppSettings.Settings["LicenseServer"].Value = textBox5.Text;
            configSave.AppSettings.Settings["LogPath"].Value = textBox6.Text;
            configSave.AppSettings.Settings["BOT-API-URL"].Value = textBox7.Text;
            configSave.AppSettings.Settings["PxUsername"].Value = txtPxUser.Text;
            configSave.AppSettings.Settings["PxPassword"].Value = txtPxPass.Text;
            configSave.AppSettings.Settings["PxIP"].Value = txtPxIP.Text;
            configSave.AppSettings.Settings["PxPort"].Value = txtPxPort.Text;
            if (checkBox1.Checked)
            {
                configSave.AppSettings.Settings["chk"].Value = "1";
            }
            else
            {
                configSave.AppSettings.Settings["chk"].Value = "0";
            }

            if (chkBox2.Checked)
            {
                configSave.AppSettings.Settings["chk2"].Value = "1";
            }
            else
            {
                configSave.AppSettings.Settings["chk2"].Value = "0";
            }

            configSave.Save();

            System.Configuration.Configuration configRefresh = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            textBox1.Text = configRefresh.AppSettings.Settings["Server"].Value;
            comboBox1.Text = configRefresh.AppSettings.Settings["DbServerType"].Value;
            textBox2.Text = configRefresh.AppSettings.Settings["CompanyDB"].Value;
            textBox3.Text = configRefresh.AppSettings.Settings["UserName"].Value;
            //textBox4.Text = crypt.DecodeSymmetric("HMCBOT", ConfigurationManager.AppSettings["Password"]);
            textBox4.Text = crypt.DecodeSymmetric("HMCBOT", configRefresh.AppSettings.Settings["Password"].Value);
            textBox5.Text = configRefresh.AppSettings.Settings["LicenseServer"].Value;
            textBox6.Text = configRefresh.AppSettings.Settings["LogPath"].Value;
            textBox7.Text = configRefresh.AppSettings.Settings["BOT-API-URL"].Value;
            txtPxUser.Text = configRefresh.AppSettings.Settings["PxUsername"].Value;
            txtPxPass.Text = crypt.DecodeSymmetric("PxHMCBOT", configRefresh.AppSettings.Settings["PxPassword"].Value);
            txtPxIP.Text = configRefresh.AppSettings.Settings["PxIP"].Value;
            txtPxPort.Text = configRefresh.AppSettings.Settings["PxPort"].Value;


            FileStream LogFile = new FileStream(string.Format("{0}\\EXC_{1}.log", textBox6.Text, DateTime.Now.ToString("yyyy-MM-dd")), FileMode.Append);
            StreamWriter StreamWrite = new StreamWriter(LogFile);
            StreamWrite.WriteLine(DateTime.Now.ToLongTimeString() + " Save LogPath");
            StreamWrite.Close();
            LogFile.Close();
            ShowMessage("Save completed");
            ConnectSAP();
        }

        private void Form3_Load(object sender, EventArgs e)
        {

            crypt = new UseEncrypt();

            textBox1.Text = ConfigurationManager.AppSettings["Server"];
            comboBox1.Text = ConfigurationManager.AppSettings["DbServerType"];
            textBox2.Text = ConfigurationManager.AppSettings["CompanyDB"];
            textBox3.Text = ConfigurationManager.AppSettings["UserName"];
            textBox4.Text = crypt.DecodeSymmetric("HMCBOT", ConfigurationManager.AppSettings["Password"]);

            textBox5.Text = ConfigurationManager.AppSettings["LicenseServer"];
            textBox6.Text = ConfigurationManager.AppSettings["LogPath"];
            textBox7.Text = ConfigurationManager.AppSettings["BOT-API-URL"];
            if (textBox6.Text == "")
                textBox6.Text = Application.StartupPath;
            if (textBox7.Text == "")
                textBox7.Text = "https://iapi.bot.or.th/Stat/Stat-ExchangeRate/DAILY_AVG_EXG_RATE_V1/";
            txtPxUser.Text = ConfigurationManager.AppSettings["PxUsername"];
            txtPxPass.Text = crypt.DecodeSymmetric("PxHMCBOT", ConfigurationManager.AppSettings["PxPassword"]);
            txtPxIP.Text = ConfigurationManager.AppSettings["PxIP"];
            txtPxPort.Text = ConfigurationManager.AppSettings["PxPort"];

            if (ConfigurationManager.AppSettings["chk"] == "1")
            {
                checkBox1.Checked = true;
            }
            else
            {
                checkBox1.Checked = false;
            }

            if (ConfigurationManager.AppSettings["chk2"] == "1")
            {
                chkBox2.Checked = true;  
            }
            else
            {
                chkBox2.Checked = false;
            }
            dateTimePicker1.Value = DateTime.Now.Date;
            dateTimePicker2.Value = DateTime.Now.Date;

            if (System.Environment.GetCommandLineArgs().Count() > 1)
            {
                OpenAuto = true;
                this.Hide();
                //MessageBox.Show(System.Environment.GetCommandLineArgs().GetValue(1).ToString());
            }
            ConnectSAP();
            LoadData(true, null, null);
            if (OpenAuto)
            {
                SAPbobsCOM.SBObob oSBObob;
                SAPbobsCOM.Recordset oRecordSet;

                oSBObob = oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);

                SAPbobsCOM.Recordset RecSet;
                RecSet = oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string sqlQuery = string.Format("SELECT DISTINCT T0.[U_MAPCurrency] FROM [dbo].[@CURRENCYMAPPING]  T0");
                RecSet.DoQuery(sqlQuery);
                while (!RecSet.EoF)
                {
                    string GlobalCurrencyCode = RecSet.Fields.Item("U_MAPCurrency").Value;
                    comboBox2.Text = GlobalCurrencyCode;
                    button2_Click(sender, e);
                    button3_Click(sender, e);

                    RecSet.MoveNext();
                }
                this.Close();
            }
        }

        private bool SetCurrencyRate(string CurrCode, System.DateTime CurrDate, System.DateTime WebDate, Post ExchData)
        {
            bool result = false;

            SAPbobsCOM.SBObob oSBObob;
            SAPbobsCOM.Recordset oRecordSet;

            oSBObob = oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);

            SAPbobsCOM.Recordset RecSet;
            RecSet = oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string sqlQuery = string.Format("SELECT *  FROM [dbo].[@CURRENCYMAPPING]  T0 WHERE T0.[U_MAPCurrency] = '{0}'", comboBox2.Text);
            RecSet.DoQuery(sqlQuery);
            while (!RecSet.EoF)
            {
                string B1CurrencyCode = RecSet.Fields.Item("U_Currency").Value;
                string B1Operator = RecSet.Fields.Item("U_Operator").Value;
                string B1RateType = RecSet.Fields.Item("U_UpdateType").Value;
                double B1Factor = RecSet.Fields.Item("U_Factor").Value;

                double B1CalcRate = 0.0;
                double ExRate = 0.0;
                switch (B1RateType)
                {
                    case "1":
                        ExRate = ExchData.selling.Value;
                        break;
                    case "2":
                        ExRate = ExchData.buying_transfer.Value;
                        break;
                    case "3":
                        ExRate = ExchData.buying_sight.Value;
                        break;
                    case "4":
                        ExRate = ExchData.mid_rate.Value;
                        break;
                }
                if (B1Operator == "*")
                {
                    B1CalcRate = ExRate * B1Factor;
                }
                else if (B1Operator == "/")
                {
                    B1CalcRate = ExRate / B1Factor;
                }


                try
                {
                    oSBObob.SetCurrencyRate(B1CurrencyCode, CurrDate.Date, B1CalcRate, true);


                    SAPbobsCOM.UserTable oUserTable;
                    oUserTable = oComp.UserTables.Item("LOGEXCHANGERATE");

                    SAPbobsCOM.Recordset RecSet2;
                    RecSet2 = oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    string sqlQuery2 = string.Format("SELECT Isnull(MAX(CAST(T1.Code AS int)),0) MaxNo FROM [dbo].[@LOGEXCHANGERATE]  T1");
                    RecSet2.DoQuery(sqlQuery2);

                    int TryAttemp = 0;
                    int AutoNo = (int)RecSet2.Fields.Item("MaxNo").Value;

                    InsertData:

                    //Set default, mandatory fields
                    oUserTable.Code = AutoNo.ToString();
                    oUserTable.Name = AutoNo.ToString();
                    //Set user field
                    oUserTable.UserFields.Fields.Item("U_LoadDate").Value = DateTime.Now.Date;
                    oUserTable.UserFields.Fields.Item("U_FromWebDate").Value = ExchData.CallDate;
                    oUserTable.UserFields.Fields.Item("U_WEBCurrency").Value = CurrCode.ToString();
                    oUserTable.UserFields.Fields.Item("U_Currency").Value = B1CurrencyCode.ToString();
                    oUserTable.UserFields.Fields.Item("U_RateDate").Value = WebDate.Date;
                    oUserTable.UserFields.Fields.Item("U_Rate").Value = B1CalcRate.ToString();

                    //if (ret <> 0)
                    //{
                    int ret;
                    // int AutoNum = RecSet2.RecordCount;
                    ret = oUserTable.Add();
                    if (ret != 0)
                    {
                        if (TryAttemp < 10)
                        {
                            TryAttemp++;
                            AutoNo++;
                            goto InsertData;


                        }
                        else
                        {
                            ShowMessage(oComp.GetLastErrorDescription());
                        }
                    }
                    result = true;
                }
                catch (Exception ex)
                {
                    ShowMessage(string.Format("{0} {1}", B1CurrencyCode, ex.Message));
                    result = false;
                }
                RecSet.MoveNext();
            }
            return result;
        }

        private void ShowMessage(string MsgLog)
        {
            if (OpenAuto)
            {
                FileStream FS = new FileStream(string.Format("{0}\\EXC_{1}.log", textBox6.Text, DateTime.Now.ToString("dd-MM-yyyy")), FileMode.Append, FileAccess.Write);
                StreamWriter SW = new StreamWriter(FS);
                SW.WriteLine(string.Format("{0}  {1}", DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"), MsgLog));
                SW.Close();
                FS.Close();
            }
            else
            {
                MessageBox.Show(MsgLog, "Result");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            searchResults.Clear();
            dataGridView1.AutoGenerateColumns = false;
            dataGridView1.DataSource = null;
            CountLoop = 0;
            LoadData(false, dateTimePicker1.Value, dateTimePicker2.Value);

        }
        private void LoadData(bool GetCurrentCode, DateTime? ParamDateFrom, DateTime? ParamDateTo)

        {
            DateTime tmpDateFrom = ParamDateFrom == null ? dateTimePicker1.Value : ParamDateFrom.Value;
            DateTime tmpDateTo = ParamDateTo == null ? dateTimePicker2.Value : ParamDateTo.Value;
            string URL = textBox7.Text;  //  "<!--https://iapi.bot.or.th/Stat/Stat-ExchangeRate/DAILY_AVG_EXG_RATE_V1/-->"
            WebClient wc = new WebClient();

            if (checkBox1.Checked)
            {
                try
                {
                    WebProxy myProxy = new WebProxy();
                    Uri newUri = new Uri(string.Format("{0}:{1}", txtPxIP.Text, txtPxPort.Text));
                    myProxy.Address = newUri;
                    if (chkBox2.Checked)
                    {
                        try
                        {
                            myProxy.Credentials = new NetworkCredential(txtPxUser.Text, txtPxUser.Text);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(string.Format("Load Data Error [{0}].", ex.Message));
                            return;
                        }
                    }
                    wc.Proxy = myProxy;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(string.Format("Load Data Error [{0}].", ex.Message));
                    return;
                }

            }

            //wc.Proxy
            wc.Headers.Add("api-key", "U9G1L457H6DCugT7VmBaEacbHV9RX0PySO05cYaGsm");
            wc.QueryString.Add("start_period", tmpDateFrom.ToString("yyyy-MM-dd", new System.Globalization.CultureInfo("en-US")));
            wc.QueryString.Add("end_period", tmpDateTo.ToString("yyyy-MM-dd", new System.Globalization.CultureInfo("en-US")));

            if (GetCurrentCode)
            {
                try
                { // wc.QueryString.Add("currency", comboBox2.Text);
                    string json = wc.DownloadString(URL);
                    JObject dataJson = JObject.Parse(json);

                    IList<JToken> results = dataJson["result"]["data"]["data_detail"].Children().ToList();
                    IList<Post> searchResults = new List<Post>();
                    foreach (JToken result in results)
                    {
                        wc.QueryString.Add("currency", comboBox2.Text);
                        Post searchResult = result.ToObject<Post>();
                        comboBox2.Items.Add(searchResult.currency_id);
                    }


                }
                catch (Exception ex)
                {
                    MessageBox.Show(string.Format("Load Data Error [{0}].", ex.Message));
                }
            }
            else
            {
                wc.QueryString.Remove("start_period");
                wc.QueryString.Remove("end_period");
                DateTime tmpDateCurrent;
                for (int i = 0; i <= (dateTimePicker2.Value.Date - dateTimePicker1.Value.Date).TotalDays; i++)
                {
                    wc.QueryString.Remove("start_period");
                    wc.QueryString.Remove("end_period");
                    wc.QueryString.Remove("currency");
                    tmpDateFrom = dateTimePicker1.Value.AddDays(i);
                    tmpDateCurrent = dateTimePicker1.Value.AddDays(i).AddDays(-1);
                    int iloop = i + 1;
                    //wc.QueryString.Add("start_period", dateTimePicker1.Value.AddDays(-1).ToString("yyyy-MM-dd", new System.Globalization.CultureInfo("en-US")));
                    //wc.QueryString.Add("end_period", dateTimePicker1.Value.AddDays(-1).ToString("yyyy-MM-dd", new System.Globalization.CultureInfo("en-US")));

                    wc.QueryString.Add("start_period", tmpDateCurrent.ToString("yyyy-MM-dd", new System.Globalization.CultureInfo("en-US")));
                    wc.QueryString.Add("end_period", tmpDateCurrent.ToString("yyyy-MM-dd", new System.Globalization.CultureInfo("en-US")));
                    wc.QueryString.Add("currency", comboBox2.Text);

                    try
                    {
                        string json = wc.DownloadString(URL);
                        Console.WriteLine(json);
                        JObject dataJson = JObject.Parse(json);

                        IList<JToken> results = dataJson["result"]["data"]["data_detail"].Children().ToList();

                        foreach (JToken result in results)
                        {
                            Post searchResult = result.ToObject<Post>();
                            while (searchResult.buying_sight == null)
                            {

                                wc.QueryString.Remove("start_period");
                                wc.QueryString.Remove("end_period");
                                wc.QueryString.Remove("currency");

                                tmpDateFrom = tmpDateFrom.AddDays(-1);
                                wc.QueryString.Add("start_period", tmpDateFrom.ToString("yyyy-MM-dd", new System.Globalization.CultureInfo("en-US")));
                                wc.QueryString.Add("end_period", tmpDateFrom.ToString("yyyy-MM-dd", new System.Globalization.CultureInfo("en-US")));

                                wc.QueryString.Add("currency", comboBox2.Text);

                                json = wc.DownloadString(URL);
                                Console.WriteLine(json);
                                dataJson = JObject.Parse(json);

                                results = dataJson["result"]["data"]["data_detail"].Children().ToList();

                                searchResult = results[0].ToObject<Post>();

                            }

                            searchResult.CallDate = dateTimePicker1.Value.AddDays(i);

                            searchResults.Add(searchResult);
                        }
                    }
                    catch (Exception ex)
                    {

                        MessageBox.Show(string.Format("Load Data Error [{0}].", ex.Message));
                    }

                }
                dataGridView1.DataSource = searchResults;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            bool retCode = false;
            List<Post> PostDataRow = (List<Post>)dataGridView1.DataSource;
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                DataGridViewRow dgRow = dataGridView1.Rows[i];
                DateTime dt = Convert.ToDateTime(dgRow.Cells["cCallDate"].Value);
                DateTime dtWeb = Convert.ToDateTime(dgRow.Cells["cCurrDate"].Value);
                retCode = SetCurrencyRate(dgRow.Cells["cCurrCode"].Value.ToString(), dt, dtWeb, PostDataRow[i]);
            }
            if (retCode)
                ShowMessage("Sync finished");
        }

        private void ConnectSAP()
        {
            if (oComp == null)
                oComp = new SAPbobsCOM.Company();
            else if (oComp.Connected)
                oComp.Disconnect();

            oComp.Server = textBox1.Text;

            switch (comboBox1.Text)
            {
                case "MSSQL2008":
                    oComp.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2008;
                    break;
                case "MSSQL2012":
                    oComp.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2012;
                    break;
                case "MSSQL2014":
                    oComp.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2014;
                    break;
                case "HANADB":
                    oComp.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB;
                    break;
                case "MSSQL2016":
                    oComp.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2016;
                    break;
            }

            //oComp.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2012;
            oComp.CompanyDB = textBox2.Text;
            oComp.UserName = textBox3.Text;
            oComp.Password = textBox4.Text;
            oComp.LicenseServer = textBox5.Text;
            try
            {
                int retCode = 0;
                retCode = oComp.Connect();
                if (retCode == 0)
                    ShowMessage("Welcome to " + oComp.CompanyName);
                else
                    ShowMessage(string.Format("Please set connection [{0}]", oComp.GetLastErrorDescription()));
            }
            catch (Exception ex)
            {
                ShowMessage(ex.Message);
            }
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void Form3_Resize(object sender, EventArgs e)
        {
            //splitContainer1.SplitterDistance = 135;
        }


        private void button4_Click(object sender, EventArgs e)
        {
            if (button4.Text == "Show Config")
            {
                button4.Text = "Hide Config";

                groupBox1.Show();
                splitContainer1.SplitterDistance = 269;
            }
            else
            {
                button4.Text = "Show Config";
                groupBox1.Hide();
                splitContainer1.SplitterDistance = 0;
            }


        }

        private void button5_Click(object sender, EventArgs e)
        {
            this.groupBox1.Hide();
            splitContainer1.SplitterDistance = 0;
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }
    }
}

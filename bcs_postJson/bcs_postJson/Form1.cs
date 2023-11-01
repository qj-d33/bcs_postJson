// bcs_postJson.Form1
using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Net;
using System.Text;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using System.Xml;
using Aras.IOM;
using bcs_postJson;
using Newtonsoft.Json.Linq;



public class WriteLog
{
    public static void SaveRecord(string content)
    {
        if (string.IsNullOrEmpty(content))
        {
            return;
        }
        content = string.Format(DateTime.Now.ToString() + "：\r\n{0}\r\n\r\n", content);
        FileStream fileStream = null;
        StreamWriter streamWriter = null;
        try
        {
            string path = Path.Combine(AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "log\\", $"{DateTime.Now:yyyy-MM-dd}" + ".log");
            using (fileStream = new FileStream(path, FileMode.Append, FileAccess.Write))
            {
                using (streamWriter = new StreamWriter(fileStream))
                {
                    streamWriter.Write(content);
                    streamWriter?.Close();
                }
                fileStream?.Close();
            }
        }
        catch (Exception ex)
        {
            SaveRecord(ex.Message);
        }
    }
}

public class Form1 : Form
{
    public Innovator inn;

    public string write_log = "";

    public string DataInteractionUrl;

    private JavaScriptSerializer serializer = new JavaScriptSerializer();

    private IContainer components = null;

    public Form1()
    {
        InitializeComponent();
    }

    private void Form1_Load(object sender, EventArgs e)
    {
        try
        {
            if (!LoginInnovator())
            {
                write_log += "登入失敗\r\n";
                Exit();
                return;
            }
            write_log += "登入成功\r\n";
            GetDataInteractionUrl();
            write_log += "物料导入 拋轉開始!!\r\n";
            Jsonpost_item("ITEM_IMPORT_IFACE_1");
            write_log += "物料导入 拋轉完成!!\r\n";
            write_log += "客戶料號對照 拋轉開始!!\r\n";
            Jsonpost_item("CI_IMPORT_INTERFACE");
            write_log += "客戶料號對照 拋轉完成!!\r\n";
            write_log += "BOM物料导入 拋轉開始!!\r\n";
            Jsonpost_item("BOM_IMPORT_IFACE");
            write_log += "BOM物料导入 拋轉完成!!\r\n";
            write_log += "ROUTING工艺路线导入 拋轉開始!!\r\n";
            Jsonpost_item("ROUTING_IMPORT_IFACE");
            write_log += "ROUTING工艺路线导入 拋轉完成!!\r\n";
            write_log += "物料变更 拋轉開始!!\r\n";
            Jsonpost_item("ITEM_UPDATE_IFACE_1");
            write_log += "物料变更 拋轉完成!!\r\n";
            write_log += "操作完成，程式關閉!!\r\n";
            Exit();
        }
        catch (Exception ex)
        {
            write_log = write_log + "例外" + ex?.ToString() + "\r\n";
            Exit();
        }
    }

    public void Jsonpost_item(string item_number)
    {
        Item item = inn.newItem("Data interaction setting", "get");
        item.setProperty("item_number", item_number);
        item = item.apply();
        string requestUriString = DataInteractionUrl + item.getProperty("url");
        string sql = "select id,created_by_id,created_on,form_item_id,setting,body,post_result from innovator.Data_interaction_Log where is_post != '1' and setting = '" + item.getProperty("id") + "' order by form_item_id,setting,created_on";
        Item item2 = inn.applySQL(sql);
        for (int i = 0; i < item2.getItemCount(); i++)
        {
            Item itemByIndex = item2.getItemByIndex(i);
            string property = itemByIndex.getProperty("body", "");
            string property2 = itemByIndex.getProperty("id", "");
            property = "[" + property + "]";
            byte[] bytes = Encoding.UTF8.GetBytes(property);
            HttpWebRequest httpWebRequest = (HttpWebRequest)WebRequest.Create(requestUriString);
            httpWebRequest.Method = "POST";
            httpWebRequest.Accept = "text/html,application/xhtml+xml,*/*";
            httpWebRequest.ContentLength = bytes.Length;
            httpWebRequest.ContentType = "application/json";
            Stream requestStream = httpWebRequest.GetRequestStream();
            requestStream.Write(bytes, 0, bytes.Length);
            httpWebRequest.Timeout = 90000;
            httpWebRequest.Headers.Set("Pragma", "no-cache");
            HttpWebResponse httpWebResponse = (HttpWebResponse)httpWebRequest.GetResponse();
            Stream responseStream = httpWebResponse.GetResponseStream();
            Encoding uTF = Encoding.UTF8;
            StreamReader streamReader = new StreamReader(responseStream, uTF);
            string json = streamReader.ReadToEnd();
            JObject jObject = JObject.Parse(json);
            responseStream.Dispose();
            streamReader.Dispose();
            string text = ((JObject)jObject["RESPONSE_HEADER"])["RESPONSE_CODE"].ToString();
            string text2 = ((JObject)jObject["RESPONSE_HEADER"])["RESPONSE_MSG_CN"].ToString();
            string text3 = ((JObject)jObject["RESPONSE_HEADER"])["RESPONSE_MSG_EN"].ToString();
            string text4 = jObject["RESPONSE_BODY"].ToString();
            string text5 = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            if (text == "1000")
            {
                string property3 = itemByIndex.getProperty("post_result", "");
                string arg = "執行日期:" + text5 + "\nRESPONSE_CODE:" + text + "\nRESPONSE_MSG_CN:" + text2 + "\nRESPONSE_MSG_EN:" + text3 + "\nRESPONSE_BODY:" + text4 + "\n" + property3;
                string sql2 = string.Format("update innovator.Data_interaction_Log set is_post='1',is_finish='1',post_result=N'{1}' where id='{0}'", property2, arg);
                inn.applySQL(sql2);
                continue;
            }
            string property4 = itemByIndex.getProperty("post_result", "");
            string text6 = "執行日期:" + text5 + "\nRESPONSE_CODE:" + text + "\nRESPONSE_MSG_CN:" + text2 + "\nRESPONSE_MSG_EN:" + text3 + "\nRESPONSE_BODY:" + text4 + "\n" + property4;
            Item itemById = inn.getItemById("Data interaction Log", property2);
            string property5 = itemById.getProperty("body");
            string[] array = property5.Split(',');
            string text7 = "";
            string text8 = "";
            string text9 = property5;
            bool flag = false;
            for (int j = 0; j < array.Length; j++)
            {
                int num = array[j].ToString().IndexOf("\"REQUEST_ID\"");
                if (num >= 0)
                {
                    text8 = array[j];
                    text7 = array[j];
                    if (text8.Substring(text8.Length - 1, 1) == "}")
                    {
                        flag = true;
                        text8 = text8.Substring(0, text8.Length - 3) + "\"}";
                    }
                    else
                    {
                        text8 = text8.Substring(0, text8.Length - 2) + "\"";
                    }
                    text9 = text9.Replace(text7, text8);
                }
            }
            int num2 = 45;
            if (flag)
            {
                num2 = 46;
            }
            if (text8.Length < num2)
            {
                inn.applySQL("update innovator.Data_interaction_Log set is_post='1' ,post_result=N'" + text6 + "' where id='" + itemById.getProperty("id") + "'");
            }
            else
            {
                inn.applySQL("update innovator.Data_interaction_Log set body =N'" + text9 + "',is_post='0' ,post_result=N'" + text6 + "' where id='" + itemById.getProperty("id") + "'");
            }
        }
    }

    public void Jsonpost_inv_Part()
    {
        Item item = inn.newItem("Data interaction setting", "get");
        item.setProperty("item_number", "ITEM_IMPORT_IFACE_1");
        item = item.apply();
        string requestUriString = DataInteractionUrl + item.getProperty("url");
        string sql = "select id,created_by_id,created_on,form_item_id,setting,body from innovator.Data_interaction_Log where is_post != '1' and setting = '988FC2E748684F649BF1725198565AD9' order by form_item_id,setting,created_on";
        Item item2 = inn.applySQL(sql);
        for (int i = 0; i < item2.getItemCount(); i++)
        {
            Item itemByIndex = item2.getItemByIndex(i);
            string property = itemByIndex.getProperty("body", "");
            string property2 = itemByIndex.getProperty("id", "");
            property = "[" + property + "]";
            byte[] bytes = Encoding.UTF8.GetBytes(property);
            HttpWebRequest httpWebRequest = (HttpWebRequest)WebRequest.Create(requestUriString);
            httpWebRequest.Method = "POST";
            httpWebRequest.Accept = "text/html,application/xhtml+xml,*/*";
            httpWebRequest.ContentLength = bytes.Length;
            httpWebRequest.ContentType = "application/json";
            Stream requestStream = httpWebRequest.GetRequestStream();
            requestStream.Write(bytes, 0, bytes.Length);
            httpWebRequest.Timeout = 90000;
            httpWebRequest.Headers.Set("Pragma", "no-cache");
            HttpWebResponse httpWebResponse = (HttpWebResponse)httpWebRequest.GetResponse();
            Stream responseStream = httpWebResponse.GetResponseStream();
            Encoding uTF = Encoding.UTF8;
            StreamReader streamReader = new StreamReader(responseStream, uTF);
            string json = streamReader.ReadToEnd();
            dynamic val = JObject.Parse(json);
            string text = val.RESPONSE_HEADER;
            string arg = val.RESPONSE_BODY;
            if (text == "1000")
            {
                string sql2 = string.Format("update innovator.Data_interaction_Log set is_post='1',is_finish='1',post_result=N'{1}' where id='{0}'", property2, arg);
                inn.applySQL(sql2);
            }
            else
            {
                string sql3 = string.Format("update innovator.Data_interaction_Log set post_result=N'{1}' where id='{0}'", property2, arg);
                inn.applySQL(sql3);
            }
        }
    }

    public void Jsonpost_upd_Part()
    {
        Item item = inn.newItem("Data interaction setting", "get");
        item.setProperty("item_number", "ITEM_UPDATE_IFACE_1");
        item = item.apply();
        string requestUriString = DataInteractionUrl + item.getProperty("url");
        string sql = "select id,created_by_id,created_on,form_item_id,setting,body from innovator.Data_interaction_Log where is_post != '1' and setting = '4918B773FFA64A07A554C3A6F6D2B397' order by form_item_id,setting,created_on";
        Item item2 = inn.applySQL(sql);
        for (int i = 0; i < item2.getItemCount(); i++)
        {
            Item itemByIndex = item2.getItemByIndex(i);
            string property = itemByIndex.getProperty("body", "");
            string property2 = itemByIndex.getProperty("id", "");
            property = "[" + property + "]";
            byte[] bytes = Encoding.UTF8.GetBytes(property);
            HttpWebRequest httpWebRequest = (HttpWebRequest)WebRequest.Create(requestUriString);
            httpWebRequest.Method = "POST";
            httpWebRequest.Accept = "text/html,application/xhtml+xml,*/*";
            httpWebRequest.ContentLength = bytes.Length;
            httpWebRequest.ContentType = "application/json";
            Stream requestStream = httpWebRequest.GetRequestStream();
            requestStream.Write(bytes, 0, bytes.Length);
            httpWebRequest.Timeout = 90000;
            httpWebRequest.Headers.Set("Pragma", "no-cache");
            HttpWebResponse httpWebResponse = (HttpWebResponse)httpWebRequest.GetResponse();
            Stream responseStream = httpWebResponse.GetResponseStream();
            Encoding uTF = Encoding.UTF8;
            StreamReader streamReader = new StreamReader(responseStream, uTF);
            string json = streamReader.ReadToEnd();
            dynamic val = JObject.Parse(json);
            string text = val.RESPONSE_HEADER;
            string arg = val.RESPONSE_BODY;
            if (text == "1000")
            {
                string sql2 = string.Format("update innovator.Data_interaction_Log set is_post='1',is_finish='1',post_result=N'{1}' where id='{0}'", property2, arg);
                inn.applySQL(sql2);
            }
            else
            {
                string sql3 = string.Format("update innovator.Data_interaction_Log set post_result=N'{1}' where id='{0}'", property2, arg);
                inn.applySQL(sql3);
            }
        }
    }

    public void Jsonpost_inv_cat()
    {
        Item item = inn.newItem("Data interaction setting", "get");
        item.setProperty("item_number", "ITEM_CAT_IMPORT_IFACE");
        item = item.apply();
        string requestUriString = DataInteractionUrl + item.getProperty("url");
        string sql = "select id,created_by_id,created_on,form_item_id,setting,body from innovator.Data_interaction_Log where is_post != '1' and setting = '222FB256430A4A5EA5142A5EB2C1453D' order by form_item_id,setting,created_on";
        Item item2 = inn.applySQL(sql);
        for (int i = 0; i < item2.getItemCount(); i++)
        {
            Item itemByIndex = item2.getItemByIndex(i);
            string property = itemByIndex.getProperty("body", "");
            string property2 = itemByIndex.getProperty("id", "");
            property = "[" + property + "]";
            byte[] bytes = Encoding.UTF8.GetBytes(property);
            HttpWebRequest httpWebRequest = (HttpWebRequest)WebRequest.Create(requestUriString);
            httpWebRequest.Method = "POST";
            httpWebRequest.Accept = "text/html,application/xhtml+xml,*/*";
            httpWebRequest.ContentLength = bytes.Length;
            httpWebRequest.ContentType = "application/json";
            Stream requestStream = httpWebRequest.GetRequestStream();
            requestStream.Write(bytes, 0, bytes.Length);
            httpWebRequest.Timeout = 90000;
            httpWebRequest.Headers.Set("Pragma", "no-cache");
            HttpWebResponse httpWebResponse = (HttpWebResponse)httpWebRequest.GetResponse();
            Stream responseStream = httpWebResponse.GetResponseStream();
            Encoding uTF = Encoding.UTF8;
            StreamReader streamReader = new StreamReader(responseStream, uTF);
            string json = streamReader.ReadToEnd();
            dynamic val = JObject.Parse(json);
            string text = val.RESPONSE_HEADER;
            string arg = val.RESPONSE_BODY;
            if (text == "1000")
            {
                string sql2 = string.Format("update innovator.Data_interaction_Log set is_post='1',is_finish='1',post_result=N'{1}' where id='{0}'", property2, arg);
                inn.applySQL(sql2);
            }
            else
            {
                string sql3 = string.Format("update innovator.Data_interaction_Log set post_result=N'{1}' where id='{0}'", property2, arg);
                inn.applySQL(sql3);
            }
        }
    }

    public void Jsonpost_inv_ROUTING()
    {
        Item item = inn.newItem("Data interaction setting", "get");
        item.setProperty("item_number", "ROUTING_IMPORT_IFACE");
        item = item.apply();
        string url = DataInteractionUrl + item.getProperty("url");
        string sql = "select id,created_by_id,created_on,form_item_id,setting,body from innovator.Data_interaction_Log where is_post != '1' and setting = '75FC20DC1FF141B4BC7BC910FAE13F24' order by form_item_id,setting,created_on";
        Item item2 = inn.applySQL(sql);
        for (int i = 0; i < item2.getItemCount(); i++)
        {
            Item itemByIndex = item2.getItemByIndex(i);
            string property = itemByIndex.getProperty("body", "");
            string property2 = itemByIndex.getProperty("id", "");
            property = "[" + property + "]";
            PostWS(url, property, property2);
        }
    }

    public void PostWS(string Url, string JSONData, string Log_id)
    {
        Hashtable hashtable = new Hashtable();
        try
        {
            byte[] bytes = Encoding.UTF8.GetBytes(JSONData);
            HttpWebRequest httpWebRequest = (HttpWebRequest)WebRequest.Create(Url);
            httpWebRequest.Method = "POST";
            httpWebRequest.Accept = "text/html,application/xhtml+xml,*/*";
            httpWebRequest.ContentLength = bytes.Length;
            httpWebRequest.ContentType = "application/json";
            Stream requestStream = httpWebRequest.GetRequestStream();
            requestStream.Write(bytes, 0, bytes.Length);
            httpWebRequest.Timeout = 90000;
            httpWebRequest.Headers.Set("Pragma", "no-cache");
            HttpWebResponse httpWebResponse = (HttpWebResponse)httpWebRequest.GetResponse();
            Stream responseStream = httpWebResponse.GetResponseStream();
            Encoding uTF = Encoding.UTF8;
            StreamReader streamReader = new StreamReader(responseStream, uTF);
            string json = streamReader.ReadToEnd();
            dynamic val = JObject.Parse(json);
            string text = val.RESPONSE_HEADER;
            string arg = val.RESPONSE_BODY;
            if (text == "1000")
            {
                string sql = string.Format("update innovator.Data_interaction_Log set is_post='1',is_finish='1',post_result=N'{1}' where id='{0}'", Log_id, arg);
                inn.applySQL(sql);
            }
            else
            {
                string sql2 = string.Format("update innovator.Data_interaction_Log set post_result=N'{1}' where id='{0}'", Log_id, arg);
                inn.applySQL(sql2);
            }
        }
        catch
        {
        }
    }

    public void Exit()
    {
        try
        {
            WriteLog.SaveRecord(write_log);
            Environment.Exit(0);
        }
        catch
        {
            WriteLog.SaveRecord(write_log);
            WriteLog.SaveRecord("关闭程序时出错");
            Environment.Exit(0);
        }
    }

    public bool LoginInnovator()
    {
        string text = "";
        string text2 = "";
        string text3 = "";
        string text4 = "";
        XmlDocument xmlDocument = new XmlDocument();
        xmlDocument.Load(AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "login.config.xml");
        XmlNodeList xmlNodeList = xmlDocument.SelectNodes("//section[@name='default']");
        foreach (XmlNode item2 in xmlNodeList)
        {
            text = item2.SelectNodes("//item[@name='url']")[0].Attributes["value"].InnerText;
            text2 = item2.SelectNodes("//item[@name='DB']")[0].Attributes["value"].InnerText;
            text3 = item2.SelectNodes("//item[@name='loginName']")[0].Attributes["value"].InnerText;
            text4 = item2.SelectNodes("//item[@name='pwd']")[0].Attributes["value"].InnerText;
        }
        if (string.IsNullOrEmpty(text) || string.IsNullOrEmpty(text2) || string.IsNullOrEmpty(text3) || string.IsNullOrEmpty(text4))
        {
            write_log += "登入資訊填寫不完整!!\r\n";
        }
        HttpServerConnection httpServerConnection = IomFactory.CreateHttpServerConnection(text, text2, text3, text4);
        Item item = httpServerConnection.Login();
        if (item.isError())
        {
            write_log = write_log + "登录失败,失败原因:" + item.getErrorString() + "\r\n";
            return false;
        }
        inn = IomFactory.CreateInnovator(httpServerConnection);
        httpServerConnection.Logout();
        return true;
    }

    public void GetDataInteractionUrl()
    {
        Item itemByKeyedName = inn.getItemByKeyedName("Variable", "TS_DataInteractionUrl2");
        DataInteractionUrl = itemByKeyedName.getProperty("value");
        if (DataInteractionUrl.Substring(DataInteractionUrl.Length - 1, 1) != "/")
        {
            DataInteractionUrl += "/";
        }
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing && components != null)
        {
            components.Dispose();
        }
        base.Dispose(disposing);
    }

    private void InitializeComponent()
    {
        System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(bcs_postJson.Form1));
        SuspendLayout();
        base.AutoScaleDimensions = new System.Drawing.SizeF(8f, 15f);
        base.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
        base.ClientSize = new System.Drawing.Size(226, 169);
        base.Icon = (System.Drawing.Icon)resources.GetObject("$this.Icon");
        base.Name = "Form1";
        Text = "Form1";
        base.Load += new System.EventHandler(Form1_Load);
        ResumeLayout(false);
    }
}

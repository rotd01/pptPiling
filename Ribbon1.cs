using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;
using System.Windows.Forms;
using System.Net;
using System.IO;
using System.Web.Script.Serialization;
using System.Net.NetworkInformation;


namespace PiliangPPT
{
    public partial class Ribbon1
    {
        private string host = "http://47.117.112.225:8080/v1/puser/getMsg";
        public static string GetHttpResponse(string url)
        {
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
            request.Method = "GET";
            request.ContentType = "text/html;charset=UTF-8";
            request.UserAgent = null;
            request.Timeout = 5000;

            HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            Stream myResponseStream = response.GetResponseStream();
            StreamReader myStreamReader = new StreamReader(myResponseStream, Encoding.GetEncoding("utf-8"));
            string retString = myStreamReader.ReadToEnd();
            myStreamReader.Close();
            myResponseStream.Close();

            return retString;

        }
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

         private static string GetFirstMacAddress()
        {
            try
            {
                string macAddress = "" ;
                NetworkInterface[] nics = NetworkInterface.GetAllNetworkInterfaces(); 
                foreach (NetworkInterface adapter in nics) 
                { 
                    if (!adapter.GetPhysicalAddress().ToString().Equals("")) 
                    { 
                        macAddress = adapter.GetPhysicalAddress().ToString(); 
                        for (int i = 1; i < 6; i++) 
                        { 
                            macAddress = macAddress.Insert(3 * i - 1, ":"); 
                        } 
                        break; 
                    } 
                }
                return macAddress;
            }
            catch (Exception)
            {
                return "";
            }

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Login login = null;
            if (login == null || login.IsDisposed)
            {
                button1.Enabled = false;
                login = new Login();
                IntPtr handle = Process.GetCurrentProcess().MainWindowHandle;
                NativeWindow win = NativeWindow.FromHandle(handle);
                login.Show();

            }
            //string mac = GetFirstMacAddress();
            //if(!"00:F1:F3:A6:2D:88".Equals(mac)){
            //    MessageBox.Show("当前机器不匹配"+ mac);
            //    return;
            //}
            //Picture_Jigsaw Picture_Jigsaw = null;
            //if (Picture_Jigsaw == null || Picture_Jigsaw.IsDisposed)
            //{
            //    Globals.Ribbons.Ribbon1.button1.Enabled = false;
            //    Picture_Jigsaw = new Picture_Jigsaw();
            //    IntPtr handle = Process.GetCurrentProcess().MainWindowHandle;
            //    NativeWindow win = NativeWindow.FromHandle(handle);
            //    Picture_Jigsaw.Show();
            //}


            //try
            //{
            //    String res = GetHttpResponse(host);
            //    if (res == null || "".Equals(res))
            //    {
            //        label3.Label = "暂时没有公告";
            //    }
            //    else
            //    {
            //        JavaScriptSerializer serializer = new JavaScriptSerializer();
            //        Dictionary<string, object> json = (Dictionary<string, object>)serializer.DeserializeObject(res);

            //        label3.Label = json["msg"].ToString();
            //    }
            //}
            //catch (Exception err)
            //{
            //    label3.Label = "暂时没有公告";
            //}

        }

        private void editBox1_TextChanged(object sender, RibbonControlEventArgs e)
        {

        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start("IEXPLORE.EXE", "https://www.o-street.cn");
            }
            catch(Exception err)
            {
                MessageBox.Show(err.ToString());
            }
            
        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show("售后微信：wangkun7991");
            MsgForm form = new MsgForm();
            form.Show();
            //Presentation ppt = new Presentation(@"C:\Users\kuniq\Desktop\pptTest\02.清新风格PPT.pptx");
            //ppt.Save(@"C:\Users\kuniq\Desktop\pptTest\02.清新风格PPT.pdf", Aspose.Slides.Export.SaveFormat.Pdf);

        }
    }
}

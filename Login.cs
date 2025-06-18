
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;

using System.Windows.Forms;
using System.Diagnostics;
using System.Web.Script.Serialization;
using System.Net.NetworkInformation;

namespace PiliangPPT
{
    public partial class Login : Form
    {
        //private string host = "http://127.0.0.1:8090/v1/puser/pptuser?email=";
        private string host = "http://47.117.112.225:8080/v1/puser/pptuser?v=11&email=";

        public static string GetHttpResponse(string url)
        {
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
            request.Method = "GET";
            request.ContentType = "text/html;charset=UTF-8";
            request.UserAgent = null;
            request.Timeout = 10000;

            HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            Stream myResponseStream = response.GetResponseStream();
            StreamReader myStreamReader = new StreamReader(myResponseStream, Encoding.GetEncoding("utf-8"));
            string retString = myStreamReader.ReadToEnd();
            myStreamReader.Close();
            myResponseStream.Close();

            return retString;

        }
        public Login()
        {
            InitializeComponent();
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
                //string macAddresses = string.Empty;

                //foreach (NetworkInterface nic in NetworkInterface.GetAllNetworkInterfaces())
                //{
                //    if (nic.OperationalStatus == OperationalStatus.Up)
                //    {
                //        macAddresses += nic.GetPhysicalAddress().ToString();
                //        //Console.WriteLine(macAddresses);
                //        break;
                //    }
                //}
                //return macAddresses.Replace(":", "-");
            }
            catch (Exception)
            {
                return "";
            }

        }
        private static string ExecuteCMD(string cmd, Func<string, string> filterFunc)
        {
            var process = new Process();
            process.StartInfo.FileName = "cmd.exe";
            process.StartInfo.UseShellExecute = false;//是否使用操作系统shell启动
            process.StartInfo.RedirectStandardInput = true;//接受来自调用程序的输入信息
            process.StartInfo.RedirectStandardOutput = true;//由调用程序获取输出信息
            process.StartInfo.RedirectStandardError = true;//重定向标准错误输出
            process.StartInfo.CreateNoWindow = true;//不显示程序窗口
            process.Start();//启动程序
            process.StandardInput.WriteLine(cmd + " &exit");
            process.StandardInput.AutoFlush = true;
            //获取cmd窗口的输出信息
            var output = process.StandardOutput.ReadToEnd();
            process.WaitForExit();
            process.Close();
            return filterFunc(output);
        }
        public static string GetCPUID()
        {
            var cmd = "wmic cpu get processorid";
            return ExecuteCMD(cmd, output =>
            {
                var cpuid = GetTextAfterSpecialText(output, "ProcessorId");
                return cpuid;
            });
        }
        private static string GetTextAfterSpecialText(string fullText, string specialText)
        {
            if (string.IsNullOrWhiteSpace(fullText) || string.IsNullOrWhiteSpace(specialText))
            {
                return null;
            }
            string lastText = null;
            var idx = fullText.LastIndexOf(specialText);
            if (idx > 0)
            {
                lastText = fullText.Substring(idx + specialText.Length).Trim();
            }
            return lastText;
        }

        public static string GetBIOSSerialNumber()
        {
            var cmd = "wmic bios get serialnumber";
            return ExecuteCMD(cmd, output =>
            {
                var serialNumber = GetTextAfterSpecialText(output, "SerialNumber");
                return serialNumber;
            });
        }


        private void checkLogin()
        {
            Globals.Ribbons.Ribbon1.button1.Enabled = false;
            //MessageBox.Show(Properties.Settings.Default.user);
            if (Properties.Settings.Default.user != "")
            {
                String mac_ = GetFirstMacAddress();
                Properties.Settings.Default.mac = mac_;
                String user = Properties.Settings.Default.user;
                label3.Text = "正在验证：" + Properties.Settings.Default.user;
                string res = "";
                try
                {
                   
                    res = GetHttpResponse(host + user + "&mac=" + mac_);
                }
                catch (Exception e)
                {

                    MessageBox.Show("连接超时了，可能是服务器错误，" + e.ToString());
                    Globals.Ribbons.Ribbon1.button1.Enabled = true;
                    label3.Visible = false;
                    panel1.Visible = true;
                   
                }
               
      
                if (res != "")
                {
                    JavaScriptSerializer serializer = new JavaScriptSerializer();
                    Dictionary<string, object> json = (Dictionary<string, object>)serializer.DeserializeObject(res);
                    if (json["err"].ToString().Equals("0"))
                    {
                        //Globals.Ribbons.Ribbon1.button170.Enabled = true;
                        Picture_Jigsaw Picture_Jigsaw = null;
                        if (Picture_Jigsaw == null || Picture_Jigsaw.IsDisposed)
                        {
                            this.Close();
                            Globals.Ribbons.Ribbon1.button1.Enabled = false;
                            Picture_Jigsaw = new Picture_Jigsaw();
                            IntPtr handle = Process.GetCurrentProcess().MainWindowHandle;
                            NativeWindow win = NativeWindow.FromHandle(handle);
                            Picture_Jigsaw.Show();
                           
                        }
                    }
                    else if (json["msg"].ToString() != "")
                    {
                        MessageBox.Show(json["msg"].ToString());
                        label3.Visible = false;
                        panel1.Visible = true;
                    }
                    else
                    {
                        MessageBox.Show("验证失败");
                        label3.Visible = false;
                        panel1.Visible = true;
                    }
                }
            }
            else
            {
                label3.Visible = false;
                panel1.Visible = true;
            }
        }
        private void offLineCheck() {
            //string str1 = GetBIOSSerialNumber();
            //string str2 = GetCPUID();
            //string date =  > ;

            if(new DateTime(2025,6,18).AddDays(30) > DateTime.Now){
                //if(str2.Equals("BFEBFBFF000906E9")){
                if(true) { 
                    Picture_Jigsaw Picture_Jigsaw = null;
                    if (Picture_Jigsaw == null || Picture_Jigsaw.IsDisposed)
                    {
                        this.Close();
                        Globals.Ribbons.Ribbon1.button1.Enabled = false;
                        Picture_Jigsaw = new Picture_Jigsaw();
                        IntPtr handle = Process.GetCurrentProcess().MainWindowHandle;
                        NativeWindow win = NativeWindow.FromHandle(handle);
                        Picture_Jigsaw.Show();
                           
                    }


                }else{
                    MessageBox.Show("机器验证失败");
                }
            }else{
                MessageBox.Show("30天使用已过期，请联系微信:wangkun7991");
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            Login.ActiveForm.Close();
            //Globals.Ribbons.Ribbon1.button170.Enabled = true;

            Picture_Jigsaw Picture_Jigsaw = null;
            if (Picture_Jigsaw == null || Picture_Jigsaw.IsDisposed)
            {
                Picture_Jigsaw = new Picture_Jigsaw();
                IntPtr handle = Process.GetCurrentProcess().MainWindowHandle;
                NativeWindow win = NativeWindow.FromHandle(handle);

                Picture_Jigsaw.Show();
            }


            //if (textBox1.Text.Length == 0)
            //{
            //    MessageBox.Show("请输入你的邮箱");
            //    return;
            //}
            //String mac_ = GetFirstMacAddress();

            //string res = "";
            //try
            //{
            //    res = GetHttpResponse(host + textBox1.Text + "&mac=" + mac_);
            //}
            //catch (Exception err)
            //{
            //    MessageBox.Show("连接超时了，可能是服务器错误," + err.ToString());
            //}
            //if (res != "")
            //{
            //    JavaScriptSerializer serializer = new JavaScriptSerializer();
            //    Dictionary<string, object> json = (Dictionary<string, object>)serializer.DeserializeObject(res);
            //    if (json["err"].ToString().Equals("0"))
            //    {
            //        Properties.Settings.Default.user = textBox1.Text;
            //        Properties.Settings.Default.Save();
            //        MessageBox.Show("验证成功");
            //        Login.ActiveForm.Close();
            //        //Globals.Ribbons.Ribbon1.button170.Enabled = true;

            //        Picture_Jigsaw Picture_Jigsaw = null;
            //        if (Picture_Jigsaw == null || Picture_Jigsaw.IsDisposed)
            //        {
            //            Picture_Jigsaw = new Picture_Jigsaw();
            //            IntPtr handle = Process.GetCurrentProcess().MainWindowHandle;
            //            NativeWindow win = NativeWindow.FromHandle(handle);

            //            Picture_Jigsaw.Show();
            //        }

            //    }
            //    else if (json["msg"].ToString() != "")
            //    {
            //        MessageBox.Show(json["msg"].ToString());
            //    }
            //    else
            //    {
            //        MessageBox.Show("验证失败");
            //    }

            //}

        }

        private void button2_Click(object sender, EventArgs e)
        {
            Login.ActiveForm.Close();
            Globals.Ribbons.Ribbon1.button1.Enabled = true;
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void Login_Load(object sender, EventArgs e)
        {
            //checkLogin();
            offLineCheck();
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            base.OnFormClosed(e);
            Globals.Ribbons.Ribbon1.button1.Enabled = true;
            //MessageBox.Show("ddd");
        }

        private void label3_Click(object sender, EventArgs e)
        {
            
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void label3_Click_1(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }
    }
}
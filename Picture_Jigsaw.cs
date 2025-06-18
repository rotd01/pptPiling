using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.Collections;
using System.Drawing.Drawing2D;
using System.Diagnostics;
using System.Drawing.Text;
using Microsoft.Office.Core;
namespace PiliangPPT
{
    public partial class Picture_Jigsaw : Form
    {

        public Picture_Jigsaw()
        {
            InitializeComponent();
            this.Deactivate += new EventHandler(color1_Deactivate);

            //Properties.Settings.Default.bottom = textBox12.Text.Trim();
            //Properties.Settings.Default.Save();

            textBox12.Text = Properties.Settings.Default.bottom;
            textBox13.Text = Properties.Settings.Default.leftRight;
            textBox14.Text = Properties.Settings.Default.right;
            textBox9.Text = Properties.Settings.Default.opacity;
            textBox10.Text = Properties.Settings.Default.splitCount;
            textBox7.Text = Properties.Settings.Default.path;
            textBox1.Text = Properties.Settings.Default.width.ToString();
            textBox2.Text = Properties.Settings.Default.col.ToString();
            textBox3.Text = Properties.Settings.Default.inner_space.ToString();
            textBox4.Text = Properties.Settings.Default.out_space.ToString();
            textBox5.Text = Properties.Settings.Default.big_num;
            textBox6.Text = Properties.Settings.Default.logo;
            trackBar1.Value = Properties.Settings.Default.water_r;
            trackBar2.Value = Properties.Settings.Default.water_c;
            trackBar3.Value = Properties.Settings.Default.water_a;

            label23.Font = Properties.Settings.Default.water_f;

            checkBox2.Checked = Properties.Settings.Default.autoBg;
            textBox11.Text = Properties.Settings.Default.autoBgOpacity;

            if(checkBox2.Checked == true){
                label33.Visible = true;
                textBox11.Visible = true;

                label9.Visible = false;
                panel1.Visible = false;

            }else{
                label33.Visible = false;
                textBox11.Visible = false;

                label9.Visible = true;
                panel1.Visible = true;
            }

            if (Properties.Settings.Default.bg_color != 0)
            {
                panel1.BackColor = Color.FromArgb(Properties.Settings.Default.bg_color);
            }
            if (Properties.Settings.Default.logo_color != 0)
            {
                panel2.BackColor = Color.FromArgb(Properties.Settings.Default.logo_color);
            }


        }

        void color1_Deactivate(object sender, EventArgs e)
        {
            //if(timer1.Enabled == true || timer2.Enabled == true){
            //    this.Focus();
            //}
            timer1.Enabled = false;
            timer2.Enabled = false;
            //this.transparentPanel1.Visible = false;
            
            //this.panel3.Visible = false;
        }

        private PowerPoint.Application app = Globals.ThisAddIn.Application;
        private ArrayList slider_array = new ArrayList();
        private List<String> name_list = new List<String>();
        private bool m_isMouseDown = false;
        private Point m_mousePos = new Point();
        private System.Threading.Thread _thread;

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            base.OnFormClosed(e);
            Globals.Ribbons.Ribbon1.button1.Enabled = true;
            //MessageBox.Show("ddd");
        }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e);
            m_mousePos = Cursor.Position;
            m_isMouseDown = true;
            if (e.Button == MouseButtons.Right){
                timer1.Enabled = false;
                timer2.Enabled = false;
            }
        }

        protected override void OnMouseUp(MouseEventArgs e)
        {
            base.OnMouseUp(e);
            
            m_isMouseDown = false;
            if (e.Button == MouseButtons.Right){
                timer1.Enabled = false;
                timer2.Enabled = false;
            }
        }

        //protected override void OnMouseMove(MouseEventArgs e)
        //{
        //    base.OnMouseMove(e);
        //    if (m_isMouseDown)
        //    {
        //        Point tempPos = Cursor.Position;
        //        this.Location = new Point(Location.X + (tempPos.X - m_mousePos.X), Location.Y + (tempPos.Y - m_mousePos.Y));
        //        m_mousePos = Cursor.Position;
        //    }
        //}

        float w; float h;

        List<string> strArr;

        private void pintu1_Load(object sender, EventArgs e)
        {
            w = app.ActivePresentation.PageSetup.SlideWidth;
            h = app.ActivePresentation.PageSetup.SlideHeight;
            label10.Text = w + "*" + h;
            label10.ForeColor = Color.Orange;
            //textBox1.Text = w + "";
        }
        private void textBox7_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.FolderBrowserDialog dialoaFolder = new FolderBrowserDialog();
            if (!textBox7.Text.Trim().Equals("") && Directory.Exists(textBox7.Text.Trim()))
            {
                dialoaFolder.SelectedPath = textBox7.Text.Trim();
            }

            if (dialoaFolder.ShowDialog() == DialogResult.OK)
            {
                Properties.Settings.Default.path = dialoaFolder.SelectedPath;
                Properties.Settings.Default.Save();
                textBox7.Text = dialoaFolder.SelectedPath;
            }
        }

        private void getFile(string path){
            
            string[] files = Directory.GetFiles(path);
            foreach (string url in files)
            {
                if(url.IndexOf("~$") > -1){
                    continue;
                }
                if(url.EndsWith(".pptx") || url.EndsWith(".ppt")){
                    listBox1.Items.Add(url);
                }
            }
            string[] d = Directory.GetDirectories(path);
            if(d.Length > 0){
                foreach (string url in d)
                {
                    getFile(url);
                }
            }
        }
        private void button5_Click(object sender, EventArgs e)
        {
            if(this.checkBox6.Checked){
                FolderBrowserDialog d = new FolderBrowserDialog();
                
                if(d.ShowDialog() == DialogResult.OK){
                    
                    //foreach (string url in d. .FileNames)
                    //{
                    //    listBox1.Items.Add(url);
                    //}
                    label16.Text = "";
                    label15.Text = "已选" + d.SelectedPath + "文件夹";
                    getFile(d.SelectedPath);
                }
                
                return;
            }

            OpenFileDialog dialog = new OpenFileDialog();

            dialog.Multiselect = true;//该值确定是否可以选择多个文件
            dialog.Title = @"请选择PPT文件";
            dialog.Filter = "Powerpoint|*.ppt;*.pptx";

            
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                //string file = dialog.FileName;

                Console.Write(dialog.FileNames);


                label16.Text = "正在选择文件";
                foreach (string url in dialog.FileNames)
                {
                    listBox1.Items.Add(url);
                }
                label16.Text = "";
                label15.Text = "已选" + listBox1.Items.Count + "个文件";

                //PowerPoint.Presentation wpresentation = app.NewPresentation.Add();

                ////FolderBrowserDialog dialog1 =
                //app.PPFileDialog(PowerPoint.PpFileDialogType.ppFileDialogOpen);
                //dialog1.ShowDialog();
                //MessageBox.Show(this, @"文件名"+ res, "提示");
            }
        }
        private delegate void SetText(String i, int idx);
        private SetText setText;

        private delegate void SetList(String i);
        private SetList setList;


        private void button6_Click(object sender, EventArgs e)
        {
            if(textBox1.Text.Trim().Equals(string.Empty)){
                MessageBox.Show("水平宽度不能为空");
                return;
            }
            if(textBox2.Text.Trim().Equals(string.Empty)){
                MessageBox.Show("图片列数不能为空");
                return;
            }
            if(textBox3.Text.Trim().Equals(string.Empty)){
                MessageBox.Show("内间距不能为空");
                return;
            }
            if(textBox4.Text.Trim().Equals(string.Empty)){
                MessageBox.Show("顶部间距不能为空");
                return;
            }

            int wresolution = int.Parse(textBox1.Text.Trim());
            int prow = int.Parse(textBox2.Text.Trim());
            int pspac = int.Parse(textBox3.Text.Trim());
            int pspac2 = int.Parse(textBox4.Text.Trim());
            int opacity = 0;
            if(!textBox9.Text.Trim().Equals(string.Empty)){
                opacity = int.Parse(textBox9.Text.Trim());
            }


            if (opacity <= 0 || opacity > 100)
            {
                MessageBox.Show("透明度输入不正确, 请输入 1 到 100的值，100为不透明");
                return;
            }
            else if (wresolution <= 0 || prow <= 0 || pspac < 0 || pspac2 < 0)
            {
                MessageBox.Show("水平宽度/列数必须是正整数，间隔大小必须是非负整数");
                return;
            }


            if (textBox7.TextLength == 0)
            {
                MessageBox.Show("请设置保存路径");
                return;
            }
            if (!Directory.Exists(textBox7.Text))
            {
                MessageBox.Show("保存路径不正确，请重新设置");
                return;
            }

            if (listBox1.Items.Count == 0)
            {
                MessageBox.Show("请选择文件");
                return;
            }

            setText = new SetText(showText);
            setList = new SetList(handelList);

            System.Threading.ThreadStart t_start = new System.Threading.ThreadStart(resolveSlider);
            _thread = new System.Threading.Thread(t_start);
            _thread.Start();
            button2.Visible = true;

            label16.Text = "正在处理";

        }

        private void handelList(String str)
        {
            listBox2.Items.Add(str);
        }

        private void showText(String str, int idx)
        {

            label16.Text = str;
            if (str == "处理完成")
            {
                if (listBox2.Items.Count > 0)
                {
                    string res = "";
                    for (int i = 0; i < listBox2.Items.Count; i++)
                    {
                        res += (listBox2.Items[i].ToString() + "\n");
                    }
                    //MessageBox.Show("出错的文件: \n" + res);

                    MsgForm msgForm = null;

                    if (msgForm == null || msgForm.IsDisposed)
                    {

                        msgForm = new MsgForm();
                        IntPtr handle = Process.GetCurrentProcess().MainWindowHandle;
                        NativeWindow win = NativeWindow.FromHandle(handle);
                        msgForm.Show();
                        msgForm.setText(res);
                        listBox2.Items.Clear();
                    }

                }
                button2.Visible = false;
            }

            //progressBar1.Value = _value;
        }
        private void resolveSlider()
        {
            for (int i = 0; i < listBox1.Items.Count; i++)
            {

                this.Invoke(setText, "正在处理第" + (i + 1) + "个文件", i);
                try
                {
                    //this.Invoke(setList, listBox1.Items[i].ToString());
                    SaveItem(i);
                }
                catch (System.Threading.ThreadAbortException e)
                {

                    MessageBox.Show("取消成功");
                    
                    this.Invoke(setText, "处理完成", i);
                    break;
                }
                catch (Exception e)
                {
                    // MessageBox.Show("有可能没有存储权限，请以管理员方式运行WPS或者Office, 此外，要关闭被占用的PPT \n" + " " + e.ToString());
                    //strArr.Add(listBox1.Items[i].ToString());
                    this.Invoke(setList, "【" + listBox1.Items[i].ToString() + "】\n错误信息：" + e.ToString() + "\n\n");
                    //this.Invoke(setText, "第" + (i + 1) + "个文件出错", i);
                    continue;
                }

                if (i + 1 == listBox1.Items.Count)
                {
                    this.Invoke(setText, "处理完成", i);
                }
            }
        }
        /// <summary>
        /// 获取一个带有透明度的ImageAttributes
        /// </summary>
        /// <param name="opcity"></param>
        /// <returns></returns>
        public ImageAttributes GetAlphaImgAttr(int opcity)
        {
             if (opcity < 0 || opcity > 100)
             {
                  throw new ArgumentOutOfRangeException("opcity 值为 0~100");
             }
             //颜色矩阵
             float[][] matrixItems =
             {
                  new float[]{
                   1,0,0,0,0},
                          new float[]{
                   0,1,0,0,0},
                          new float[]{
                   0,0,1,0,0},
                          new float[]{
                   0,0,0,(float)opcity / 100,0},
                          new float[]{
                   0,0,0,0,1}
             };
             ColorMatrix colorMatrix = new ColorMatrix(matrixItems);
             ImageAttributes imageAtt = new ImageAttributes();
             imageAtt.SetColorMatrix(colorMatrix, ColorMatrixFlag.Default, ColorAdjustType.Bitmap);
             return imageAtt;
        }

        private void SaveItem(Object _index)
        {
            String ppt_path = listBox1.Items[(int)_index].ToString();
            PowerPoint.Presentation _ppt = null;
            try
            {

                _ppt = app.Presentations.Open(
                   ppt_path, Office.MsoTriState.msoTrue, Office.MsoTriState.msoFalse, Office.MsoTriState.msoFalse
                );
                //if((int)_index == 4)
                //{
                //    throw new Exception("wosdd");
                //}

                //            PowerPoint.Slides _slides = _ppt.Slides;
                //            slider_array.Add(_slides);
                //            name_list.Add(_ppt.Name);

            
                
                float h = _ppt.PageSetup.SlideHeight;
                float w = _ppt.PageSetup.SlideWidth;
                //PowerPoint.Slides item = (PowerPoint.Slides)slider_array[(int)_index];
                PowerPoint.Slides item = _ppt.Slides;
                //string save_name_1 = textBox7.Text + @"\shipin" + ".mp4";


                int wresolution = int.Parse(textBox1.Text.Trim());
                int prow = int.Parse(textBox2.Text.Trim());
                int pspac = int.Parse(textBox3.Text.Trim());
                int pspac2 = int.Parse(textBox4.Text.Trim());
                int opacity = int.Parse(textBox9.Text.Trim());


                // 添加上下左右边距
                int leftRight = int.Parse(textBox13.Text.Trim());
                int right = int.Parse(textBox14.Text.Trim());
                int top = int.Parse(textBox4.Text.Trim());
                int bottom = int.Parse(textBox12.Text.Trim());


                                // 分几页导出
                    //if (splitCount > 0)
                    //{
                    //    if(splitCount >= countOriginal){
                    //        return;
                    //    }

                    //    
                    //    MessageBox.Show(phase.ToString());
                    //    for(int pItem = 0; pItem < phase; pItem ++){
                    //         MessageBox.Show(pItem.ToString());


                int exportCount = 0;
                int splitCount = 0;

                if (!textBox8.Text.Trim().Equals("") && int.Parse(textBox8.Text.Trim()) > 0)
                {
                    exportCount = int.Parse(textBox8.Text.Trim());
                }
                if (!textBox10.Text.Trim().Equals("") && int.Parse(textBox10.Text.Trim()) > 0)
                {
                    splitCount = int.Parse(textBox10.Text.Trim());
                }
             
                if( splitCount > 0 && splitCount < item.Count){
                    int phase = (int)Math.Ceiling((double)item.Count / (double)splitCount);
                    //MessageBox.Show(((double)item.Count / (double)splitCount).ToString(), phase.ToString());
                    Color masterColor = Color.Empty;
                    for(int currentPhase = 0; currentPhase < phase; currentPhase ++){
                        
                        String[] arr = { };
                        if (!textBox5.Text.Trim().Equals(""))
                        {
                            arr = textBox5.Text.Trim().Split(char.Parse(","), char.Parse(" "), char.Parse("，")).ToArray();
                        }
                        if (false)
                        {


                        }
                        else
                        {
                            //PowerPoint.SlideRange srange = item.R;
                            int count = item.Count;
                            //if (exportCount > 0)
                            //{
                            //    if (exportCount >= item.Count)
                            //    {

                            //    }
                            //    else
                            //    {
                            //        count = exportCount;
                            //    }
                            //}
                            if(currentPhase == phase -1){
                                if(item.Count % splitCount == 0){
                                    count = splitCount;
                                }else{
                                    count = item.Count % splitCount;
                                }
                            }else if(currentPhase < phase -1){
                                count = splitCount;
                            }

                            int[] index = new int[count];                      //index数组用于将页面的选择顺序强制从前到后
                            for (int i = 1; i <= count; i++)
                            {
                                index[i - 1] = item[currentPhase * splitCount + i].SlideIndex;
                            }
                            Array.Reverse(index);
                            Array.Sort(index);
                            //item.Unselect();
                            //slides.Range(index).Select();

                            int[,] arrnew = new int[index.Count(), 2];             //arrnew用于初始化，arrnew[i,0]是所有页面的序号，arrnew[i,1]默认是小图
                            for (int i = 0; i < index.Count(); i++)
                            {
                                arrnew[i, 0] = index[i];
                                arrnew[i, 1] = 0;
                            }

                            if (arr.Count() > 0)
                            {
                                if (arr.Count() == 1 && arr[0] == "n")
                                {
                                    arr[0] = count.ToString();
                                }
                                else
                                {
                                    for (int i = 0; i < arr.Count(); i++)
                                    {
                                        if (arr[i] == "n")
                                        {
                                            arr[i] = count.ToString();
                                        }
                                    }
                                }
                            }

                            int arrprow1 = 0;
                            for (int i = 0; i < arrnew.Length / 2; i++)                  //根据数组arr标记arrnew[i,1]中哪些页面是大图
                            {
                                if (prow == 1)
                                {
                                    arrnew[i, 1] = 1;
                                    arrprow1 += 1;
                                }
                                else
                                {
                                    for (int j = 0; j < arr.Count(); j++)
                                    {
                                        if (arrnew[i, 0] == int.Parse(arr[j]))
                                        {
                                            arrnew[i, 1] = 1;
                                            arrprow1 += 1;
                                        }
                                    }
                                }
                            }

                            // 上下左右边距
                            //int wlarge = wresolution - pspac2 * 2;         //根据用户设置的分辨率，计算大图和小图的宽度和高度
                            int wlarge = wresolution - leftRight - right;         //根据用户设置的分辨率，计算大图和小图的宽度和高度

                            int hlarge = (int)(wlarge * h / w);
                            // 上下左右边距
                            //int wsmall = (int)((wlarge - (prow - 1) * pspac) / prow);
                            int wsmall = (int)((wlarge - (prow - 1) * pspac) / prow);

                            int hsmall = (int)(wsmall * h / w);

                            int arrcount = arrnew.Length / 2;
                            int[,] narr = new int[arrcount, 2];  //narr[0,0]是大图小图标识、narr[0,1]是水平序号
                            int wcount = 0;                      //wcount是水平序号、hcount是垂直序号、hscount是垂直方向上小图的行数、hscan是小图是否重新起一行
                            int hcount = 0;
                            int hscount = 0;
                            int hscan = 0;
                            for (int i = 0; i < arrcount; i++)
                            {
                                if (arrnew[i, 1] == 1)
                                {
                                    narr[i, 0] = 1;
                                    hcount += 1;
                                    wcount = 0;
                                    hscan = 0;
                                }
                                else
                                {
                                    narr[i, 0] = 0;
                                    if (wcount == 0)
                                    {
                                        if (hscan == 0)
                                        {
                                            narr[i, 1] = wcount;
                                            wcount += 1;
                                            hscount += 1;
                                            hcount += 1;
                                        }
                                        else
                                        {
                                            wcount += 1;
                                            narr[i, 1] = wcount;
                                            wcount += 1;
                                        }
                                    }
                                    else
                                    {
                                        if (wcount < prow)
                                        {
                                            narr[i, 1] = wcount;
                                            wcount += 1;
                                        }
                                        else
                                        {
                                            wcount = 0;
                                            narr[i, 1] = wcount;
                                            hscount += 1;
                                            hcount += 1;
                                            hscan = 1;
                                        }
                                    }
                                }
                            }
                            int _w = wresolution;
                            // 上下左右边距
                            //int _h = hlarge * arrprow1 + hsmall * hscount + pspac * (hcount - 1) + pspac2 * 2;
                            int _h = hlarge * arrprow1 + hsmall * hscount + pspac * (hcount - 1) + top + bottom;


                            //MessageBox.Show(_w.ToString(), _h.ToString());
                            Bitmap bmp0 = new Bitmap(_w, _h);    //计算长图的尺寸、设置长图的分辨率
                            float dpi = Properties.Settings.Default.dpi;
                            bmp0.SetResolution(dpi, dpi);

                            //String name = name_list[(int)_index];
                            String name = _ppt.Name;


                            //根据演示文稿的文件名创建长图文件夹
                            if (name.Contains(".pptx"))
                            {
                                name = name.Replace(".pptx", "");
                            }
                            else if (name.Contains(".ppt"))
                            {
                                name = name.Replace(".ppt", "");
                            }
                            string cPath = textBox7.Text + @"\" + name + @" 的幻灯片\";
                            string delPath = textBox7.Text + @"\" + name + @" 的幻灯片";

                            if (Directory.Exists(cPath))
                            {
                                new DirectoryInfo(cPath).Delete(true);

                            }
                            Directory.CreateDirectory(cPath);


                            if (checkBox1.Checked)
                            {

                                PowerPoint.Slide nslide = item[1];
                                nslide.Export(textBox7.Text + @"\" + _ppt.Name + ".png", "png", wresolution, (int)(wresolution * h / w));

                                item = null;
                                _ppt.Close();
                                return;
                            }

                            string fistSliderPath = "";
                            for (int i = 1; i <= count; i++)                //导出所选的页面为图片
                            {
                                PowerPoint.Slide nslide = item[currentPhase * splitCount + i];
                                string shname = name + "_临时_" + i;
                                nslide.Export(cPath + shname + ".png", "png", wresolution, (int)(wresolution * h / w));
                                if(i == 1){
                                   fistSliderPath = cPath + shname + ".png";
                                }
                            }

                            //if (Directory.Exists(delPath))
                            //{
                            //    new DirectoryInfo(delPath).Delete(true);
                            //}
                            //_ppt.SaveCopyAs(delPath, PowerPoint.PpSaveAsFileType.ppSaveAsPNG, Office.MsoTriState.msoTriStateMixed);

                            Graphics g = Graphics.FromImage(bmp0);
                            if (!checkBox3.Checked)                                         //设置长图的底色
                            {
                                //SolidBrush _br = new SolidBrush(panel1.BackColor);
                                //g.FillRectangle(_br, 0, 0, bmp0.Width, bmp0.Height);


                                SolidBrush _br = new SolidBrush(Color.White);
                                g.FillRectangle(_br, 0, 0, bmp0.Width, bmp0.Height);


                                if(checkBox2.Checked){
                                    int autoBgOpacity = int.Parse(textBox11.Text);
                                    if(autoBgOpacity <= 0 || autoBgOpacity > 255){
                                        autoBgOpacity = 100;
                                    }
                                    if(!fistSliderPath.Equals("")){
                                        if(masterColor == Color.Empty){
                                            
                                            masterColor = getMasterColor(_ppt, autoBgOpacity, fistSliderPath);
                                        }
                                        _br = new SolidBrush(masterColor);
                                        g.FillRectangle(_br, 0, 0, bmp0.Width, bmp0.Height);
                                    }
                                }else{
                                    _br = new SolidBrush(panel1.BackColor);
                                    g.FillRectangle(_br, 0, 0, bmp0.Width, bmp0.Height);
                                }

                            }

                            // 上下边距
                            //int ny = pspac2;
                            int ny = top;

                            int sc = 1;
                            for (int i = 1; i <= count; i++)                                //读取之前导出的临时图片，根据之前的narr数组计算该图片所在的位置和尺寸
                            {
                                string shname2 = name + "_临时_" + i + @".png";
                                //string shname2 = @"幻灯片" + i + @".PNG";

                                Bitmap bmp1 = new Bitmap(cPath + shname2);
                                int x = 0;
                                int y = 0;
                                int wd = 0;
                                int ht = 0;

                                if (narr[i - 1, 0] == 1)
                                {
                                    wd = wlarge;
                                    ht = hlarge;

                                    // 上下边距
                                    //x = pspac2;
                                    x = leftRight;

                                    y = y + ny;
                                    ny = ny + ht + pspac;
                                }
                                else
                                {
                                    wd = wsmall;
                                    ht = hsmall;
                                    // 上下边距
                                    //x = pspac2 + narr[i - 1, 1] * (wd + pspac);
                                    x = leftRight + narr[i - 1, 1] * (wd + pspac);

                                    y = y + ny;
                                    if (sc < prow)
                                    {
                                        if (i < count && narr[i, 0] == 1)
                                        {
                                            sc = 1;
                                            // 上下边距
                                            //ny = ny + ht + pspac;

                                            ny = ny + ht + pspac;
                                        }
                                        else
                                        {
                                            sc += 1;
                                        }
                                    }
                                    else
                                    {
                                        sc = 1;
                                        ny = ny + ht + pspac;
                                    }
                                }

                                if(false || checkBox4.Checked){
                                    if(File.Exists(cPath + "d1d.png")){
                                        //FileInfo _fileInfo = new FileInfo(cPath + "d1d.png");
                                        //_fileInfo.Delete();
                                    }else{
                                        PowerPoint.Shape shapeOrigin = app.ActivePresentation.Slides[1].Shapes.AddShape(
                                        Office.MsoAutoShapeType.msoShapeRectangle, 0, 0, bmp1.Width / 2f, bmp1.Height / 2f);
                                        shapeOrigin.Export(cPath + "shapeOrigin.png", PowerPoint.PpShapeFormat.ppShapeFormatPNG);
                                        shapeOrigin.Delete();

                                        PowerPoint.Shape shape = app.ActivePresentation.Slides[1].Shapes.AddShape(
                                        Office.MsoAutoShapeType.msoShapeRectangle, 0, 0, bmp1.Width / 2f, bmp1.Height / 2f);
                                        float OffsetX = (float)numericUpDown2.Value;
                                        float OffsetY = (float)numericUpDown3.Value;
                                        float Blur = (float)numericUpDown5.Value;
                                        float Transparency = (float)numericUpDown4.Value/10;

                                        //shape.Shadow.OffsetX = (float)bmp1.Width/(200.0f/10.0f);
                                        //shape.Shadow.OffsetY = (float)bmp1.Width/(200.0f/10.0f);
                                        shape.Shadow.Blur = Blur;
                                        shape.Shadow.Transparency = Transparency;
                                        //shape.Width = "px";
                                        //shape.Shadow.Type = Office.MsoShadowType.msoShadow9;
                                        shape.Shadow.Style = Office.MsoShadowStyle.msoShadowStyleOuterShadow;
                                        shape.Shadow.Visible = MsoTriState.msoCTrue;
                                        shape.Shadow.OffsetX = OffsetX;
                                        shape.Shadow.OffsetY = OffsetY;
                                        //shape.Shadow.Blur = 5.0f;
                                        shape.Export(cPath + "d1d.png", PowerPoint.PpShapeFormat.ppShapeFormatPNG);
                                        shape.Delete();

                                    }
                                    Bitmap _bitMapOrigin = new Bitmap(cPath + "shapeOrigin.png");
                                    Bitmap _bitMap = new Bitmap(cPath + "d1d.png");
                                    float w_Ratio = (float)_bitMapOrigin.Width /(float)_bitMap.Width ;
                                    float h_Ratio = (float)_bitMapOrigin.Height /(float)_bitMap.Height ;

                                    g.DrawImage(_bitMap, x, y, wd, ht);
                                    _bitMap.Dispose();
                                    _bitMapOrigin.Dispose();

                                    g.DrawImage(bmp1, x, y, wd * w_Ratio, ht * h_Ratio); 
                                    bmp1.Dispose();
                                }else{
                                    g.DrawImage(bmp1, x, y, wd, ht); 
                                    bmp1.Dispose();
                                }

                                //g.DrawImage(bmp1, x, y, wd, ht);                      //在长图中进行绘制

                                //bmp1.Dispose();
                            }


                            //如果为空，则不设置水印
                            if (!textBox6.Text.Trim().Equals(""))
                            {
                                int angle = Properties.Settings.Default.water_a;
                                float rowWaterCount = Properties.Settings.Default.water_r;
                                float colWaterCount = Properties.Settings.Default.water_c;

                                Bitmap textBitmap = drawText(bmp0, angle, rowWaterCount, colWaterCount, opacity);

                                ImageAttributes imageAtt = GetAlphaImgAttr(50);

                                Rectangle rectangle = new Rectangle(0, 0, textBitmap.Width, textBitmap.Height);

                                g.DrawImage(textBitmap, rectangle, -textBitmap.Width / 2, -textBitmap.Height / 2, textBitmap.Width, textBitmap.Height, GraphicsUnit.Pixel, imageAtt);
                                textBitmap.Dispose();
                            }

                            new DirectoryInfo(delPath).Delete(true);
                            int k = 0;
                            //File.Delete(delPath);
                            try
                            {
                                System.IO.DirectoryInfo dir = new System.IO.DirectoryInfo(textBox7.Text);   //保存长图为png或jpg
                                k = dir.GetFiles().Length + 1;
                            }
                            catch (Exception e)
                            {
                                MessageBox.Show(e.ToString());
                                bmp0.Dispose();
                                item = null;
                                _ppt.Close();
                                return;
                            }

                            if (!Directory.Exists(textBox7.Text + @"\" + name))
                            {
                                Directory.CreateDirectory(textBox7.Text + @"\" + name);
                            }
                            string ddd = textBox7.Text;
                            if(checkBox6.Checked){
                                ddd =  Path.GetDirectoryName(ppt_path);
                            }

                            if (checkBox3.Checked)
                            {
                                
                                string save_name = ddd + @"\" + name + @"\"+ "第"+ (currentPhase+1)+  "块" + ".png";
                                
                                
                                if (File.Exists(save_name))
                                {
                                    new FileInfo(save_name).Delete();
                                    //save_name = textBox7.Text + @"\" + name + "_" + k + @"\"+ "第"+ (currentPhase+1)+  "块" + ".png";
                                }

                                bmp0.Save(save_name, ImageFormat.Png);
                            }
                            else
                            {
                                string save_name = ddd + @"\" + name + @"\"+ "第"+ (currentPhase+1)+  "块" + ".jpg";
                                if (File.Exists(save_name))
                                {
                                    new FileInfo(save_name).Delete();
                                    //save_name = textBox7.Text + @"\" + name + "_" + k + @"\"+ "第"+ (currentPhase+1)+  "块" + ".jpg";
                                }

                                //MessageBox.Show(save_name);
                                bmp0.Save(save_name, ImageFormat.Jpeg);
                            }

                            bmp0.Dispose();
                            
                           
                        }
                        
                    }
                    item = null;
                    _ppt.Close();
                }else{
                    String[] arr = { };
                    if (!textBox5.Text.Trim().Equals(""))
                    {
                        arr = textBox5.Text.Trim().Split(char.Parse(","), char.Parse(" "), char.Parse("，")).ToArray();
                    }
                    if (false)
                    {


                    }
                    else
                    {
                        //PowerPoint.SlideRange srange = item.R;
                        int count = item.Count;
                        if (exportCount > 0)
                        {
                            if (exportCount >= item.Count)
                            {

                            }
                            else
                            {
                                count = exportCount;
                            }
                        }

                        int[] index = new int[count];                      //index数组用于将页面的选择顺序强制从前到后
                        for (int i = 1; i <= count; i++)
                        {
                            index[i - 1] = item[i].SlideIndex;
                        }
                        Array.Reverse(index);
                        Array.Sort(index);
                        //item.Unselect();
                        //slides.Range(index).Select();

                        int[,] arrnew = new int[index.Count(), 2];             //arrnew用于初始化，arrnew[i,0]是所有页面的序号，arrnew[i,1]默认是小图
                        for (int i = 0; i < index.Count(); i++)
                        {
                            arrnew[i, 0] = index[i];
                            arrnew[i, 1] = 0;
                        }

                        if (arr.Count() > 0)
                        {
                            if (arr.Count() == 1 && arr[0] == "n")
                            {
                                arr[0] = count.ToString();
                            }
                            else
                            {
                                for (int i = 0; i < arr.Count(); i++)
                                {
                                    if (arr[i] == "n")
                                    {
                                        arr[i] = count.ToString();
                                    }
                                }
                            }
                        }

                        int arrprow1 = 0;
                        for (int i = 0; i < arrnew.Length / 2; i++)                  //根据数组arr标记arrnew[i,1]中哪些页面是大图
                        {
                            if (prow == 1)
                            {
                                arrnew[i, 1] = 1;
                                arrprow1 += 1;
                            }
                            else
                            {
                                for (int j = 0; j < arr.Count(); j++)
                                {
                                    if (arrnew[i, 0] == int.Parse(arr[j]))
                                    {
                                        arrnew[i, 1] = 1;
                                        arrprow1 += 1;
                                    }
                                }
                            }
                        }
                        

                        
                        // 上下左右边距
                        //int wlarge = wresolution - pspac2 * 2;         //根据用户设置的分辨率，计算大图和小图的宽度和高度
                        int wlarge = wresolution - leftRight - right;         //根据用户设置的分辨率，计算大图和小图的宽度和高度

                        int hlarge = (int)(wlarge * h / w);

                         // 上下左右边距
                        //int wsmall = (int)((wlarge - (prow - 1) * pspac) / prow);
                        int wsmall = (int)((wlarge - (prow - 1) * pspac) / prow);

                        int hsmall = (int)(wsmall * h / w);

                        int arrcount = arrnew.Length / 2;
                        int[,] narr = new int[arrcount, 2];  //narr[0,0]是大图小图标识、narr[0,1]是水平序号
                        int wcount = 0;                      //wcount是水平序号、hcount是垂直序号、hscount是垂直方向上小图的行数、hscan是小图是否重新起一行
                        int hcount = 0;
                        int hscount = 0;
                        int hscan = 0;
                        for (int i = 0; i < arrcount; i++)
                        {
                            if (arrnew[i, 1] == 1)
                            {
                                narr[i, 0] = 1;
                                hcount += 1;
                                wcount = 0;
                                hscan = 0;
                            }
                            else
                            {
                                narr[i, 0] = 0;
                                if (wcount == 0)
                                {
                                    if (hscan == 0)
                                    {
                                        narr[i, 1] = wcount;
                                        wcount += 1;
                                        hscount += 1;
                                        hcount += 1;
                                    }
                                    else
                                    {
                                        wcount += 1;
                                        narr[i, 1] = wcount;
                                        wcount += 1;
                                    }
                                }
                                else
                                {
                                    if (wcount < prow)
                                    {
                                        narr[i, 1] = wcount;
                                        wcount += 1;
                                    }
                                    else
                                    {
                                        wcount = 0;
                                        narr[i, 1] = wcount;
                                        hscount += 1;
                                        hcount += 1;
                                        hscan = 1;
                                    }
                                }
                            }
                        }
                        int _w = wresolution;

                        // 上下左右边距
                        //int _h = hlarge * arrprow1 + hsmall * hscount + pspac * (hcount - 1) + pspac2 * 2;
                        int _h = hlarge * arrprow1 + hsmall * hscount + pspac * (hcount - 1) + top + bottom;

                        Bitmap bmp0 = new Bitmap(_w, _h);    //计算长图的尺寸、设置长图的分辨率
                        float dpi = Properties.Settings.Default.dpi;
                        bmp0.SetResolution(dpi, dpi);

                        //String name = name_list[(int)_index];
                        String name = _ppt.Name;


                        //根据演示文稿的文件名创建长图文件夹
                        if (name.Contains(".pptx"))
                        {
                            name = name.Replace(".pptx", "");
                        }
                        else if (name.Contains(".ppt"))
                        {
                            name = name.Replace(".ppt", "");
                        }
                        string cPath = textBox7.Text + @"\" + name + @" 的幻灯片\";
                        string delPath = textBox7.Text + @"\" + name + @" 的幻灯片";

                        if (Directory.Exists(cPath))
                        {
                            new DirectoryInfo(cPath).Delete(true);

                        }
                        Directory.CreateDirectory(cPath);


                        if (checkBox1.Checked)
                        {

                            PowerPoint.Slide nslide = item[1];
                            nslide.Export(textBox7.Text + @"\" + _ppt.Name + ".png", "png", wresolution, (int)(wresolution * h / w));

                            item = null;
                            _ppt.Close();
                            return;
                        }

                        string fistSliderPath = "";
                        for (int i = 1; i <= count; i++)                //导出所选的页面为图片
                        {
                            PowerPoint.Slide nslide = item[i];
                            string shname = name + "_临时_" + i;
                            nslide.Export(cPath + shname + ".png", "png", wresolution, (int)(wresolution * h / w));
                            if(i == 1){
                                fistSliderPath = cPath + shname + ".png";
                            }

                        }

                        //if (Directory.Exists(delPath))
                        //{
                        //    new DirectoryInfo(delPath).Delete(true);
                        //}
                        //_ppt.SaveCopyAs(delPath, PowerPoint.PpSaveAsFileType.ppSaveAsPNG, Office.MsoTriState.msoTriStateMixed);

                        Graphics g = Graphics.FromImage(bmp0);
                        if (!checkBox3.Checked)                                         //设置长图的底色
                        {
                            SolidBrush _br = new SolidBrush(Color.White);
                            g.FillRectangle(_br, 0, 0, bmp0.Width, bmp0.Height);


                            if(checkBox2.Checked){
                                int autoBgOpacity = int.Parse(textBox11.Text);
                                if(autoBgOpacity <= 0 || autoBgOpacity > 255){
                                    autoBgOpacity = 100;
                                }
                                if(!fistSliderPath.Equals("")){
                                     Color masterColor = getMasterColor(_ppt, autoBgOpacity, fistSliderPath);
                                    _br = new SolidBrush(masterColor);
                                    g.FillRectangle(_br, 0, 0, bmp0.Width, bmp0.Height);
                                }
                            }else{
                                _br = new SolidBrush(panel1.BackColor);
                                g.FillRectangle(_br, 0, 0, bmp0.Width, bmp0.Height);
                            }
                            
                        }

                        // 上下边距
                        //int ny = pspac2;
                        int ny = top;

                        int sc = 1;
                        for (int i = 1; i <= count; i++)                                //读取之前导出的临时图片，根据之前的narr数组计算该图片所在的位置和尺寸
                        {
                            string shname2 = name + "_临时_" + i + @".png";
                            //string shname2 = @"幻灯片" + i + @".PNG";

                            Bitmap bmp1 = new Bitmap(cPath + shname2);
                            int x = 0;
                            int y = 0;
                            int wd = 0;
                            int ht = 0;

                            if (narr[i - 1, 0] == 1)
                            {
                                wd = wlarge;
                                ht = hlarge;
                                
                                // 上下边距
                                //x = pspac2;
                                x = leftRight;

                                y = y + ny;
                                ny = ny + ht + pspac;
                            }
                            else
                            {
                                wd = wsmall;
                                ht = hsmall;
                                // 上下边距
                                //x = pspac2 + narr[i - 1, 1] * (wd + pspac);
                                x = leftRight + narr[i - 1, 1] * (wd + pspac);

                                y = y + ny;
                                if (sc < prow)
                                {
                                    if (i < count && narr[i, 0] == 1)
                                    {
                                        sc = 1;
                                        // 上下边距
                                        //ny = ny + ht + pspac;
                                        ny = ny + ht + pspac;
                                    }
                                    else
                                    {
                                        sc += 1;
                                    }
                                }
                                else
                                {
                                    sc = 1;
                                    ny = ny + ht + pspac;
                                }
                            }

                            
                            if(false || checkBox4.Checked){
                                if(File.Exists(cPath + "d1d.png")){
                                    //FileInfo _fileInfo = new FileInfo(cPath + "d1d.png");
                                    //_fileInfo.Delete();
                                }else{
                                    PowerPoint.Shape shapeOrigin = app.ActivePresentation.Slides[1].Shapes.AddShape(
                                    Office.MsoAutoShapeType.msoShapeRectangle, 0, 0, bmp1.Width / 2f, bmp1.Height / 2f);
                                    shapeOrigin.Export(cPath + "shapeOrigin.png", PowerPoint.PpShapeFormat.ppShapeFormatPNG);
                                    shapeOrigin.Delete();

                                    PowerPoint.Shape shape = app.ActivePresentation.Slides[1].Shapes.AddShape(
                                    Office.MsoAutoShapeType.msoShapeRectangle, 0, 0, bmp1.Width / 2f, bmp1.Height / 2f);
                                    float OffsetX = (float)numericUpDown2.Value;
                                    float OffsetY = (float)numericUpDown3.Value;
                                    float Blur = (float)numericUpDown5.Value;
                                    float Transparency = (float)numericUpDown4.Value/10;
                                    //shape.Shadow.OffsetX = (float)bmp1.Width/(200.0f/10.0f);
                                    //shape.Shadow.OffsetY = (float)bmp1.Width/(200.0f/10.0f);
                                    shape.Shadow.Blur = Blur;
                                    shape.Shadow.Transparency = Transparency;
                                    //shape.Width = "px";
                                    //shape.Shadow.Type = Office.MsoShadowType.msoShadow9;
                                    shape.Shadow.Style = Office.MsoShadowStyle.msoShadowStyleOuterShadow;
                                    shape.Shadow.Visible = MsoTriState.msoCTrue;
                                    shape.Shadow.OffsetX = OffsetX;
                                    shape.Shadow.OffsetY = OffsetY;
                                    //shape.Shadow.Blur = 5.0f;
                                    shape.Export(cPath + "d1d.png", PowerPoint.PpShapeFormat.ppShapeFormatPNG);
                                    shape.Delete();

                                }
                                Bitmap _bitMapOrigin = new Bitmap(cPath + "shapeOrigin.png");
                                Bitmap _bitMap = new Bitmap(cPath + "d1d.png");
                                float w_Ratio = (float)_bitMapOrigin.Width /(float)_bitMap.Width ;
                                float h_Ratio = (float)_bitMapOrigin.Height /(float)_bitMap.Height ;

                                g.DrawImage(_bitMap, x, y, wd, ht);
                                _bitMap.Dispose();
                                _bitMapOrigin.Dispose();

                                g.DrawImage(bmp1, x, y, wd * w_Ratio, ht * h_Ratio); 
                                bmp1.Dispose();
                            }else{
                                g.DrawImage(bmp1, x, y, wd, ht); 
                                bmp1.Dispose();
                            }
                        }


                        //如果为空，则不设置水印
                        if (!textBox6.Text.Trim().Equals(""))
                        {
                            int angle = Properties.Settings.Default.water_a;
                            float rowWaterCount = Properties.Settings.Default.water_r;
                            float colWaterCount = Properties.Settings.Default.water_c;

                            Bitmap textBitmap = drawText(bmp0, angle, rowWaterCount, colWaterCount, opacity);

                            ImageAttributes imageAtt = GetAlphaImgAttr(opacity);

                            Rectangle rectangle = new Rectangle(0,  0, textBitmap.Width, textBitmap.Height);

                            g.DrawImage(textBitmap, rectangle, textBitmap.Width/2, textBitmap.Height/2, textBitmap.Width, textBitmap.Height, GraphicsUnit.Pixel, imageAtt);


                            //g.DrawImage(textBitmap, -textBitmap.Width / 2, -textBitmap.Height / 2, textBitmap.Width, textBitmap.Height);
                            textBitmap.Dispose();
                        }

                        new DirectoryInfo(delPath).Delete(true);
                        int k = 0;
                        //File.Delete(delPath);
                        try
                        {
                            System.IO.DirectoryInfo dir = new System.IO.DirectoryInfo(textBox7.Text);   //保存长图为png或jpg
                            k = dir.GetFiles().Length + 1;
                        }
                        catch (Exception e)
                        {
                            MessageBox.Show(e.ToString());
                            bmp0.Dispose();
                            item = null;
                            _ppt.Close();
                            return;
                        }


                        if (checkBox3.Checked)
                        {   
                           
                            string save_name = textBox7.Text + @"\" + name + ".png";
                            
                            if(checkBox6.Checked){
                                save_name =  Path.GetDirectoryName(ppt_path) + @"\" + name + ".png";
                            }
                            if (File.Exists(save_name))
                            {   
                                if(checkBox6.Checked){
                                    save_name =  Path.GetDirectoryName(ppt_path) + @"\" + name + "_" + k + ".png";
                                }else{
                                    save_name = textBox7.Text + @"\" + name + "_" + k + ".png";
                                }
                                
                                
                            }

                            bmp0.Save(save_name, ImageFormat.Png);
                        }
                        else
                        {
                            string save_name = textBox7.Text + @"\" + name + ".jpg";
                            if(checkBox6.Checked){
                                save_name =  Path.GetDirectoryName(ppt_path) + @"\" + name + ".jpg";
                            }

                            if (File.Exists(save_name))
                            {
                                if(checkBox6.Checked){
                                    save_name =  Path.GetDirectoryName(ppt_path) + @"\" + name + "_" + k + ".jpg";
                                }else{
                                    save_name = textBox7.Text + @"\" + name + "_" + k + ".jpg";
                                }
                            }

                            //MessageBox.Show(save_name);
                            bmp0.Save(save_name, ImageFormat.Jpeg);
                        }

                        bmp0.Dispose();
                        item = null;
                        _ppt.Close();
                    }
                }
            }
            catch (Exception e)
            {   
                if(_ppt != null){
                    _ppt.Close();
                    this.Invoke(setList, "【" + _ppt.Name + "】\n出错信息：" + e.ToString() + "\n\n");
                }else{
                    this.Invoke(setList, "出错信息：" + e.ToString() + "\n\n");
                }
            }
        }


        private Bitmap drawText(Bitmap bmp0, int angle, float rowWaterCount, float colWaterCount, int opacity)
        {
            int bw = bmp0.Width * 4, bh = bmp0.Height * 4;

            String text = textBox6.Text;

            Font font = label23.Font;
            
            //Color col = Color.FromArgb(50, panel2.BackColor);

            //SolidBrush br = new SolidBrush(col);
            SolidBrush br = new SolidBrush(panel2.BackColor);

            Bitmap textBigmap = new Bitmap(bw, bh);
            float dpi = Properties.Settings.Default.dpi;
            textBigmap.SetResolution(dpi, dpi);
            Graphics textBigmapG = Graphics.FromImage(textBigmap);
            //SizeF sf = textBigmapG.MeasureString(text, font);


            float height = bh / colWaterCount;
            float width = bw / rowWaterCount;

            float x0 = (float)(width * 0.2);
            float y0 = (float)(height * 0.5);
            textBigmapG.TranslateTransform(textBigmap.Width / 2, textBigmap.Height / 2);
            textBigmapG.RotateTransform(angle);
            textBigmapG.TranslateTransform(-textBigmap.Width / 2, -textBigmap.Height / 2);
            textBigmapG.SmoothingMode = SmoothingMode.HighSpeed;
            textBigmapG.CompositingQuality = CompositingQuality.HighSpeed;
            textBigmapG.TextRenderingHint = TextRenderingHint.AntiAlias;
            //textBigmapG.InterpolationMode = InterpolationMode.HighQualityBicubic;
            //textBigmapG.CompositingQuality = CompositingQuality.HighQuality;
            //textBigmapG.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

            for (int i = 0; i < (int)(bh / height); i++)
            {
                for (int j = 0; j < (int)(bw / width); j++)
                {
                    //height
                    float strX = x0 + width * j;
                    float strY = y0 + height * i;

                    //if (strX < textBigmap.Width && strY < textBigmap.Height)
                    //{
                    textBigmapG.DrawString(text, font, br, strX, strY);
                    //}

                }
            }

            textBigmapG.Dispose();
            //textBigmap.MakeTransparent(Color.Black);
            return textBigmap;


        }
        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void panel1_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                if (this.colorDialog1.ShowDialog() == DialogResult.OK)
                {

                    Properties.Settings.Default.bg_color = colorDialog1.Color.ToArgb();
                    Properties.Settings.Default.Save();

                    this.panel1.BackColor = colorDialog1.Color;
                }
            }
            if (e.Button == MouseButtons.Right)
            {
                //this.transparentPanel1.Width = this.Width;
                //this.transparentPanel1.Height = this.Height;
                //this.transparentPanel1.Location = new Point(0, 0);
                //this.transparentPanel1.Visible = true;


                //MessageBox.Show(this.transparentPanel1.Visible.ToString());
                timer1.Start();
                //if (this.colorDialog1.ShowDialog() == DialogResult.OK)
                //{
                //    Properties.Settings.Default.bg_color = colorDialog1.Color.ToArgb();
                //    Properties.Settings.Default.Save();
                //    this.panel1.BackColor = colorDialog1.Color;
                //}
            }
        }
        private void panel2_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                if (this.colorDialog1.ShowDialog() == DialogResult.OK)
                {
                    Properties.Settings.Default.logo_color = colorDialog1.Color.ToArgb();
                    Properties.Settings.Default.Save();
                    this.panel2.BackColor = colorDialog1.Color;
                }
            }
            if (e.Button == MouseButtons.Right)
            {
                

                //this.transparentPanel1.Visible = true;
                //this.transparentPanel1.Location = new Point(0, 0);
                //this.transparentPanel1.Width = this.Width;
                //this.transparentPanel1.Height = this.Height;

                //this.transparentPanel1.Visible = true;

                //MessageBox.Show(this.transparentPanel1.Visible.ToString());
                timer2.Start();

                //if (this.colorDialog1.ShowDialog() == DialogResult.OK)
                //{
                //    Properties.Settings.Default.logo_color = colorDialog1.Color.ToArgb();
                //    Properties.Settings.Default.Save();
                //    this.panel2.BackColor = colorDialog1.Color;
                //}
            }
        }
        private void timer2_Tick(object sender, EventArgs e)
        {
            Point p = new Point(MousePosition.X, MousePosition.Y);
            Bitmap shotImage = new Bitmap(1, 1);
            Graphics dc = Graphics.FromImage(shotImage);
            dc.CopyFromScreen(p, new Point(0, 0), shotImage.Size);
            Color c = shotImage.GetPixel(0, 0);

            panel2.BackColor = c;
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            Point p = new Point(MousePosition.X, MousePosition.Y);
            Bitmap shotImage = new Bitmap(1, 1);
            Graphics dc = Graphics.FromImage(shotImage);
            dc.CopyFromScreen(p, new Point(0, 0), shotImage.Size);
            Color c = shotImage.GetPixel(0, 0);

            panel1.BackColor = c;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string name = app.ActivePresentation.Name;
            if (name.Contains(".pptx"))
            {
                name = name.Replace(".pptx", "");
            }
            if (name.Contains(".ppt"))
            {
                name = name.Replace(".ppt", "");
            }
            string cPath = app.ActivePresentation.Path + @"\" + name + @" 的幻灯片\";
            if (!Directory.Exists(cPath))
            {
                MessageBox.Show("不存在拼图文件夹");
            }
            else
            {
                System.IO.DirectoryInfo dir = new System.IO.DirectoryInfo(cPath);
                FileInfo[] files = dir.GetFiles();
                foreach (FileInfo file in files)
                {
                    file.Delete();
                }
                MessageBox.Show("已清空拼图文件夹");
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string name = app.ActivePresentation.Name;
            if (name.Contains(".pptx"))
            {
                name = name.Replace(".pptx", "");
            }
            if (name.Contains(".ppt"))
            {
                name = name.Replace(".ppt", "");
            }
            string cPath = app.ActivePresentation.Path + @"\" + name + @" 的幻灯片\";
            if (!Directory.Exists(cPath))
            {
                MessageBox.Show("不存在拼图文件夹");
            }
            else
            {
                Directory.Delete(cPath, true);
                MessageBox.Show("已删除拼图文件夹");
            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            try
            {
                if (int.Parse(textBox2.Text.Trim()) <= 0)
                {
                    //MessageBox.Show("请输入有效的列数");
                    //textBox2.Text = "2";
                }
                else
                {
                    Properties.Settings.Default.col = int.Parse(textBox2.Text.Trim());
                    Properties.Settings.Default.Save();
                }
            }
            catch (FormatException)
            {
                //MessageBox.Show("请输入有效的列数");
                //textBox2.Text = "2";
            }


        }

        private void label10_Click(object sender, EventArgs e)
        {
            string[] w = label10.Text.Split(char.Parse("*")).ToArray();
            textBox1.Text = w[0];
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Picture_Jigsaw.ActiveForm.Close();
            Globals.Ribbons.Ribbon1.button1.Enabled = true;
        }



        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void toolTip1_Popup(object sender, PopupEventArgs e)
        {

        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            try
            {
                if (int.Parse(textBox3.Text.Trim()) <= 0)
                {
                    //MessageBox.Show("请输入有效的内间隔");

                    //textBox3.Text = "4";
                }
                else
                {
                    Properties.Settings.Default.inner_space = int.Parse(textBox3.Text.Trim());
                    Properties.Settings.Default.Save();
                }
            }
            catch (FormatException)
            {
                //MessageBox.Show("请输入有效的内间隔");
                //textBox3.Text = "4";
            }

        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.big_num = textBox5.Text;
            Properties.Settings.Default.Save();
        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            try
            {
                if (int.Parse(textBox4.Text.Trim()) <= 0)
                {
                    //MessageBox.Show("请输入有效的外间隔");

                    //textBox4.Text = "4";
                }
                else
                {
                    Properties.Settings.Default.out_space = int.Parse(textBox4.Text.Trim());
                    
                    Properties.Settings.Default.top = textBox4.Text.Trim();
                    Properties.Settings.Default.Save();
                }
            }
            catch (FormatException)
            {
                //MessageBox.Show("请输入有效的外间隔");
                //textBox4.Text = "4";
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            try
            {

                if (int.Parse(textBox1.Text.Trim()) <= 0)
                {
                    //MessageBox.Show("请输入有效的水平宽度");
                    //textBox1.Text = "750";
                }
                else
                {
                    Properties.Settings.Default.width = int.Parse(textBox1.Text.Trim());
                    Properties.Settings.Default.Save();
                }
            }
            catch (FormatException)
            {
                //MessageBox.Show("请输入有效的水平宽度");
                //textBox1.Text = "750";
            }
        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.logo = textBox6.Text.Trim();
            Properties.Settings.Default.Save();
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void button7_Click(object sender, EventArgs e)
        {
            int all_count = listBox1.SelectedIndices.Count;

            if (all_count < 1)
            {
                return;
            }

            for (int i = all_count - 1; i >= 0; i--)
            {
                listBox1.Items.RemoveAt(listBox1.SelectedIndices[i]);
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            listBox1.Items.Clear();
        }

        private void panel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void button1_Click_1(object sender, EventArgs e)
        {

            try
            {
                if (this.fontDialog1.ShowDialog() == DialogResult.OK)
                {
                    label23.Font = fontDialog1.Font;
                    Properties.Settings.Default.water_f = fontDialog1.Font;
                    Properties.Settings.Default.Save();
                }
            }
            catch (Exception)
            {
                MessageBox.Show("不支持这种字体");
            }
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            try
            {
                if (_thread.ThreadState == System.Threading.ThreadState.Running)
                {
                    _thread.Abort();
                    button2.Visible = false;
                }
                else
                {
                    MessageBox.Show("暂无任务可以取消");
                    
                    this.showText("处理完成", listBox1.Items.Count - 1);
                }
            }
            catch (NullReferenceException)
            {
                MessageBox.Show("暂无任务可以取消");
                this.showText("处理完成..", listBox1.Items.Count - 1);
            }
        }

        private void trackBar1_Scroll(object sender, EventArgs e)
        {
            //label27.Text = "横向"+trackBar1.Value.ToString();
            Properties.Settings.Default.water_r = trackBar1.Value;
            Properties.Settings.Default.Save();
        }

        private void trackBar2_Scroll(object sender, EventArgs e)
        {
            //label28.Text = "纵向" + trackBar2.Value.ToString();
            Properties.Settings.Default.water_c = trackBar2.Value;
            Properties.Settings.Default.Save();
        }

        private void trackBar3_Scroll(object sender, EventArgs e)
        {
            //label29.Text = "角度" + trackBar3.Value.ToString();
            Properties.Settings.Default.water_a = trackBar3.Value;
            Properties.Settings.Default.Save();
        }

        private void textBox10_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.splitCount = textBox10.Text;
            Properties.Settings.Default.Save();
        }

        private void label16_Click(object sender, EventArgs e)
        {

        }

        private void textBox9_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.opacity = textBox9.Text;
            Properties.Settings.Default.Save();
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            //String ppt_path = listBox1.Items[0].ToString();

            //PowerPoint.Presentation _ppt = app.Presentations.Open(
            //   ppt_path, Office.MsoTriState.msoTrue, Office.MsoTriState.msoFalse, Office.MsoTriState.msoFalse
            //);
            //string _extra = this.getColor(_ppt);

            
            //MsgForm msgForm = new MsgForm();
            //msgForm.Show();
            ////Color color =  Color.FromArgb(_extra);
            //msgForm.setText(getColor());
        }

        
        private Color getMasterColor(PowerPoint.Presentation _ppt, int opacity, string firstPath){
            if(!checkBox5.Checked){
                int c = _ppt.SlideMaster.Theme.ThemeColorScheme.Colors(Office.MsoThemeColorSchemeIndex.msoThemeAccent1).RGB;
                return Color.FromArgb(opacity, Color.FromArgb(c));
            }

            Bitmap bitmap = new Bitmap(firstPath);

            //色调的总和
            var sum_hue = 0d;
            //色差的阈值
            var threshold = 30;
            //计算色调总和
            for (int h = 0; h < bitmap.Height; h++)
            {
                for (int w = 0; w < bitmap.Width; w++)
                {
                    var hue = bitmap.GetPixel(w, h).GetHue();
                    sum_hue += hue;
                }
            }
            var avg_hue = sum_hue / (bitmap.Width * bitmap.Height);
 
            //色差大于阈值的颜色值
            var rgbs = new List<Color>();
            for (int h = 0; h < bitmap.Height; h++)
            {
                for (int w = 0; w < bitmap.Width; w++)
                {
                    var color = bitmap.GetPixel(w, h);
                    var hue = color.GetHue();
                    //如果色差大于阈值，则加入列表
                    if (Math.Abs(hue - avg_hue) > threshold)
                    {
                        rgbs.Add(color);
                    }
                }
            }
            bitmap.Dispose();
            if (rgbs.Count == 0)
                return Color.Black;
            //计算列表中的颜色均值，结果即为该图片的主色调
            int sum_r = 0, sum_g = 0, sum_b = 0;
            foreach (var rgb in rgbs)
            {
                sum_r += rgb.R;
                sum_g += rgb.G;
                sum_b += rgb.B;
            }
           
            return Color.FromArgb(opacity, Color.FromArgb(
                sum_r / rgbs.Count,
                sum_g / rgbs.Count,
                sum_b / rgbs.Count
            ));

        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.autoBg = checkBox2.Checked;
            Properties.Settings.Default.Save();

            if(checkBox2.Checked){
                label33.Visible = true;
                textBox11.Visible = true;

                label9.Visible = false;
                panel1.Visible = false;

            }else{
                label33.Visible = false;
                textBox11.Visible = false;

                label9.Visible = true;
                panel1.Visible = true;
            }
        }

        private void textBox11_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.autoBgOpacity = textBox11.Text;
            Properties.Settings.Default.Save();
        }

        private void label34_Click(object sender, EventArgs e)
        {

        }
        private void SaveVideoItem(Object _index, int videoWidth, int duration)
        {
            String ppt_path =  this.listBox1.Items[(int)_index].ToString();

            PowerPoint.Presentation _ppt = app.Presentations.Open(
               ppt_path, Office.MsoTriState.msoFalse, Office.MsoTriState.msoFalse, Office.MsoTriState.msoFalse
            );
 
            try
            {
                for (int i = 1; i < _ppt.Slides.Count; i++)
                {

                    PowerPoint.SlideShowTransition se = _ppt.Slides[i + 1].SlideShowTransition;
                    se.EntryEffect = PowerPoint.PpEntryEffect.ppEffectPushUp;

                    //se[-1].EffectType = PowerPoint.MsoAnimEffect.msoAnimEffectArcUp;

                }
                string path = this.textBox7.Text;
                if(checkBox6.Checked){
                    path = Path.GetDirectoryName(ppt_path);
                }
                string cuName = _ppt.Name;
                if (cuName.Contains(".pptx") || cuName.Contains(".ppt"))
                {
                    cuName = cuName.Replace(".pptx", "");
                    cuName = cuName.Replace(".ppt", "");
                }
                //string saveName = path + @"\" + _ppt + "(" + videoWidth + ")" + ".mp4";
                string saveName = path + @"\" + cuName + ".mp4";

                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }
                if(File.Exists(saveName)){
                    new FileInfo(saveName).Delete();
                }
                
                _ppt.CreateVideo(saveName, false, duration, videoWidth, 30, 85);
                //_ppt.Close();
            }
            catch (Exception e)
            {
                this.handelList("【" + _ppt.Name + "】\n出错信息：" + e.ToString() + "\n\n");
                
                _ppt.Close();
            }
        }

       
        private void resolveSliderVideo()
        {
            
        }

        private void button3_Click_2(object sender, EventArgs ea)
        {
            if(this.checkBox6.Checked == false){
                if (textBox7.TextLength == 0)
                {
                    MessageBox.Show("请设置保存路径");
                    return;
                }
                if (!Directory.Exists(textBox7.Text))
                {
                    MessageBox.Show("保存路径不正确，请重新设置");
                    return;
                }
            }

            if (listBox1.Items.Count == 0)
            {
                MessageBox.Show("请选择PPT文件");
                return;
            }


            for (int i = 0; i < listBox1.Items.Count; i++)
            {

                this.showText("正在导出第" + (i + 1) + "个视频", i);
                try
                {
                    int videoWidth = 480;
                    if (!this.comboBox1.Text.Equals(""))
                    {
                        videoWidth = int.Parse(this.comboBox1.Text);
                    }
                    int duration = 5;
                    if(this.numericUpDown1.Value > 0){
                        duration = (int)(this.numericUpDown1.Value);
                    }

                    SaveVideoItem(i, videoWidth, duration);
                }
                catch (System.Threading.ThreadAbortException e)
                {

                    MessageBox.Show("取消成功");
                    
                    this.showText("处理完成", i);
                    break;
                }
                catch (Exception e)
                {
                    // MessageBox.Show("有可能没有存储权限，请以管理员方式运行WPS或者Office, 此外，要关闭被占用的PPT \n" + " " + e.ToString());
                    //strArr.Add(listBox1.Items[i].ToString());
                    this.handelList("【" + listBox1.Items[i].ToString() + "】\n错误信息：" + e.ToString() + "\n\n");
                    //this.Invoke(setText, "第" + (i + 1) + "个文件出错", i);
                    continue;
                }

                if (i + 1 == listBox1.Items.Count)
                {
                    //this.showText("处理完成", i);
                    this.Close();
                    this.Dispose();
                    MsgForm msgForm = new MsgForm();
                    msgForm.setText("导出任务添加完成, 窗口已关闭，等Office下边的导出状态完成后，视频即可导出成功。");
                    msgForm.ShowDialog();
                    
                }
            }


            button2.Visible = true;

            // SaveAs(path + "/a.mp4", PowerPoint.PpSaveAsFileType.ppSaveAsMP4);
            //PowerPoint.Shape shape = app.ActivePresentation.Slides[1].Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, 0, 0, 200, 100);
            //shape.Shadow.OffsetX = 5f;
            //shape.Shadow.OffsetY = 5f;
            //shape.Shadow.Blur = 5f;
            //shape.Shadow.Transparency = 0.5f;

            ////shape.Shadow.Type = Office.MsoShadowType.msoShadow9;
            //shape.Shadow.Style = Office.MsoShadowStyle.msoShadowStyleOuterShadow;
            //shape.Shadow.Visible = MsoTriState.msoCTrue;


            //shape.Export(@"C:\Users\kuniq\Desktop\pptTest\dd.png", PowerPoint.PpShapeFormat.ppShapeFormatPNG, 220, 120);
            //app.ActivePresentation.Slides[1].Select();
            //app.ActiveWindow.Selection.ShapeRange[1].Export();
        }

        private void textBox12_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.bottom = textBox12.Text.Trim();
            Properties.Settings.Default.Save();
        }

        private void textBox13_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.leftRight = textBox13.Text.Trim();
            Properties.Settings.Default.Save();
        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void Picture_Jigsaw_MouseUp(object sender, MouseEventArgs e)
        {

        }

        private void Picture_Jigsaw_MouseClick(object sender, MouseEventArgs e)
        {
            this.timer1.Stop();
            this.timer2.Stop();
        }

        private void panel3_MouseClick(object sender, MouseEventArgs e)
        {
            this.timer1.Stop();
            this.timer2.Stop();
            //this.panel3.Visible = false;
        }

        private void transparentPanel1_MouseClick(object sender, MouseEventArgs e)
        {
            //this.transparentPanel1.Visible = false;
            this.timer1.Stop();
            this.timer2.Stop();
        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void textBox14_TextChanged(object sender, EventArgs e)
        {
             Properties.Settings.Default.right = textBox14.Text.Trim();
             Properties.Settings.Default.Save();
        }

        private void label45_Click(object sender, EventArgs e)
        {

        }

        private void label46_Click(object sender, EventArgs e)
        {

        }
    }
}

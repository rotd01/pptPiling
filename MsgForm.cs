using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace PiliangPPT
{
    public partial class MsgForm : Form
    {
        public MsgForm()
        {
            InitializeComponent();
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {

        }
        public void setText(String text)
        {
            this.richTextBox1.Text = text;
        }

        private void MsgForm_Load(object sender, EventArgs e)
        {
            //MessageBox.Show("ddd");
        }
    }
}

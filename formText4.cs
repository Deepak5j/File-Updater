using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace File_Updater
{
    public partial class formText4 : Form
    {
        public formText4()
        {
            InitializeComponent();
            richTextBox1.ReadOnly = true;
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {

        }
        private void RichTextBox1_LinkClicked(object sender, LinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start("explorer.exe", e.LinkText);

        }

    }
}

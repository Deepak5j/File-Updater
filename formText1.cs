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
    public partial class formText1 : Form
    {
        public formText1()
        {
            InitializeComponent();
            richTextBox1.ReadOnly = true;
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {

            richTextBox1.SelectionAlignment = HorizontalAlignment.Center;

        }

        private void formText1_Load(object sender, EventArgs e)
        {

        }
    }
}

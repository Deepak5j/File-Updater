using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Runtime.Remoting.Contexts;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Spire.Xls;
using ExcelApp = Microsoft.Office.Interop.Excel;

namespace File_Updater
{
    public partial class Form1 : Form
    {
        Dictionary<string, string> myDick = new Dictionary<string, string>();
        string fileToUpdate;
        public Form1()
        {
            InitializeComponent();
            button2.Enabled = false;
            button3.Enabled = false;
            button4.Enabled = false;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DialogResult result = openFileDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
                string file = openFileDialog1.FileName; //FILE path
                try
                {
                    string[] lines = System.IO.File.ReadAllLines(file);
                    //MessageBox.Show(Convert.ToString(lines[0].Split('\t')[1].ToLower()));
                    if (!lines[0].Split('\t')[1].ToLower().Equals("dues left"))
                    {
                        throw new Exception();
                    }
                    foreach (string line in lines)
                    {
                        var cols = line.Split('\t')[1];
                        if (cols.Equals("Yes"))
                        {
                            try
                            {
                                myDick.Add(line.Split('\t')[0], line.Split('\t')[1]);
                            }
                            catch (Exception E)
                            {
                                //MessageBox.Show(E.Message);                            
                            }
                        }
                    }/*
                    string s = "";
                    foreach (KeyValuePair<string,string> i in myDick)
                    {
                        s += i;
                    }
                    MessageBox.Show(s);*/
                    button1.Enabled = false;
                    button2.Enabled = true;
                    MessageBox.Show("File Loaded Successfuly!");
                }
                catch (Exception E)
                {
                    button1.Enabled = true;
                    button2.Enabled = false;
                    button3.Enabled = false;
                    button4.Enabled = false;
                    MessageBox.Show("Wrong File Loaded!");
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DialogResult result2 = openFileDialog2.ShowDialog();
            if (result2 == DialogResult.OK)
            {
                string file2 = openFileDialog2.FileName;
                try
                {
                    string[] lines2 = System.IO.File.ReadAllLines(file2);
                    //MessageBox.Show(Convert.ToString(lines2[0].Split('\t')[1].ToLower()));
                    if (!lines2[0].Split('\t')[1].ToLower().Equals("amount applicable"))
                    {
                        throw new Exception();
                    }
                    foreach (string line in lines2)
                    {
                        var cols = line.Split('\t')[0];
                        try
                        {
                            List<string> keys = new List<string>(myDick.Keys);
                            foreach (string key in keys)
                            {
                                if (key == cols)
                                {
                                    myDick[cols] = line.Split('\t')[1];
                                }
                            }
                        }
                        catch (Exception E)
                        {
                            //MessageBox.Show("###"+E.Message);
                        }
                    }/*
                    string s = "";
                    foreach (KeyValuePair<string, string> i in myDick)
                    {
                        s += i;
                    }
                    MessageBox.Show(s);*/
                    button2.Enabled = false;
                    button3.Enabled = true;
                    MessageBox.Show("File Loaded Successfuly!");
                }
                catch (Exception)
                {
                    button1.Enabled = false;
                    button2.Enabled = true;
                    button3.Enabled = false;
                    button4.Enabled = false;
                    MessageBox.Show("Wrong File Loaded!");
                }
            }
        }

        [Obsolete]
        private void button3_Click(object sender, EventArgs e)
        {
            DialogResult result3 = openFileDialog3.ShowDialog();
            if (result3 == DialogResult.OK)
            {
                string file3 = openFileDialog3.FileName;
                fileToUpdate = file3;
                try
                {
                    string[] lines3 = System.IO.File.ReadAllLines(file3);
                    //MessageBox.Show(Convert.ToString(lines3[0].Split('\t')[2].ToLower()));
                    if (!lines3[0].Split('\t')[2].ToLower().Equals("dues"))
                    {
                        throw new Exception();
                    }
                    foreach (string line in lines3)
                    {
                        var cols = line.Split('\t')[0];
                        try
                        {
                            List<string> keys = new List<string>(myDick.Keys);
                            foreach (string key in keys)
                            {
                                if (key == cols)
                                {
                                    myDick[cols] = Convert.ToString(Convert.ToDouble(myDick[cols]) + Convert.ToDouble(line.Split('\t')[2]));
                                }
                            }
                        }
                        catch (Exception E)
                        {
                            //MessageBox.Show(E.Message);
                        }
                    }/*
                    string s = "";
                    foreach (KeyValuePair<string, string> i in myDick)
                    {
                        s += i;
                    }
                    MessageBox.Show(s);*/
                    button3.Enabled = false;
                    button4.Enabled = true;
                    MessageBox.Show("File Loaded Successfuly!");

                }
                catch (Exception)
                {
                    button1.Enabled = false;
                    button2.Enabled = false;
                    button3.Enabled = true;
                    button4.Enabled = false;
                    MessageBox.Show("Wrong File Loaded!");
                }
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            List<string> keys = new List<string>(myDick.Keys);
            string[] arrLine = File.ReadAllLines(fileToUpdate);
            try
            {
                foreach (string key in keys)
                {
                    foreach (string line in arrLine)
                    {
                        if (key == line.Split('\t')[0])
                        {

                            arrLine[Convert.ToInt32(key)] = line.Split('\t')[0] + '\t' + line.Split('\t')[1] + '\t' + myDick[key];
                        }
                    }
                }
                File.WriteAllLines(fileToUpdate, arrLine);
                button1.Enabled = true;
                button2.Enabled = false;
                button3.Enabled = false;
                button4.Enabled = false;
                MessageBox.Show("File Updated Successfuly!");
            }
            catch (Exception E)
            {
                button1.Enabled = true;
                button2.Enabled = false;
                button3.Enabled = false;
                button4.Enabled = false;
                MessageBox.Show(E.Message);
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void menuClose(object sender, EventArgs e)
        {
            System.Windows.Forms.Application.Exit();
        }
        private void menuViewHelpText1(object sender, EventArgs e)
        {
            formText1 formText1 = new formText1();
            formText1.Show();
            //formText1.ShowDialog() //parent window freezes
        }

        private void menuViewHelpText2(object sender, EventArgs e)
        {
            formText2 formText2 = new formText2();
            formText2.Show();
        }
        private void menuViewHelpText3(object sender, EventArgs e)
        {
            formText3 formText3 = new formText3();
            formText3.Show();
        }
        private void menuViewHelpAbout(object sender, EventArgs e)
        {
            formText4 formText4 = new formText4();
            formText4.Show();
        }



        /*
        protected override void OnPaint(PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics, ClientRectangle, System.Drawing.Color.Transparent, ButtonBorderStyle.Solid);
            
        }*/
    }
}

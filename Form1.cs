using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace Converter
{
    public partial class Form1 : MetroFramework.Forms.MetroForm
    {
        public Form1()
        {
           
            InitializeComponent();
            metroTextBox1.Text = ConfigurationManager.AppSettings["InputFolder"];
            metroTextBox2.Text = ConfigurationManager.AppSettings["OutputFolder"];
            metroTextBox3.Text = ConfigurationManager.AppSettings["RefreshRate"];

        }
        private string UD;
        private string AE;
        private string NK;
        private string AER;
        private string SK;
        public int i;
        private void timer1_Tick(object sender, EventArgs e)
        {

            UD = null;
            AE = null;
            NK = null;
            AER = null;
            SK = null;


            i = 0;
            try
            {
                
                foreach (var item in Directory.GetFiles(metroTextBox1.Text))
                {
                    i++;
                }
                
                if (i != 0)
                {
                    foreach (var item in Directory.GetFiles(metroTextBox1.Text))
                    {
                        string localtxt = Environment.CurrentDirectory + "\\" + item.Replace(".pdf", "").Replace(metroTextBox1.Text, "") + ".txt";
                        string strCmdText = "/C " + @"d:\pdftotext.exe """  + item + "\" \"" + localtxt + "\"";
                        //****************************************************************************************************************************
                        ProcessStartInfo procStartInfo = new ProcessStartInfo("cmd", strCmdText) { RedirectStandardError = true, RedirectStandardOutput = true, UseShellExecute = false, CreateNoWindow = true };

                        using (Process proc = new Process())
                        {
                            proc.StartInfo = procStartInfo;
                            proc.Start();
                            string output = proc.StandardOutput.ReadToEnd();
                            if (string.IsNullOrEmpty(output)) output = proc.StandardError.ReadToEnd();
                            this.Hide();
                        }

                        Thread.Sleep(100);
                        ////////////////////////////////////////////////////////////////////////////////////
                        ////////////////////////////////////////////////////////////////////////////////////
                        ////////////////////////////////////////////////////////////////////////////////////
                        ////////////////////////////////////////////////////////////////////////////////////
                        bool writing = false;
                        List<string> doc = new List<string>();

                        XmlDocument xmlDoc = new XmlDocument();
                        //using (StreamReader sr = new StreamReader(metroTextBox2.Text + "\\" + item.Replace(".pdf", "").Replace(metroTextBox1.Text, "") + ".txt"))
                        //{
                        //    while (sr.Peek() >= 0)
                        //    {
                        //        doc.Add(sr.ReadLine());
                        //    }
                        //}
                        foreach (var line in File.ReadAllLines(localtxt,Encoding.UTF7))
                        {
                            if (line.Contains("Utdanning"))
                            {
                                writing = true;
                            }
                            if (line.Contains("Arbeidserfaring"))
                            {
                                writing = false;
                            }
                            if (writing && !line.Contains("Utdanning"))
                            {
                                UD = UD + line + " ";
                            }
                        }
                        foreach (var line in File.ReadAllLines(localtxt, Encoding.UTF7))
                        {
                            if (line.Contains("Arbeidserfaring"))
                            {
                                writing = true;
                            }
                            if (line.Contains("Nøkkelkompetanse"))
                            {
                                writing = false;
                            }
                            if (writing && !line.Contains("Arbeidserfaring"))
                            {
                                AE = AE + line + " ";
                            }
                        }



                        foreach (var line in File.ReadAllLines(localtxt, Encoding.UTF7))
                        {
                            if (line.Contains("Nøkkelkompetanse"))
                            {
                                writing = true;
                            }
                            if (line.Contains("Annen erfaring"))
                            {
                                writing = false;
                            }
                            if (writing && !line.Contains("Nøkkelkompetanse"))
                            {
                                NK = NK +  line;
                            }
                        }
                        foreach (var line in File.ReadAllLines(localtxt, Encoding.UTF7))
                        {
                            if (line.Contains("Annen erfaring"))
                            {
                                writing = true;
                            }
                            if (line.Contains("Språk"))
                            {
                                writing = false;
                            }
                            if (writing && !line.Contains("Annen erfaring"))
                            {
                                AER = AER + line;
                            }
                        }
                        foreach (var line in File.ReadAllLines(localtxt, Encoding.UTF7))
                        {
                            if (line.Contains("Språk"))
                            {
                                writing = true;
                            }
                            if (line.Contains("Ønsker for neste jobb"))
                            {
                                writing = false;
                            }
                            if (writing && !line.Contains("Språk"))
                            {
                                SK = SK + line;
                            }

                        }

                        XmlWriterSettings settings = new XmlWriterSettings();
                        settings.Indent = true;
                        settings.IndentChars = ("    ");
                        settings.CloseOutput = true;
                        //settings.OmitXmlDeclaration = true;
                        using (XmlWriter writer = XmlWriter.Create(metroTextBox2.Text + "\\" + item.Replace(".pdf", "").Replace(metroTextBox1.Text, "") + ".xml", settings))
                        {
                            writer.WriteStartElement("data");
                            writer.WriteElementString("Utdanning", UD);
                            writer.WriteElementString("Arbeidserfaring", AE);
                            writer.WriteElementString("Nøkkelkompetanse", NK);
                            writer.WriteElementString("AnnenErfaring", AER);
                            writer.WriteElementString("Språk", SK.Replace("\r\n\fSpråk\r\n",""));
                            writer.WriteEndElement();
                            writer.Flush();
                        }

                       
                        ////////////////////////////////////////////////////////////////////////////////////
                        ////////////////////////////////////////////////////////////////////////////////////
                        ////////////////////////////////////////////////////////////////////////////////////
                        ////////////////////////////////////////////////////////////////////////////////////
                        File.Delete(item);
                    }
                }
            }
            catch (Exception d)
            {
                MessageBox.Show(d.Message);
                timer1.Enabled = false;
                metroButton1.Text = "Start";
                throw;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
       
        }

        private void metroButton1_Click(object sender, EventArgs e)
        {
            if (timer1.Enabled)
            {
                timer1.Enabled = false;
                metroButton1.Text = "Start";
            }
            if (!timer1.Enabled)
            {
                timer1.Enabled = true;
                metroButton1.Text = "Stop";
            }
        }

        private void metroTextBox1_TextChanged(object sender, EventArgs e)
        {
            try
            {
                timer1.Interval = int.Parse(metroTextBox3.Text);
            }
            catch (Exception)
            {
            }
            
        }
    }
}

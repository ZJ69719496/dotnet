using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        string dataNameFile = "";
        private void Button1_Click(object sender, EventArgs e)
        {
            this.openFileDialog1.FileName = "";
            label1.Text = "";

            this.openFileDialog1.Multiselect = false;
            this.openFileDialog1.Filter = "XML files(*.xml)|*.xml";
            if (this.openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                if (openFileDialog1.FileName != "")
                {
                    dataNameFile = openFileDialog1.FileName;
                    label1.Text = dataNameFile;
                    if (File.Exists(dataNameFile))
                    {
                        ParseDataName(dataNameFile);
                    }
                }
            }

        }

        /// <summary>
        /// 解析数据名称
        /// </summary>
        /// <param name="dataNameFile"></param>
        private void ParseDataName(string path)
        {
            string strContent = File.ReadAllText(path);
            strContent = Regex.Replace(strContent, "ss:", "");
            File.WriteAllText(path + "temp", strContent);

            XDocument doc = XDocument.Load(path + "temp");
            if (doc.Root.Name.LocalName == "ErcExpQDataTechConfigData")
            {
                Console.WriteLine("ErcExpQDataTechConfigData");
            }
            else
            {
                textBox1.Text = "";
                textBox1.AppendText("Parse error");
                Console.WriteLine("Parse error");
                return;
            }

            textBox1.Text = "";
            foreach (XElement item in doc.Root.Descendants("TechData"))
            {
                textBox1.AppendText("TagName：" + item.Element("TagName").Value + "        number：" + item.Element("TagName").Value.ToLower().TrimStart("tech".ToCharArray()) + "        ValueName：" + item.Element("ValueName").Value + Environment.NewLine);
            }

            File.Delete(path + "temp");

        }


        private void Button2_Click(object sender, EventArgs e)
        {
            this.openFileDialog1.FileName = "";
            label2.Text = "";

            this.openFileDialog1.Multiselect = false;
            this.openFileDialog1.Filter = "XML files(*.xml)|*.xml";
            if (this.openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                if (openFileDialog1.FileName != "")
                {
                    dataNameFile = openFileDialog1.FileName;
                    label2.Text = dataNameFile;
                    if (File.Exists(dataNameFile))
                    {
                        ParseData(dataNameFile);
                    }
                }
            }
        }

        /// <summary>
        /// 解析数据
        /// </summary>
        /// <param name="dataNameFile"></param>
        private void ParseData(string path)
        {
            string strContent = File.ReadAllText(path);
            strContent = Regex.Replace(strContent, "xmlns=\"rds\"", "");
            File.WriteAllText(path + "temp", strContent);

            strContent = Regex.Replace(strContent, "ss:", "");//防止加载成报错信息的XML

            XDocument doc2 = XDocument.Load(path + "temp");

            if (doc2.Root.Name.LocalName == "file")
            {
                Console.WriteLine("file");
            }
            else
            {
                textBox2.Text = "";
                textBox2.AppendText("Parse error");
                Console.WriteLine("Parse error");
                return;
            }

            textBox2.Text = "";
            foreach (XElement item in doc2.Root.Descendants("identification"))
            {
                var workstation = item.Element("workstation");
                textBox2.AppendText("workstation: " + workstation.Attribute("name").Value.ToString() + Environment.NewLine);

                var facility = workstation.Element("facility");
                textBox2.AppendText("facility: " + facility.Attribute("name").Value.ToString() + Environment.NewLine);

                var station = facility.Element("station");
                textBox2.AppendText("station: " + station.Attribute("name").Value.ToString() + Environment.NewLine);

                var controller = station.Element("controller");
                textBox2.AppendText("controller: " + controller.Attribute("name").Value.ToString() + Environment.NewLine);
                textBox2.AppendText("product: " + controller.Attribute("product").Value.ToString() + Environment.NewLine);
                textBox2.AppendText("ordinal_join_tech: " + controller.Attribute("ordinal_join_tech").Value.ToString() + Environment.NewLine);
                textBox2.AppendText("manufacturer: " + controller.Attribute("manufacturer").Value.ToString() + Environment.NewLine);
                textBox2.AppendText("IP_number: " + controller.Attribute("IP_number").Value.ToString() + Environment.NewLine);

                textBox2.AppendText(Environment.NewLine);
                textBox2.AppendText(Environment.NewLine);
            }

            int count = 0;
            foreach (XElement item in doc2.Root.Descendants("data"))
            {
                textBox2.AppendText("tech01：" + item.Attribute("tech01").Value.ToString());
                textBox2.AppendText(" tech02：" + item.Attribute("tech02").Value.ToString());
                textBox2.AppendText(" tech03：" + item.Attribute("tech03").Value.ToString());
                textBox2.AppendText(" tech04：" + item.Attribute("tech04").Value.ToString());
                textBox2.AppendText(" tech05：" + item.Attribute("tech05").Value.ToString());
                textBox2.AppendText(" tech30：" + item.Attribute("tech30").Value.ToString());


                textBox2.AppendText(" joining_spot：" + item.Attribute("joining_spot").Value.ToString());
                textBox2.AppendText(" joining_spot：" + item.Attribute("error_code").Value.ToString());
                textBox2.AppendText(" quality：" + item.Attribute("quality").Value.ToString());
                textBox2.AppendText(" time：" + item.Attribute("time").Value.ToString());
                textBox2.AppendText(" date：" + item.Attribute("date").Value.ToString());
                textBox2.AppendText(" serial：" + item.Attribute("serial").Value.ToString());
                textBox2.AppendText(Environment.NewLine);
                if (++count == 20)
                {
                    MessageBox.Show("DEMO 只显示20条数据");
                    break;
                }

            }
            File.Delete(path + "temp");

        }

        private void Button3_Click(object sender, EventArgs e)
        {
            this.openFileDialog1.FileName = "";
            label3.Text = "";

            this.openFileDialog1.Multiselect = false;
            this.openFileDialog1.Filter = "XML files(*.xml)|*.xml";
            if (this.openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                if (openFileDialog1.FileName != "")
                {
                    dataNameFile = openFileDialog1.FileName;
                    label3.Text = dataNameFile;
                    if (File.Exists(dataNameFile))
                    {
                        ParseFaultData(dataNameFile);
                    }
                }
            }
        }

        private void ParseFaultData(string path)
        {
            string strContent = File.ReadAllText(path);
            strContent = strContent.Replace("xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\"", "");
            strContent = Regex.Replace(strContent, "ss:", "");

            File.WriteAllText(path + "temp", strContent);
            XDocument doc3 = XDocument.Load(path + "temp");

            if (doc3.Root.Name.LocalName == "Workbook")
            {
                Console.WriteLine("Workbook");
            }
            else
            {
                textBox3.Text = "";
                textBox3.AppendText("Parse error");
                Console.WriteLine("Parse error");
                return;
            }

            textBox3.Text = "";

            int count = 0;
            foreach (XElement row in doc3.Root.Descendants("Row"))
            {
                foreach (XElement cell in row.Descendants("Cell"))
                {
                    textBox3.AppendText(cell.Element("Data").Value.ToString());
                    textBox3.AppendText("  ——  ");
                }
                textBox3.AppendText(Environment.NewLine);

                if (++count == 20)
                {
                    MessageBox.Show("DEMO 只显示20条数据");
                    break;
                }
            }
            File.Delete(path + "temp");

        }

    }
}

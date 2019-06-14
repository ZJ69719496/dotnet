using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;
using System.Xml;
using System.Xml.Linq;

namespace ExcelImport
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        XmlDocument xml;
        DataTable dt1;
        DataTable dt2;
        DataTable dt3;
        DataTable dt4;

        private void Button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.FileName = "";
            openFileDialog1.Filter = "Excel 2003|*.xls|Excel 2007|*.xlsx|All files(*.*)|*.*";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string ExcelPath = openFileDialog1.FileName;
                label1.Text = ExcelPath;
                ImportExcel(ExcelPath);

                string xmltemp = "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
                                 "<Settings Name=\"工位信号配置\" Description=\"工位信号配置\">" +
                                    "<TagClassTemplate Name=\"通用标签模板\" Description=\"\" DeviceType=\"Siemens\" ConnectTimeOut=\"1000\" PlcType=\"S7_300\" CreateTime=\"" + DateTime.Now.ToString("hh:mm:ss") + "\">" +
                                    "</TagClassTemplate>" +
                                 "</Settings>";

                xml = new XmlDocument();
                xml.LoadXml(xmltemp);

                if (dt1 != null)
                {
                    dataGridView1.DataSource = dt1;
                    ParseToTagNodeTemplate(dt1);
                }

                if (dt2 != null)
                {
                    dataGridView2.DataSource = dt2;
                    ParseToWorkStation(dt2);
                }

                if (dt3 != null)
                {
                    dataGridView3.DataSource = dt3;
                    ParseToQuality(dt3);
                }

                if (dt1 != null)
                {
                    dataGridView4.DataSource = dt4;
                    ParseToAlarm(dt4);
                }


            }
        }

        #region TagNodeTemplate 增加

        /// <summary>
        /// 解析到TagNodeTemplate
        /// </summary>
        /// <param name="dt"></param>
        private void ParseToTagNodeTemplate(DataTable dt)
        {
            foreach (DataRow dr in dt.Rows)
            {
                addTagNodeTemplate(dr, xml);
            }
        }

        /// <summary>
        /// 增加TagNodeTemplate
        /// </summary>
        /// <param name="dr"></param>
        /// <param name="xml"></param>
        private void addTagNodeTemplate(DataRow dr, XmlDocument xml)
        {
            XmlElement TagNodeTemplate = xml.CreateElement("TagNodeTemplate");
            TagNodeTemplate.SetAttribute("TagName", dr[1].ToString());
            TagNodeTemplate.SetAttribute("Description", dr[7].ToString());
            TagNodeTemplate.SetAttribute("TagType", dr[3].ToString());
            TagNodeTemplate.SetAttribute("Length", dr[2].ToString());
            TagNodeTemplate.SetAttribute("DB", dr[6].ToString().TrimStart("DB".ToCharArray())); ;
            TagNodeTemplate.SetAttribute("StartAddress", dr[4].ToString() + dr[5].ToString());
            TagNodeTemplate.SetAttribute("IsMonitor", dr[9].ToString());
            TagNodeTemplate.SetAttribute("IsForweb", dr[10].ToString());
            TagNodeTemplate.SetAttribute("IsEnable", dr[8].ToString());
            TagNodeTemplate.SetAttribute("DataProperty", "0");
            TagNodeTemplate.SetAttribute("GroupID", "0");
            TagNodeTemplate.SetAttribute("GroupOrderBy", "0");
            TagNodeTemplate.SetAttribute("BelongEngineType", "0");
            TagNodeTemplate.SetAttribute("ListNo", dr[0].ToString());
            TagNodeTemplate.SetAttribute("Address", "自动生成，生成规则待定");

            XmlNode Settings = xml.SelectSingleNode("Settings");
            Settings.SelectSingleNode("TagClassTemplate").AppendChild(TagNodeTemplate);
        }

        #endregion

        #region WorkStation 增加

        /// <summary>
        /// 解析到WorkStation
        /// </summary>
        /// <param name="dt"></param>
        private void ParseToWorkStation(DataTable dt)
        {
            foreach (DataRow dr in dt.Rows)
            {
                addWorkStation(dr, xml);
            }
        }

        /// <summary>
        /// 增加Workstation
        /// </summary>
        /// <param name="dr"></param>
        private void addWorkStation(DataRow dr, XmlDocument xml)
        {
            XmlElement WorkStation = xml.CreateElement("WorkStation");

            WorkStation.SetAttribute("Name", dr[1].ToString());
            WorkStation.SetAttribute("OpCode", dr[0].ToString());
            WorkStation.SetAttribute("Description", dr[4].ToString());
            WorkStation.SetAttribute("DeviceType", dr[8].ToString());
            WorkStation.SetAttribute("ConnectTimeOut", dr[9].ToString());
            WorkStation.SetAttribute("IpAddress", dr[11].ToString());
            WorkStation.SetAttribute("Port", dr[12].ToString());
            WorkStation.SetAttribute("DBOffset", dr[13].ToString());
            WorkStation.SetAttribute("AddressOffset", dr[14].ToString());
            WorkStation.SetAttribute("PlcType", dr[10].ToString());
            WorkStation.SetAttribute("function", dr[26].ToString());
            WorkStation.SetAttribute("MonitorPropertyCode", dr[31].ToString());
            WorkStation.SetAttribute("CreateTime", DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss"));

            WorkStation.AppendChild(CreateCommomTagClass(xml, dr[0].ToString()));

            var QualityTagClass = CreateQualityTagClass(xml);
            WorkStation.AppendChild(QualityTagClass);
            QualityTagClasses.Add(dr[1].ToString(), new pack { opcode = dr[0].ToString(), element = QualityTagClass });

            var AlarmTagClass = CreateAlarmTagClass(xml);
            WorkStation.AppendChild(AlarmTagClass);
            AlarmTagClasses.Add(dr[1].ToString(), new pack { opcode = dr[0].ToString(), element = AlarmTagClass });

            xml.SelectSingleNode("Settings").AppendChild(WorkStation);

        }
        struct pack
        {
            public string opcode;
            public XmlElement element;
        }

        Dictionary<string, pack> QualityTagClasses = new Dictionary<string, pack>();
        Dictionary<string, pack> AlarmTagClasses = new Dictionary<string, pack>();

        /// <summary>
        /// 创建通用标签组
        /// </summary>
        /// <returns></returns>
        private XmlElement CreateCommomTagClass(XmlDocument xml, string OpCode)
        {
            XmlElement tagClass = xml.CreateElement("TagClass");
            tagClass.SetAttribute("Name", "通用标签组");
            tagClass.SetAttribute("TagClassType", "0");
            tagClass.SetAttribute("Description", "通用标签组");

            int index = 1;
            foreach (DataRow dr in dt1.Rows)
            {
                tagClass.AppendChild(addCommomTagNode(dr, OpCode, string.Format("{0:d3}", index)));
                index++;
            }

            return tagClass;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="dr"></param>
        private XmlElement addCommomTagNode(DataRow dr, string OpCode, string index)
        {
            XmlElement TagNodeTemplate = xml.CreateElement("TagNode");
            TagNodeTemplate.SetAttribute("TagName", dr[1].ToString());
            TagNodeTemplate.SetAttribute("Description", dr[7].ToString());
            TagNodeTemplate.SetAttribute("TagType", dr[3].ToString());
            TagNodeTemplate.SetAttribute("Length", dr[2].ToString());
            TagNodeTemplate.SetAttribute("DB", dr[6].ToString().TrimStart("DB".ToCharArray()));
            TagNodeTemplate.SetAttribute("StartAddress", dr[4].ToString() + dr[5].ToString());
            TagNodeTemplate.SetAttribute("IsMonitor", dr[9].ToString());
            TagNodeTemplate.SetAttribute("IsForweb", dr[10].ToString());
            TagNodeTemplate.SetAttribute("IsEnable", dr[8].ToString());
            TagNodeTemplate.SetAttribute("DataProperty", "0");
            TagNodeTemplate.SetAttribute("GroupID", "0");
            TagNodeTemplate.SetAttribute("GroupOrderBy", "0");
            TagNodeTemplate.SetAttribute("BelongEngineType", "0");
            TagNodeTemplate.SetAttribute("ListNo", dr[0].ToString());
            TagNodeTemplate.SetAttribute("Address", dr[6].ToString() + "." + dr[4].ToString() + "@" + dr[3].ToString());
            TagNodeTemplate.SetAttribute("TagID", "7" + OpCode + index);

            return TagNodeTemplate;
        }

        /// <summary>
        /// 创建质量标签组
        /// </summary>
        /// <returns></returns>
        private XmlElement CreateQualityTagClass(XmlDocument xml)
        {
            XmlElement tagClass = xml.CreateElement("TagClass");
            tagClass.SetAttribute("Name", "质量标签组");
            tagClass.SetAttribute("TagClassType", "1");
            tagClass.SetAttribute("Description", "质量标签组");
            return tagClass;
        }

        /// <summary>
        /// 创建个性标签组
        /// </summary>
        /// <returns></returns>
        private XmlElement CreateAlarmTagClass(XmlDocument xml)
        {
            XmlElement tagClass = xml.CreateElement("TagClass");
            tagClass.SetAttribute("Name", "个性标签组");
            tagClass.SetAttribute("TagClassType", "2");
            tagClass.SetAttribute("Description", "个性标签组");
            return tagClass;
        }

        #endregion

        #region 质量标签组 增加

        /// <summary>
        /// 解析到质量标签组
        /// </summary>
        /// <param name="dt3"></param>
        private void ParseToQuality(DataTable dt)
        {
            int index = 1;
            foreach (DataRow dr in dt.Rows)
            {
                addQualityTagNode(dr, xml, string.Format("{0:d3}", index));
                index++;
            }
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="dr"></param>
        /// <param name="xml"></param>
        private void addQualityTagNode(DataRow dr, XmlDocument xml, string index)
        {
            string OPNAME = dr[2].ToString();

            var tagclass = QualityTagClasses[OPNAME];//更加工位查找对应的 tagclass 质量标签组

            if (tagclass.element != null)
            {
                var tagnode = xml.CreateElement("TagNode");
                tagnode.SetAttribute("TagID", "8" + tagclass.opcode + index);
                tagnode.SetAttribute("TagName", dr[1].ToString());
                tagnode.SetAttribute("Description", dr[20].ToString());
                tagnode.SetAttribute("TagType", dr[6].ToString());
                tagnode.SetAttribute("Length", dr[4].ToString());
                tagnode.SetAttribute("DB", dr[5].ToString().TrimStart("DB".ToCharArray()));
                tagnode.SetAttribute("StartAddress", dr[7].ToString());
                tagnode.SetAttribute("Address", dr[5].ToString() + "." + dr[7].ToString() + "@" + dr[6].ToString());
                tagnode.SetAttribute("IsMonitor", dr[17].ToString());
                tagnode.SetAttribute("IsForweb", dr[18].ToString());
                tagnode.SetAttribute("IsEnable", dr[16].ToString());
                tagnode.SetAttribute("DataProperty", dr[11].ToString());
                tagnode.SetAttribute("GroupID", dr[9].ToString());
                tagnode.SetAttribute("GroupOrderBy", dr[10].ToString());
                tagnode.SetAttribute("BelongEngineType", dr[14].ToString());
                tagnode.SetAttribute("ListNo", dr[0].ToString());

                tagclass.element.AppendChild(tagnode);
            }

        }

        #endregion

        #region 个性标签组 增加

        /// <summary>
        /// 解析到个性标签组
        /// </summary>
        /// <param name="dt4"></param>
        private void ParseToAlarm(DataTable dt)
        {
            int index = 1;
            foreach (DataRow dr in dt.Rows)
            {
                addAlarmTagNode(dr, xml, string.Format("{0:d3}", index));
                index++;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="dr"></param>
        /// <param name="xml"></param>
        private void addAlarmTagNode(DataRow dr, XmlDocument xml, string index)
        {
            string OPNAME = dr[1].ToString();

            var tagclass = AlarmTagClasses[OPNAME];//更加工位查找对应的 tagclass 质量标签组

            if (tagclass.element != null)
            {
                var tagnode = xml.CreateElement("TagNode");
                tagnode.SetAttribute("TagID", "9" + tagclass.opcode + index);
                tagnode.SetAttribute("TagName", dr[0].ToString());
                tagnode.SetAttribute("Description", dr[14].ToString());
                tagnode.SetAttribute("TagType", dr[5].ToString());
                tagnode.SetAttribute("Length", dr[3].ToString());
                tagnode.SetAttribute("DB", dr[4].ToString().TrimStart("DB".ToCharArray()));
                tagnode.SetAttribute("StartAddress", dr[6].ToString());
                tagnode.SetAttribute("Address", dr[4].ToString() + "." + dr[6].ToString() + "@" + dr[5].ToString());
                tagnode.SetAttribute("IsMonitor", dr[16].ToString());
                tagnode.SetAttribute("IsForweb", dr[17].ToString());
                tagnode.SetAttribute("IsEnable", dr[15].ToString());
                tagnode.SetAttribute("DataProperty", dr[10].ToString());
                tagnode.SetAttribute("GroupID", dr[8].ToString());
                tagnode.SetAttribute("GroupOrderBy", dr[9].ToString());
                tagnode.SetAttribute("BelongEngineType", dr[13].ToString());
                tagnode.SetAttribute("ListNo", "0");

                tagclass.element.AppendChild(tagnode);
            }
        }

        #endregion

        #region Excel操作

        /// <summary>
        /// 获取excel内容
        /// </summary>
        /// <param name="filePath">excel文件路径</param>
        /// <returns></returns>
        public void ImportExcel(string filePath)
        {
            try
            {
                using (FileStream fsRead = File.OpenRead(filePath))
                {
                    IWorkbook wk = null;
                    //获取后缀名
                    string extension = filePath.Substring(filePath.LastIndexOf(".")).ToString().ToLower();
                    //判断是否是excel文件
                    if (extension == ".xlsx" || extension == ".xls")
                    {
                        //判断excel的版本
                        if (extension == ".xlsx")
                        {
                            wk = new XSSFWorkbook(fsRead);
                        }
                        else
                        {
                            wk = new HSSFWorkbook(fsRead);
                        }

                        getSheetByIndex(out dt1, wk, 0);
                        getSheetByIndex(out dt2, wk, 1);
                        getSheetByIndex(out dt3, wk, 2);
                        getSheetByIndex(out dt4, wk, 3);
                    }

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Excel文件已被打开，请先关闭");
            }
        }

        /// <summary>
        /// 按照索引读取数据表
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="wk"></param>
        /// <param name="index"></param>
        private void getSheetByIndex(out DataTable dt, IWorkbook wk, int index)
        {
            dt = new DataTable();
            //获取第一个sheet
            ISheet sheet = wk.GetSheetAt(index);
            //获取第一行
            IRow headrow = sheet.GetRow(0);

            //创建列
            for (int i = headrow.FirstCellNum; i < headrow.Cells.Count; i++)
            {
                DataColumn datacolum = new DataColumn(headrow.GetCell(i).StringCellValue);
                dt.Columns.Add(datacolum);
            }

            //读取每行,从第二行起
            for (int r = 1; r <= sheet.LastRowNum; r++)
            {
                bool result = false;
                DataRow dr = dt.NewRow();
                //获取当前行
                IRow row = sheet.GetRow(r);
                //读取每列
                for (int j = 0; j < row.Cells.Count; j++)
                {
                    ICell cell = row.GetCell(j); //一个单元格
                    dr[j] = GetCellValue(cell); //获取单元格的值
                                                //全为空则不取
                    if (dr[j].ToString() != "")
                    {
                        result = true;
                    }
                }

                if (result == true)
                {
                    dt.Rows.Add(dr); //把每行追加到DataTable
                }
            }
        }

        /// <summary>
        /// 对单元格进行判断取值
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        private static string GetCellValue(ICell cell)
        {
            if (cell == null)
                return string.Empty;
            switch (cell.CellType)
            {
                case CellType.Blank: //空数据类型 这里类型注意一下，不同版本NPOI大小写可能不一样,有的版本是Blank（首字母大写)
                    return string.Empty;
                case CellType.Boolean: //bool类型
                    return cell.BooleanCellValue.ToString();
                case CellType.Error:
                    return cell.ErrorCellValue.ToString();
                case CellType.Numeric: //数字类型
                    if (HSSFDateUtil.IsCellDateFormatted(cell))//日期类型
                    {
                        return cell.DateCellValue.ToString();
                    }
                    else //其它数字
                    {
                        return cell.NumericCellValue.ToString();
                    }
                case CellType.Unknown: //无法识别类型
                default: //默认类型
                    return cell.ToString();//
                case CellType.String: //string 类型
                    return cell.StringCellValue;
                case CellType.Formula: //带公式类型
                    try
                    {
                        HSSFFormulaEvaluator e = new HSSFFormulaEvaluator(cell.Sheet.Workbook);
                        e.EvaluateInCell(cell);
                        return cell.ToString();
                    }
                    catch
                    {
                        return cell.NumericCellValue.ToString();
                    }
            }
        }

        #endregion

        private void Button2_Click(object sender, EventArgs e)
        {
            saveFileDialog1.Filter = "xml|*.xml";
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string path = saveFileDialog1.FileName;
                xml.Save(path);
            }
        }
    }
}

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
using System.IO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Data.SqlClient;
using System.Configuration;

namespace CraftImport
{


    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 工作簿
        /// </summary>
        IWorkbook workbook;
        /// <summary>
        /// 工位表对象集合
        /// </summary>
        Dictionary<string, WorkStationSheet> workStationSheets;

        /// <summary>
        /// 工位数量
        /// </summary>
        int sheetCount = 0;
        /// <summary>
        /// 每页46行
        /// </summary>
        int pageHeight = 46;

        /// <summary>
        /// 取零件 行开始位置
        /// </summary>
        int craftStart = 8;
        /// <summary>
        /// 取零件 行数量
        /// </summary>
        int craftLength = 7;

        /// <summary>
        /// 使用工具 行开始位置
        /// </summary>
        int toolStart = 17;
        /// <summary>
        /// 使用工具 行数量
        /// </summary>
        int toolLength = 5;

        /// <summary>
        /// 作业内容 行开始位置
        /// </summary>
        int workStart = 24;
        /// <summary>
        /// 作业内容 行数量
        /// </summary>
        int workLength = 19;

        //加载Excel按钮
        private void LoadBtn_Click(object sender, EventArgs e)
        {
            workStationSheets = new Dictionary<string, WorkStationSheet>();

            openFileDialog1.FileName = "";
            openFileDialog1.Filter = "Excel2007|*.xlsx|Excel2003|*.xls|ALL File|*.*";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                String filename = openFileDialog1.FileName;

                LoadExcel(filename);
            }
        }

        //加载Excel
        private void LoadExcel(string filename)
        {
            LoadBtn.Enabled = false;
            using (FileStream fs = new FileStream(filename, FileMode.Open, FileAccess.Read))
            {
                // 2007版本  
                if (filename.IndexOf(".xlsx") > 0)
                    workbook = new XSSFWorkbook(fs);
                // 2003版本  
                else if (filename.IndexOf(".xls") > 0)
                    workbook = new HSSFWorkbook(fs);

                if (workbook != null)
                {
                    sheetCount = workbook.NumberOfSheets;
                    ParseXLS2Object();
                }
            }
        }

        //解析Excel到对象
        private void ParseXLS2Object()
        {
            //异步解析Excel
            Task.Factory.StartNew(() =>
            {
                if (comboBox1.InvokeRequired)
                {
                    comboBox1.Invoke(new Action(() =>
                    {
                        comboBox1.Items.Clear();
                    }));
                }

                for (int i = 0; i < sheetCount; i++)
                {
                    ISheet sheet = workbook.GetSheetAt(i);

                    WorkStationSheet WorkStation = CreateWorkStation(sheet);
                    workStationSheets.Add(sheet.SheetName, WorkStation);

                    if (comboBox1.InvokeRequired)
                    {
                        comboBox1.Invoke(new Action(() =>
                        {
                            comboBox1.Items.Add(sheet.SheetName);
                        }));
                    }
                }

                if (LoadBtn.InvokeRequired)
                {
                    comboBox1.Invoke(new Action(() =>
                    {
                        comboBox1.SelectedIndex = 0;
                    }));
                }

                if (LoadBtn.InvokeRequired)
                {
                    LoadBtn.Invoke(new Action(() =>
                    {
                        LoadBtn.Enabled = true;
                        MessageBox.Show("导入成功！");
                    }));
                }
            });
        }

        //创建工位对象
        private WorkStationSheet CreateWorkStation(ISheet sheet)
        {
            WorkStationSheet workStation = new WorkStationSheet();
            workStation.PageCount = sheet.LastRowNum / pageHeight;
            workStation.Opname = sheet.SheetName;
            for (int i = 0; i < workStation.PageCount; i++)
            {
                workStation.Pages.Add(CreatePage(sheet, i));
            }
            return workStation;
        }

        //创建页面对象
        private Page CreatePage(ISheet sheet, int pageIndex)
        {
            Page page = new Page();
            page.Operations = CreateOperations(sheet, pageIndex);
            page.Tools = CreateTools(sheet, pageIndex);
            page.WorkConents = CreateWorkContents(sheet, pageIndex);
            return page;
        }

        //创建取零件操作集合
        private List<Operation> CreateOperations(ISheet sheet, int pageIndex)
        {
            List<Operation> operations = new List<Operation>();
            Operation op = null;
            int startindex = pageIndex * pageHeight + craftStart;
            for (int i = 0; i < craftLength; i++)
            {
                op = CreateOperation(sheet, startindex, pageIndex + 1);
                if (op == null) break;

                operations.Add(op);
                startindex++;
            }
            return operations;
        }

        //创建取零件操作
        private Operation CreateOperation(ISheet sheet, int startindex, int page)
        {
            Operation operation = new Operation();
            var row = sheet.GetRow(startindex);

            var cellO = row.GetCell(14);
            operation.id = cellO.ToString();

            var cellP = row.GetCell(15);
            operation.itemCode = cellP.StringCellValue;

            var cellU = row.GetCell(20);
            operation.itemName = cellU.StringCellValue;

            var cellAB = row.GetCell(27);
            operation.count = cellAB.ToString();

            var cellAC = row.GetCell(28);
            operation.supplier = cellAC.StringCellValue;

            var cellAD = row.GetCell(29);
            operation.standard = cellAD.StringCellValue;

            var cellAF = row.GetCell(31);
            operation.option = cellAF.StringCellValue;

            var cellAH = row.GetCell(33);
            operation.standard2 = cellAH.StringCellValue;

            var cellAI = row.GetCell(34);
            operation.option2 = cellAI.StringCellValue;
            operation.page = page;

            if (operation.id == "" && operation.itemCode == "" && operation.itemName == "" && operation.count == "" &&
                operation.supplier == "" && operation.standard == "" && operation.standard2 == "" && operation.option2 == "")
            {
                return null;
            }
            else
            {
                return operation;
            }
        }

        //创建使用工具集合
        private List<Tool> CreateTools(ISheet sheet, int pageIndex)
        {
            List<Tool> tools = new List<Tool>();
            Tool tool = null;
            int startindex = pageIndex * pageHeight + toolStart;
            for (int i = 0; i < toolLength; i++)
            {
                tool = CreateTool(sheet, startindex, pageIndex + 1);
                if (tool == null) break;

                tools.Add(tool);
                startindex++;
            }
            return tools;
        }

        //创建使用工具
        private Tool CreateTool(ISheet sheet, int startindex, int page)
        {
            Tool tool = new Tool();
            var row = sheet.GetRow(startindex);

            var cellO = row.GetCell(14);
            tool.id = cellO.ToString();

            var cellP = row.GetCell(15);
            tool.toolName = cellP.StringCellValue;

            var cellAA = row.GetCell(26);
            tool.toolType = cellAA.StringCellValue;

            var cellAF = row.GetCell(30);
            tool.defaultValue = cellAF.StringCellValue;
            tool.page = page;
            if (tool.id == "" && tool.toolName == "" && tool.toolType == "" && tool.defaultValue == "")
            {
                return null;
            }
            else
            {
                return tool;
            }
        }

        //创建作业内容集合
        private List<WorkContent> CreateWorkContents(ISheet sheet, int pageIndex)
        {
            List<WorkContent> workContents = new List<WorkContent>();
            WorkContent wc = null;
            int startindex = pageIndex * pageHeight + workStart;
            for (int i = 0; i < workLength; i++)
            {
                wc = CreateWorkContent(sheet, startindex, pageIndex + 1);
                if (wc == null) break;

                workContents.Add(wc);
                startindex++;
            }
            return workContents;
        }

        //创建作业内容
        private WorkContent CreateWorkContent(ISheet sheet, int startindex, int page)
        {
            WorkContent workContent = new WorkContent();
            var row = sheet.GetRow(startindex);

            var cellO = row.GetCell(14);
            workContent.workStep = cellO.StringCellValue;

            var cellX = row.GetCell(23);
            workContent.operationNotice = cellX.StringCellValue;

            var cellAG = row.GetCell(32);
            workContent.techStandard = cellAG.StringCellValue;
            workContent.page = page;
            if (workContent.workStep == "" && workContent.operationNotice == "" && workContent.techStandard == "")
            {
                return null;
            }
            else
            {
                return workContent;
            }

        }

        //生成SQL语句按钮
        private void BtnSql_Click(object sender, EventArgs e)
        {
            //TODO 读取工位  生成SQL
            if (comboBox1.Items.Count <= 0)
            {
                MessageBox.Show("请先加载格式正确的Excel工艺卡文件!", "错误");
                return;
            }
            string opname = comboBox1.SelectedItem.ToString();
            WorkStationSheet wss;
            string sql = "";

            if (workStationSheets.ContainsKey(opname))
            {
                wss = workStationSheets[opname];
                if (wss == null)
                {
                    MessageBox.Show("工位" + opname + "内容为空!", "错误");

                    return;
                }
                else
                {
                    sql = CreateSQL(wss);
                }
            }

            Form2 sqlform = new Form2();
            sqlform.textBox1.Text = sql;
            sqlform.ShowDialog();
        }

        //创建sql语句
        private string CreateSQL(WorkStationSheet wss)
        {
            StringBuilder sql = new StringBuilder();
            StringBuilder operationsql = new StringBuilder();
            StringBuilder toolsql = new StringBuilder();
            StringBuilder workContentsql = new StringBuilder();

            int step = 1;//工序内容 工步号
            string typeCode = "0";//发动机机型标志  发动机型号查询

            string machineType = wss.Pages[0].Operations[0].itemCode.Split('-')[0];
            typeCode = GetTypeCode(machineType);

            foreach (var page in wss.Pages)
            {
                if (chkOperation.Checked)
                    CreateOperationSQL(wss, operationsql, typeCode, page);

                if (chkTool.Checked)
                    CreateToolSQL(wss, toolsql, typeCode, page);

                if (chkWorkContent.Checked)
                    CreateWorkContentSQL(wss, workContentsql, step++, typeCode, page);
            }

            if (chkOperation.Checked)
            {
                sql.AppendFormat("DELETE FROM 电子看板_装配零件 WHERE 工位号 = '{0}'{1}", wss.Opname, Environment.NewLine);
                sql.Append(operationsql);
            }

            if (chkTool.Checked)
            {
                sql.Append(Environment.NewLine);
                sql.Append(Environment.NewLine);
                sql.AppendFormat("DELETE FROM 电子看板_使用工具 WHERE 工位号 = '{0}'{1}", wss.Opname, Environment.NewLine);
                sql.Append(toolsql);
            }

            if (chkWorkContent.Checked)
            {
                sql.Append(Environment.NewLine);
                sql.Append(Environment.NewLine);
                sql.AppendFormat("DELETE FROM 电子看板_工序内容 WHERE 工位号 = '{0}'{1}", wss.Opname, Environment.NewLine);
                sql.Append(workContentsql);
            }

            return sql.ToString();
        }

        //查询 发动机机型标志
        private string GetTypeCode(string machineType)
        {
            string result = "";
            string getMachineFlag = "SELECT 发动机机型标志 FROM MES_基本变量_机型代码 WHERE 发动机型号 = @machineType";
            string sqlconstr = ConfigurationManager.ConnectionStrings["strCon"].ConnectionString;
            using (SqlConnection conn = new SqlConnection(sqlconstr))
            {
                using (SqlCommand cmd = conn.CreateCommand())
                {
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = getMachineFlag;   //sql语句
                    cmd.Parameters.Add("@machineType", SqlDbType.NVarChar);
                    cmd.Parameters["@machineType"].Value = machineType;
                    result = cmd.ExecuteScalar().ToString();
                }
                //给参数sql语句的参数赋值
            }
            return result;
        }

        //创建 工序内容sql
        private static void CreateWorkContentSQL(WorkStationSheet wss, StringBuilder workContentsql, int step, string typeCode, Page page)
        {
            //生成 电子看板_装配零件 插入语句
            foreach (var workContent in page.WorkConents)
            {
                workContentsql.AppendFormat(@"INSERT INTO [dbo].[电子看板_工序内容] ([工位号] ,[产品型号代码] ,[工步号] ,[作业步骤] ,[操作内容] ,[技术标准] ,[工步属性] ,[页码])
                VALUES ('{0}', {1}, {2}, '{3}', '{4}', '{5}', {6}, {7}){8}",
                wss.Opname,
                typeCode,
                step++,
                workContent.workStep,
                workContent.operationNotice,
                workContent.techStandard,
                512,
                workContent.page,
                Environment.NewLine);
            }

        }

        //创建 使用工具sql
        private static void CreateToolSQL(WorkStationSheet wss, StringBuilder toolsql, string typeCode, Page page)
        {
            //生成 电子看板_使用工具 插入语句
            foreach (var tool in page.Tools)
            {
                toolsql.AppendFormat(@"INSERT INTO [dbo].[电子看板_使用工具] ([工位名称] ,[产品型号代码] ,[序号] ,[工具名称] ,[工具类型] ,[预设值] ,[页码])
                VALUES ('{0}', {1}, {2}, '{3}', '{4}', '{5}', {6}){7}",
                wss.Opname,
                typeCode,
                tool.id,
                tool.toolName,
                tool.toolType,
                tool.defaultValue,
                tool.page,
                Environment.NewLine);
            }
        }

        //创建 装配零件sql
        private static void CreateOperationSQL(WorkStationSheet wss, StringBuilder operationsql, string typeCode, Page page)
        {
            //生成 电子看板_装配零件 插入语句
            foreach (var operation in page.Operations)
            {
                string[] arr = operation.itemCode.Split('-');
                operationsql.AppendFormat(@"INSERT INTO [dbo].[电子看板_装配零件] ([工位号] ,[产品型号代码] ,[标号] ,[零件名称] ,[物料号] ,[零件数量] ,[生产厂家] ,[发动机型号] ,[零件图号] ,[页码])
                VALUES ('{0}', {1}, '{2}', '{3}', '{4}', {5}, '{6}', '{7}', '{8}', {9}){10}",
                wss.Opname,
                typeCode,
                operation.id,
                operation.itemName,
                arr[1],
                operation.count,
                operation.supplier,
                arr[0],
                arr[1],
                operation.page,
                Environment.NewLine);
            }
        }
    }

    #region 封装的类

    /// <summary>
    /// 工位表
    /// </summary>
    public class WorkStationSheet
    {
        public string Opname;
        /// <summary>
        /// 页面数量
        /// </summary>
        public int PageCount;
        /// <summary>
        /// 页面集合
        /// </summary>
        public List<Page> Pages = new List<Page>();
    }

    /// <summary>
    /// 工位页面
    /// </summary>
    public class Page
    {
        /// <summary>
        /// 取零件操作
        /// </summary>
        public List<Operation> Operations = new List<Operation>();
        /// <summary>
        /// 使用工具
        /// </summary>
        public List<Tool> Tools;
        /// <summary>
        /// 作业内容
        /// </summary>
        public List<WorkContent> WorkConents = new List<WorkContent>();
    }

    /// <summary>
    /// 操作
    /// </summary>
    public class Operation
    {
        /// <summary>
        /// 序号
        /// </summary>
        public string id;
        /// <summary>
        /// 料号
        /// </summary>
        public string itemCode;
        /// <summary>
        /// 零件名称
        /// </summary>
        public string itemName;
        /// <summary>
        /// 数量
        /// </summary>
        public string count;
        /// <summary>
        /// 供应商
        /// </summary>
        public string supplier;
        /// <summary>
        /// 标准配置
        /// </summary>
        public string standard;
        /// <summary>
        /// 可选配置
        /// </summary>
        public string option;
        /// <summary>
        /// 标准配置2
        /// </summary>
        public string standard2;
        /// <summary>
        /// 可选配置2
        /// </summary>
        public string option2;
        /// <summary>
        /// 当前页码
        /// </summary>
        public int page;
    }

    /// <summary>
    /// 工具
    /// </summary>
    public class Tool
    {
        /// <summary>
        /// 序号
        /// </summary>
        public string id;
        /// <summary>
        /// 工具名称
        /// </summary>
        public string toolName;
        /// <summary>
        /// 工具类型(特性)
        /// </summary>
        public string toolType;
        /// <summary>
        /// 预设值
        /// </summary>
        public string defaultValue;
        /// <summary>
        /// 当前页码
        /// </summary>
        public int page;
    }

    /// <summary>
    /// 作业内容
    /// </summary>
    public class WorkContent
    {
        /// <summary>
        /// 作业步骤
        /// </summary>
        public string workStep;
        /// <summary>
        /// 操作要点
        /// </summary>
        public string operationNotice;
        /// <summary>
        /// 技术标准
        /// </summary>
        public string techStandard;
        /// <summary>
        /// 当前页码
        /// </summary>
        public int page;
    }

    #endregion

}
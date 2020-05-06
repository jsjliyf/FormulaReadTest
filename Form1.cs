using NPOI.HPSF;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace FormulaReadTest
{
    public partial class Form1 : Form
    {
        public static string serverIP;
        public static bool isAudit;

        public static string[] array_Bank =  { "合计",//0
        "安平惠民村镇银行","深州丰源村镇银行","阜城家银村镇银行","故城家银村镇银行","武强家银村镇银行",//1-5
            "饶阳民商村镇银行","武邑邢农商村镇银行","枣强丰源村镇银行","冀州丰源村镇银行","景州丰源村镇银行",//6-10
            "衡水农商行","阜城农商行","武强农商行","景州农商行",//11-14
            "冀州农商行","枣强农商行","安平农商行","深州农商行",//15-18
            "饶阳联社","故城联社","武邑联社" };//19-21

        public static string userInfo;
        public static string auditDate;

        private DataSet dsForImportingChecking;
        private HSSFWorkbook wbForImportingChecking;
        private List<DataGridView> list_Dgv;  //待显示的DataGridView
        private BindingSource bs_Dgv;  //绑定到待显示的DataGridView上的数据源

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            serverIP = comboBox_ServerIP.SelectedItem.ToString();
        }

        private void DisplaySheet(List<DataGridView> list_dgv)
        {
            //先清空tabControl1上已有的内容
            tabControl1.TabPages.Clear();

            foreach (DataGridView dgv in list_dgv)
            {
                if (dgv != null)
                {
                    TabPage tp = new TabPage(dgv.Name);
                    tp.Controls.Add(dgv);
                    dgv.Dock = DockStyle.Fill;

                    tabControl1.TabPages.Add(tp);

                    dgv.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
                    //最后一行颜色
                    //dgv.Rows[dgv.Rows.Count - 1].DefaultCellStyle.BackColor = Color.AliceBlue;

                    /*
                    //将颜色赋值给dgv
                    if (list_CheckedDgv_Color.Count > 0)
                    {
                        foreach (int[] array_Color in list_CheckedDgv_Color)  //需注意dgv的空值情况
                        {
                            dgv.Rows[array_Color[0]].Cells[array_Color[1]].Style.BackColor = Color.FromArgb(array_Color[2]);
                        }
                        foreach(string[] array_formulaText in list_CheckedDgv_ToolTip)
                        {
                            dgv.Rows[int.Parse(array_formulaText[0].ToString())].Cells[int.Parse(array_formulaText[1].ToString())].
                                ToolTipText = array_formulaText[3];
                        }
                    }
                    */

                    //禁用dataGridView列排序
                    for (int j = 0; j < dgv.Columns.Count; j++)
                    {
                        dgv.Columns[j].SortMode = DataGridViewColumnSortMode.NotSortable;
                    }
                }
            }
        }

        //导入原始表
        private void Button_Import_Click(object sender, EventArgs e)
        {
            ProcessingFile("原始");
        }

        //导入审核用表
        private void Button_ImportAudit_Click(object sender, EventArgs e)
        {
            ProcessingFile("审核");
        }

        //导入汇总表
        private void Button_ImportSum_Click(object sender, EventArgs e)
        {
            ProcessingFile("汇总");
        }
        //打开报表并生成标准格式的报表,然后导入数据库
        private void ProcessingFile(string ReportType)
        {
            OpenFileDialog openExcelFiles = new OpenFileDialog
            {
                Multiselect = true,  //可同时打开多个文件
                Title = "请选择要导入的报表",
                Filter = "Excel文件(*.xls,*.xlsx) | *.xls;*.xlsx"
            };

            //读取Excel到DataSet和WorkBook
            if (openExcelFiles.ShowDialog() == DialogResult.OK)
            {
                list_Dgv = new List<DataGridView>();
                dsForImportingChecking = new DataSet(ReportType);
                wbForImportingChecking = new HSSFWorkbook();

                foreach (string excelPath in openExcelFiles.FileNames)
                {
                    //清空
                    DataSet dsFromExcel = new DataSet();
                    HSSFWorkbook wbFromExcel = new HSSFWorkbook();
                    bs_Dgv = new BindingSource();

                    //NPoI方式读取Excel，分别存储到DS和WB
                    dsFromExcel = OperatingData.OperatingData.DSFromExcel(excelPath, ReportType);
                    wbFromExcel = OperatingData.OperatingData.WBFromExcel(excelPath);

                    if (dsFromExcel == null ) continue;
                    else if (dsFromExcel != null && dsFromExcel.Tables.Count == 0) continue;

                    if (ReportType == "原始" || ReportType == "汇总")
                    {
                        HSSFSheet st;

                        DataTable dtFromExcel = dsFromExcel.Tables[0];
                        if (dsForImportingChecking.Tables.Contains(dtFromExcel.TableName))  //如果两个表名字相同，直接跳过，不再进行校验，目前是G4及其附注表会出现这种情况
                        {
                            continue;
                        }
                        dsForImportingChecking.Tables.Add(dtFromExcel.Copy());

                        st = wbFromExcel.GetSheetAt(0) as HSSFSheet;
                        st.CopyTo(wbForImportingChecking, dtFromExcel.TableName.Substring(10), true, true);
                       
                        wbForImportingChecking.SummaryInformation = wbFromExcel.SummaryInformation;

                        bs_Dgv.DataSource = dtFromExcel;
                        DataGridView dgv = new DataGridView
                        {
                            Name = dtFromExcel.TableName.Substring(10),
                            DataSource = bs_Dgv
                        };

                        dgv.EndEdit();
                        bs_Dgv.EndEdit();
                        list_Dgv.Add(dgv);
                    }
                    else if (ReportType == "审核")
                    {
                        foreach (DataTable dtFromExcel in dsFromExcel.Tables)
                        {
                            bs_Dgv = new BindingSource();
                            dsForImportingChecking.Tables.Add(dtFromExcel.Copy());

                            bs_Dgv.DataSource = dtFromExcel;
                            DataGridView dgv = new DataGridView
                            {
                                Name = "Checking " + dtFromExcel.TableName,
                                DataSource = bs_Dgv
                            };

                            dgv.EndEdit();
                            bs_Dgv.EndEdit();
                            list_Dgv.Add(dgv);
                        }
                    }
                }
                DisplaySheet(list_Dgv);

                //导入数据库
                int importedDT = 0;
                foreach (DataTable dt in dsForImportingChecking.Tables)
                {
                    bool isImported = false;

                    if (ReportType == "原始" || ReportType == "汇总")
                        isImported = OperatingData.OperatingData.DTtoDB(dt, "db_1104Check");
                    else if (ReportType == "审核")
                        isImported = OperatingData.OperatingData.DTtoDB(dt, "db_1104Formula");

                    if (isImported == true) importedDT++;
                }
                MessageBox.Show("导入了" + importedDT + "张报表");
            }
        }

        //校验  从1104Formula数据库中取出审核公式，在原始表或汇总表中进行校验
        private void Button_Check_Click(object sender, EventArgs e)
        {
            if (dsForImportingChecking.DataSetName == "审核")
            {
                MessageBox.Show("当前导入的是审核表，不能进行校验，请导入汇总表或1104原始表");
                return;
            }

            #region 从数据库取出公式表
            DataSet dsFrom1104Formula = new DataSet();
            List<string> list_FormulaTableName = new List<string>();

            foreach (DataTable dt_Sum in dsForImportingChecking.Tables)
            {
                string reportName = dt_Sum.TableName.Substring(10);

                DataTable dtFrom1104Formula = OperatingData.OperatingData.DTfromDB("select * from " + reportName, "db_1104Formula");
                if (dtFrom1104Formula == null) continue;
                else if (dtFrom1104Formula.Rows.Count == 0) continue;

                dtFrom1104Formula.TableName = reportName;
                dsFrom1104Formula.Tables.Add(dtFrom1104Formula);
                list_FormulaTableName.Add(reportName);
            }
            #endregion

            //将dsForImporting转换成Excel，然后利用NPoI进行计算
            //HSSFWorkbook wbForChecking = OperatingData.OperatingData.DSForCheckingToExcel(dsForImportingChecking);
            CheckReport(wbForImportingChecking, dsForImportingChecking, dsFrom1104Formula, list_FormulaTableName);
        }

        private void CheckReport(HSSFWorkbook wb, DataSet ds, DataSet dsFrom1104Formula, List<string> list_FormulaTableName)
        {
            for (int i = 0; i < wb.NumberOfSheets; i++)
            {
                HSSFSheet sheet = wb.GetSheetAt(i) as HSSFSheet;   //读取wb中当前sheet的数据
                DataTable dt_FromSum = ds.Tables[ds.Tables[0].TableName.Substring(0, 10) + sheet.SheetName]; //取出对应ds中的dt，需加上前缀，如“汇总20190930”
                DataGridView dgv_FromSum = new DataGridView();

                //从list_Dgv中取出对应的Dgv
                foreach (DataGridView _dgv in list_Dgv)
                {
                    if (_dgv.Name == sheet.SheetName)
                    {
                        dgv_FromSum = _dgv;
                        break;
                    }
                }

                string dt_Formula_Name = sheet.SheetName;
                DataTable dtOfCheck = new DataTable();
                //如果DB中存在此表校验公式的表
                if (list_FormulaTableName.Contains(dt_Formula_Name))
                {
                    foreach (DataTable dt_Checking in dsFrom1104Formula.Tables)
                    {
                        if (dt_Checking.TableName == dt_Formula_Name)
                        {
                            dtOfCheck = dt_Checking;
                            break;
                        }
                    }

                    if (dtOfCheck.Rows.Count > 0)
                    {
                        //依次取出dtOfCheck中的每行，即每个公式，赋值到wbForImportingChecking中对应的sheet中，开始校验
                        foreach (DataRow checking_Row in dtOfCheck.Rows)
                        {
                            int checking_RowIndex = int.Parse(checking_Row[0].ToString());
                            int checking_ColumnIndex = int.Parse(checking_Row[1].ToString());

                            HSSFCell checking_Cell = new HSSFCell(wb, sheet, checking_RowIndex, short.Parse(checking_ColumnIndex.ToString()));
                            //checking_Cell = sheet.GetRow(checking_RowIndex).GetCell(checking_ColumnIndex) as HSSFCell;  //这样会出现null的Cell

                            string str_Formula_Original = checking_Row[2].ToString();
                            string str_Formula = str_Formula_Original;

                            dgv_FromSum.Rows[checking_RowIndex].Cells[checking_ColumnIndex].ToolTipText = str_Formula;

                            //检验单元格的值：（此时cell的ToString返回的是公式，需将其计算后，才能得到校验的值)
                            //计算公式
                            try
                            {
                                HSSFFormulaEvaluator ev = new HSSFFormulaEvaluator(wb);
                                HSSFFormulaEvaluator.SetupEnvironment(new string[] { wb.SummaryInformation.Title }, new HSSFFormulaEvaluator[] { ev });

                                Dictionary<string, IFormulaEvaluator> dic_Wb = new Dictionary<string, IFormulaEvaluator>();
                                dic_Wb.Add(wb.SummaryInformation.Title, ev as IFormulaEvaluator);

                                ev.SetupReferencedWorkbooks(dic_Wb);
                                ev.IgnoreMissingWorkbooks = true;

                                checking_Cell.SetCellFormula(str_Formula);
                                ev.EvaluateInCell(checking_Cell);
                            }
                            catch (Exception ex)
                            {
                                //公式计算有问题的单元格设置颜色和备注
                                //设置颜色
                                HSSFCellStyle cellStyleOfWrong= wbForImportingChecking.CreateCellStyle() as HSSFCellStyle;
                                cellStyleOfWrong.FillBackgroundColor = HSSFColor.DarkYellow.Index;
                                checking_Cell.CellStyle = cellStyleOfWrong;
                                //创建批注
                                HSSFPatriarch patr = sheet.CreateDrawingPatriarch() as HSSFPatriarch;
                                HSSFComment comment_NullWarning = patr.CreateComment(new HSSFClientAnchor(0, 0, 0, 0, checking_ColumnIndex, checking_RowIndex, checking_ColumnIndex, checking_RowIndex));
                                //设置批注的内容和作者
                                comment_NullWarning.String = new HSSFRichTextString("此单元格无法计算公式(有可能是未导入引用的汇总表)，请自行校验");
                                comment_NullWarning.Author = "LYF";
                                checking_Cell.CellComment = comment_NullWarning;

                                //赋值到Dgv中
                                dgv_FromSum.Rows[checking_RowIndex].Cells[checking_ColumnIndex].Style.BackColor = Color.SandyBrown;
                                dgv_FromSum.Rows[checking_RowIndex].Cells[checking_ColumnIndex].ToolTipText = "此单元格无法计算公式，请自行校验";

                                continue;
                                //throw new Exception(ex.Message);
                            }
                            string result = checking_Cell.ToString();

                            //赋值给对应的dt和dgv
                            dt_FromSum.Rows[checking_RowIndex][checking_ColumnIndex] = result;
                            dgv_FromSum.Rows[checking_RowIndex].Cells[checking_ColumnIndex].Value = result;
                            //设置颜色
                            HSSFCellStyle cellStyleForChecking = wbForImportingChecking.CreateCellStyle() as HSSFCellStyle;
                            if (double.Parse(result) == 0)
                            {
                                cellStyleForChecking.FillBackgroundColor = HSSFColor.Aqua.Index;
                                dgv_FromSum.Rows[checking_RowIndex].Cells[checking_ColumnIndex].Style.BackColor = Color.Aqua;
                            }
                            else if (double.Parse(result) != 0)
                            {
                                cellStyleForChecking.FillBackgroundColor = HSSFColor.Red.Index;
                                dgv_FromSum.Rows[checking_RowIndex].Cells[checking_ColumnIndex].Style.BackColor = Color.Red;
                            }
                            checking_Cell.CellStyle = cellStyleForChecking;
                        }

                        dgv_FromSum.CellClick += new DataGridViewCellEventHandler(Dgv_CellClick);  //添加点击单元格事件
                        dgv_FromSum.EndEdit();
                        bs_Dgv.EndEdit();
                    }
                }
            }
        }

        //单击单元格显示说明
        public void Dgv_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (sender != null)
                if (e.RowIndex > 0 && e.ColumnIndex > 0)
                    textBox1.Text = (sender as DataGridView).Rows[e.RowIndex].Cells[e.ColumnIndex].ToolTipText;
        }

        private void Button_DeleteSum_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("确定删除所有汇总表吗", "是否删除", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
            {
                OperatingData.OperatingData.DeleteTableInDB("db_1104Check");
            }
        }

        
    }
}
using FormulaReadTest;
using NPOI.HPSF;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml;

namespace OperatingData
{
    public class OperatingData
    {
        static string db_UserName = "lyf";
        static string db_PassWord = "HiAmigo168F";

        public static string Db_UserName { get => db_UserName; }  //只读
        public static string Db_PassWord { get => db_PassWord; }  //只读

        /// <summary>
        /// NPOI方式读取Excel
        /// </summary>
        /// <param name="excelPath">Excel文件路径</param>
        /// <returns>对应的DataSet</returns>
        public static DataSet DSFromExcel(string excelPath, string ReportType) //如果是1104或是汇总表的话，则只取第一个sheet即可
        {
            try
            {
                DataSet ds = new DataSet();

                using (FileStream fs = File.OpenRead(excelPath))   //打开excel文件
                {
                    IWorkbook wk = null;

                    string fileName = Path.GetFileNameWithoutExtension(fs.Name);
                    string extension = Path.GetExtension(fs.Name);

                    //判断excel文件类型
                    if (extension == ".xlsx" || extension == ".xls")
                    {
                        //判断excel的版本
                        if (extension == ".xlsx")
                        {
                            wk = new XSSFWorkbook(fs);
                        }
                        else
                        {
                            wk = new HSSFWorkbook(fs);
                        }
                    }

                    int sheetNum = 1;
                    if (ReportType == "原始" || ReportType == "汇总") //原始表或汇总表每个workbook都只有一个有效sheet，所以只取第一个sheet
                        sheetNum = 1;
                    else if (ReportType == "审核") //审核表一般都是多个sheet，所以需要取多次
                        sheetNum = wk.NumberOfSheets;

                    for (int i = 0; i < sheetNum; i++)
                    {
                        ISheet sheet = wk.GetSheetAt(i);   //读取当前表数据

                        #region DataTable命名--取文件名中的要素，根据原始表/汇总表/审核表
                        string dtName = null, dtName03 = null;

                        if (ReportType == "原始")
                        {

                        }

                        else if (ReportType == "汇总")
                        {
                            //汇总表格式为"汇总20190930G0100"
                            dtName = "汇总";

                            Regex reg = new Regex(@"20[1-7]\d-\d{2}-\d{2}");
                            Match m = reg.Match(fileName);

                            dtName += m.Value.Replace("-", "");

                            //取出报表名称，如：G01_I，接着转换成标准表名，如：G0101
                            string originalReportName = fileName.Substring(fileName.IndexOf("汇总_") + 3, fileName.IndexOf(" ") - fileName.IndexOf("汇总_") - 3);
                            dtName03 = InitTables.GenerateStandardSheetName(originalReportName);
                            dtName += dtName03;
                        }

                        else if (ReportType == "审核")
                        {
                            dtName = sheet.SheetName;  //审核表直接将sheet的名字取出来即可，为了效率暂时先这样取，后期会根据sheet表中的内容来确定dtName
                        }
                        #endregion

                        DataTable dt;
                        if (dtName != null) dt = new DataTable(dtName);
                        else dt = new DataTable("NoStandardDTName");

                        if (sheet.LastRowNum == 0)
                        {
                            ds.Tables.Add(dt);
                            continue;
                        }

                        #region 确定最大列数
                        //IRow headrow = sheet.GetRow(0);
                        int columnNumMax = 0;
                        if (ReportType == "原始")
                            columnNumMax = InitTables.InitOriginalColumnCount(dtName03);
                        else if (ReportType == "汇总")
                            columnNumMax = InitTables.InitAuditSumColumnCount(dtName03);
                        else if (ReportType == "审核")
                            columnNumMax = InitTables.InitAuditSumColumnCount(dtName);

                        if (columnNumMax == 0)
                        {
                            for (int rowCnt = 0; rowCnt <= sheet.LastRowNum; rowCnt++)//找所有行中最大的值
                            {
                                IRow row = sheet.GetRow(rowCnt);
                                if (row != null && row.LastCellNum > columnNumMax)
                                {
                                    columnNumMax = row.LastCellNum;
                                }
                            }
                        }
                        #endregion

                        #region 创建列
                       //创建列，按照最大的列数（审核表只需创建三列：带公式的单元格的行索引、列索引和公式）
                        if (ReportType == "审核")
                        {
                            DataColumn dataColumn01 = new DataColumn("RowIndex");
                            DataColumn dataColumn02 = new DataColumn("ColumnIndex");
                            DataColumn dataColumn03 = new DataColumn("公式");
                            dt.Columns.Add(dataColumn01);
                            dt.Columns.Add(dataColumn02);
                            dt.Columns.Add(dataColumn03);
                        }
                        else
                        {
                            for (int c = 0; c < columnNumMax; c++)
                            {
                                DataColumn dataColum = new DataColumn(c.ToString());
                                dt.Columns.Add(dataColum);
                            }
                        }
                        #endregion

                        #region 填充数据到DataTable
                        if (ReportType == "审核")
                        {
                            //审核表只取公式，其他数据不要
                            int formulaCount = 0;  //公式个数

                            for (int r = 0; r <= sheet.LastRowNum; r++)  //LastRowNum 是当前表的总行数
                            {
                                IRow row = sheet.GetRow(r);  //读取当前行数据

                                if (row != null)
                                {
                                    for (int k = 0; k < columnNumMax; k++)
                                    {
                                        ICell cell = row.GetCell(k);  //当前表格
                                        if (cell == null) continue;
                                        if (cell.CellType == CellType.Formula)
                                        {
                                            formulaCount++;
                                            DataRow dr = dt.NewRow();
                                            dr[0] = cell.RowIndex;
                                            dr[1] = cell.ColumnIndex;
                                            dr[2] = cell.CellFormula;
                                            dt.Rows.Add(dr);
                                        }
                                    }
                                }
                            }
                            ds.Tables.Add(dt);
                        }
                        else //填充其他表
                        {
                            for (int r = 0; r <= sheet.LastRowNum; r++)  //LastRowNum 是当前表的总行数
                            {
                                IRow row = sheet.GetRow(r);  //读取当前行数据
                                DataRow dr = dt.NewRow();

                                if (row != null)
                                {
                                    for (int k = 0; k < columnNumMax; k++)
                                    {
                                        ICell cell = row.GetCell(k);  //当前表格
                                        if (cell != null)
                                        {
                                            dr[k] = GetCellValue(cell);
                                        }
                                        else if (cell == null)
                                        {
                                            dr[k] = string.Empty;
                                        }
                                    }
                                }
                                else if (row == null)
                                {
                                    dr[0] = string.Empty;
                                }
                                dt.Rows.Add(dr);
                            }
                            ds.Tables.Add(dt);
                        }
                        #endregion
                    }
                }
                return ds;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return null;
            }
        }

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
                        return cell.DateCellValue.ToString("yyyyMMdd");
                    }
                    else //其它数字
                    {
                        return cell.NumericCellValue.ToString().Trim();
                    }
                case CellType.Unknown: //无法识别类型
                default: //默认类型
                    return cell.ToString().Trim();//
                case CellType.String: //string 类型
                    return cell.StringCellValue.Trim();
                case CellType.Formula: //带公式类型
                    try
                    {
                        /*这是计算出公式的值的步骤
                        HSSFFormulaEvaluator e = new HSSFFormulaEvaluator(cell.Sheet.Workbook);
                        e.EvaluateInCell(cell);
                        return cell.ToString();
                        */
                        return cell.CellFormula.Trim();
                    }
                    catch
                    {
                        return cell.NumericCellValue.ToString().Trim();
                    }
            }
        }

        public static HSSFWorkbook WBFromExcel(string excelPath)
        {
            HSSFWorkbook wb = null;

            using (FileStream fs = File.OpenRead(excelPath))   //打开excel文件
            {
                string extension = Path.GetExtension(fs.Name);

                //判断excel文件类型
                if (extension == ".xlsx" || extension == ".xls")
                {
                    //判断excel的版本
                    if (extension == ".xlsx")
                    {
                        //wb = new XSSFWorkbook(fs);
                    }
                    else
                    {
                        try
                        {
                            wb = new HSSFWorkbook(fs);
                            //创建SummaryInformation
                            SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
                            si.Title = Path.GetFileName(fs.Name);
                            //将文件名保存至workbook的SummaryInformation中（Title）
                            if (wb.SummaryInformation == null)
                                wb.SummaryInformation = si;
                            else
                                wb.SummaryInformation.Title = si.Title;
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message+"\n\""+ Path.GetFileName(fs.Name)+"\" 读取有问题");
                            
                        }
                    }
                }
            }
            return wb;
        }

        /// <summary>
        /// 从数据库取出数据到DataTable
        /// </summary>
        /// <param name="strSql">查询语句</param>
        /// <param name="dbName">数据库名称</param>
        /// <returns>取出的DataTable</returns>
        public static DataTable DTfromDB(string strSql, string dbName)
        {
            using (SqlConnection conn = new SqlConnection("server=" +Form1.serverIP+ ";database=" + dbName + ";uid="+Db_UserName+";pwd="+Db_PassWord))
            {
                try
                {
                    SqlDataAdapter sda = new SqlDataAdapter(strSql, conn);

                    DataTable dtSelected = new DataTable();

                    sda.Fill(dtSelected);
                    return dtSelected;
                }

                catch (Exception ex)
                {
                    return null;
                }
            }
        }

        /*暂时不做登录
        /// <summary>
        /// 从数据库取出数据到DataTable，专用做登录
        /// </summary>
        /// <param name="strSql">查询语句</param>
        /// <param name="dbName">数据库名称</param>
        /// <returns>取出的DataTable</returns>
        public static DataTable DT_LoginfromDB(string strSql, string dbName)
        {
            using (SqlConnection conn = new SqlConnection("server=" + Form_Login.serverIP + ";database=" + dbName + ";uid="+Db_UserName+";pwd="+Db_PassWord))
            {
                try
                {
                    SqlDataAdapter sda = new SqlDataAdapter(strSql, conn);
                    DataTable dtSelected = new DataTable();

                    sda.Fill(dtSelected);
                    return dtSelected;
                }

                catch (Exception ex)
                {
                    return null;
                }
            }
        }
        */

        /// <summary>
        /// 利用SqlBulkCopy将DataTable数据导入数据库
        /// </summary>
        /// <param name="dt_ToDB">待导入的DataTable</param>
        /// <param name="dbName">数据库名称</param>
        /// <returns>是否导入成功</returns>
        public static bool DTtoDB(DataTable dt_ToDB, string dbName)
        {
            if (dt_ToDB == null || dt_ToDB.Rows.Count == 0) { MessageBox.Show("要导入的表不存在数据，请重新导入。"); return false; }

            //判断要导入的表在数据库中是否存在
            string strSql_GetTableName = "SELECT NAME FROM sysobjects WHERE XTYPE = 'U' AND NAME='" + dt_ToDB.TableName + "' ORDER BY NAME";
            DataTable dt_TableName = DTfromDB(strSql_GetTableName, dbName);

            //不存在此表则创建新表
            if (dt_TableName == null || dt_TableName.Rows.Count == 0)
                CreateTableInDB(dt_ToDB, dbName);

            //存在此表则清空原表（选择是否导入）
            else if (dt_TableName.Rows.Count > 0)
            {
                //默认清空原表，直接导入即可
                //DialogResult dr = MessageBox.Show("数据库中已经存在"+ dt_ToDB.TableName+"，继续导入吗？将会覆盖原表", "是否清空表", MessageBoxButtons.OKCancel);
                //仍然导入，则先清空原数据
                //if (dr == DialogResult.OK)
                //{
                    using (SqlConnection sqlConn = new SqlConnection("server=" + Form1.serverIP + ";database=" + dbName + ";uid="+Db_UserName+";pwd="+Db_PassWord))
                    {
                        if (sqlConn.State == ConnectionState.Closed)
                            sqlConn.Open();
                        //清空数据库中对应的表数据
                        try
                        {
                            SqlCommand sqlComm = new SqlCommand("truncate table " + dt_ToDB.TableName, sqlConn);
                            int i =
                                sqlComm.ExecuteNonQuery();
                        }
                        catch (Exception ex)
                        {
                            if (sqlConn.State == ConnectionState.Open)
                                sqlConn.Close();

                            MessageBox.Show(ex.ToString());
                            return false;
                        }
                        if (sqlConn.State == ConnectionState.Open)
                            sqlConn.Close();
                    }
                //}
                //不导入数据
                //else if (dr == DialogResult.Cancel)
                //{
                //    return false;
                //}
            }

            //SqlBulkCopy批量导入数据
            using (SqlConnection sqlConn = new SqlConnection("server=" + Form1.serverIP + ";database=" + dbName + ";uid="+Db_UserName+";pwd="+Db_PassWord))
            {
                if (sqlConn.State == ConnectionState.Closed)
                    sqlConn.Open();

                using (SqlBulkCopy sbc = new SqlBulkCopy(sqlConn))
                {
                    try
                    {
                        sbc.BatchSize = dt_ToDB.Rows.Count;
                        sbc.DestinationTableName = dt_ToDB.TableName;

                        for (int c = 0; c < dt_ToDB.Columns.Count; c++)
                        {
                            sbc.ColumnMappings.Add(dt_ToDB.Columns[c].ColumnName, dt_ToDB.Columns[c].ColumnName);
                        }

                        sbc.WriteToServer(dt_ToDB);
                    }
                    catch (Exception ex)
                    {
                        if (sqlConn.State == ConnectionState.Open)
                            sqlConn.Close();

                        MessageBox.Show(ex.StackTrace+"\n\""+ dt_ToDB.TableName+"\" 导入数据库未成功");
                        return false;
                    }
                }

                if (sqlConn.State == ConnectionState.Open)
                    sqlConn.Close();
            }

            //MessageBox.Show("\"" + dt_ToDB.TableName + "\"" + "成功导入数据库");
            return true;
        }

        /// <summary>
        /// 在数据库dbName中创建表，表结构与(DataTable)dt一样
        /// </summary>
        /// <param name="dt">表(结构)</param>
        /// <param name="dbName">数据库的名字</param>
        /// <returns>是否创建成功</returns>
        private static bool CreateTableInDB(DataTable dt, string dbName)
        {
            //创建表（只有结构，没有数据）
            try
            {
                string connString = "Initial Catalog=" + dbName + ";" + "Data Source=" + Form1.serverIP + ";uid="+Db_UserName+";pwd="+Db_PassWord;
                SqlConnection conn = new SqlConnection();
                conn.ConnectionString = connString;
                conn.Open();

                string strSql = "CREATE TABLE " + dt.TableName + "(";

                if (dbName == "db_Scores")
                {
                    strSql += "[申报积分项目] nvarchar(100) NOT NULL," +
                        "[分值] nvarchar(10) NOT NULL," +
                        "[备注] nvarchar(100))";  //！！注意，不要随便设主键，否则SqlBulkCopy导入数据库后，会按照主键来进行排序！
                }
                else
                {
                    //！！注意列名中如果含有特殊字符，要加中括号[]引起来，防止sql无法识别
                    for (int c = 0; c < dt.Columns.Count; c++)
                    {
                        if (c != dt.Columns.Count - 1)
                            strSql += "[" + dt.Columns[c].ColumnName + "] nvarchar(50) NOT NULL,";  //！！注意，不要随便设主键，否则SqlBulkCopy导入数据库后，会按照主键来进行排序！
                        else if (c == dt.Columns.Count - 1)
                            strSql += "[" + dt.Columns[c].ColumnName + "] nvarchar(50) NOT NULL)";
                    }
                }

                SqlCommand cmd = new SqlCommand(strSql, conn);
                int i =
                    cmd.ExecuteNonQuery(); //对于对于 UPDATE、INSERT 和 DELETE 语句，返回值为该命令所影响的行数。对于所有其他 DML 语句，返回值都为 - 1。
                                          //对于DDL语句，比如 CREATE TABLE 或 ALTER TABLE，返回值为 0。
                conn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return false;
            }

            return true;
        }

        public static void DeleteTableInDB(string dbName)
        {
            int deletedCount = 0;
            string strSql_GetTableName = "SELECT NAME FROM sysobjects WHERE XTYPE = 'U' ORDER BY NAME";
            DataTable dt_TableName = DTfromDB(strSql_GetTableName, dbName);

            using (SqlConnection sqlConn = new SqlConnection("server=" + Form1.serverIP + ";database=" + dbName + ";uid="+Db_UserName+";pwd="+Db_PassWord))
            {
                for (int t = 0; t < dt_TableName.Rows.Count; t++)
                {
                    if (sqlConn.State == ConnectionState.Closed)
                        sqlConn.Open();
                    //清空数据库中对应的表数据
                    try
                    {
                        SqlCommand sqlComm = new SqlCommand("drop table " + dt_TableName.Rows[t][0].ToString(), sqlConn);
                        int i =
                            sqlComm.ExecuteNonQuery();

                        deletedCount++;
                    }
                    catch (Exception ex)
                    {
                        if (sqlConn.State == ConnectionState.Open)
                            sqlConn.Close();

                    }
                }

                if (sqlConn.State == ConnectionState.Open)
                    sqlConn.Close();
                MessageBox.Show("删除了" + deletedCount + "张报表");
            }
        }

        
        public static bool UpdateAuditState(int auditState, string dbName)
        {
            //if (dt_ToDB == null || dt_ToDB.Rows.Count == 0) {  return false; }

            //判断要导入的表在数据库中是否存在
            string strSql_GetTableName = "SELECT NAME FROM sysobjects WHERE XTYPE = 'U' AND NAME = 'tb_AuditState' ORDER BY NAME";
            DataTable dt_TableName = DTfromDB(strSql_GetTableName, dbName);

            //不存在此表
            if (dt_TableName == null || dt_TableName.Rows.Count == 0 || dt_TableName.Rows.Count >1)
            {
                MessageBox.Show("没有找到审核状态表，请先在数据库中创建此表");
                return false;
            }

            //进行更新，有则更新，无则插入
            else if (dt_TableName.Rows.Count == 1)
            {
                using (SqlConnection conn = new SqlConnection("server=" + Form1.serverIP + ";database=" + dbName + ";uid="+Db_UserName+";pwd="+Db_PassWord))
                {
                    try
                    {
                        int i = 0;
                        string str_Find = string.Format(@"select * from tb_AuditState where county = '{0}' and reportDate = '{1}'", Form1.userInfo.Substring(0, 2), Form1.auditDate);
                        SqlCommand comm = new SqlCommand(str_Find, conn);
                        if (conn.State == ConnectionState.Closed)
                            conn.Open();
                        
                        SqlDataReader sdr = comm.ExecuteReader();
                        if (sdr.HasRows == true)
                        {
                            i = 1;
                        }
                        sdr.Close();

                        string str_InsertUpdate = null;
                        if (i ==0) //无则插入
                        {
                            str_InsertUpdate = string.Format(@"INSERT INTO tb_AuditState (county, reportDate, state) VALUES ('{0}','{1}','{2}')",
                           Form1.userInfo.Substring(0, 2), Form1.auditDate, auditState);
                        }
                        else if (i == 1)  //有则更新
                        {
                            str_InsertUpdate = string.Format(@"update tb_AuditState set state='{0}' where county='{1}' and reportDate='{2}'", 
                                auditState, Form1.userInfo.Substring(0,2), Form1.auditDate);

                        }
                        //string str_InsertUpdate = string.Format(@"INSERT INTO tb_AuditState (county, reportDate, state) VALUES ('{0}','{1}','{2}') ON DUPLICATE KEY UPDATE state = '{2}'",
                        //    Form_Audit.userInfo.Substring(0, 2), Form_Audit.auditDate, auditState);
                        //string strUpdate = string.Format(@"update {0} set state='{1}' where county='{2}' and reportDate='{3}'", dtName, auditState, Form_Audit.userInfo.Substring(0,2), Form_Audit.auditDate);
                        SqlCommand commUpdate = new SqlCommand(str_InsertUpdate, conn);
                        if (conn.State == ConnectionState.Closed)
                            conn.Open();

                        int j =
                            commUpdate.ExecuteNonQuery();
                        if (j < 1) return false;
                        return true;
                    }

                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message + "\n");
                        return false;
                    }
                }
            }
            return true;
        }

        /// <summary>
        /// 将导入的汇总表转成Excel，以便计算公式然后审核
        /// </summary>
        /// <param name="dsForChecking">待转换的DS</param>
        /// <param name="filename">转换后的Excel</param>
        /// <returns>是否转换成功</returns>
        public static HSSFWorkbook DSForCheckingToExcel(DataSet dsForChecking)
        {
            using (dsForChecking)
            {
                HSSFWorkbook workBook = new HSSFWorkbook();              

                for (int i = 0; i < dsForChecking.Tables.Count; i++)
                {
                    string sheetName = dsForChecking.Tables[i].TableName;

                    ISheet sheet = workBook.CreateSheet(sheetName);
                    SheetFromDT(sheet, dsForChecking.Tables[i]);
                }
                return workBook;
            }
        }

        #region NPOI保存数据到excel
        /// <summary>
                /// 导出数据到excel中
                /// </summary>
                /// <param name="dataSet"></param>
                /// <param name="filename"></param>
                /// <returns></returns>
        public static bool DSToExcel(DataSet dataSet, string filename)
        {
            MemoryStream ms = new MemoryStream();
            using (dataSet)
            {
                IWorkbook workBook;
                //IWorkbook workBook=WorkbookFactory.Create(filename);

                string suffix = Path.GetExtension(filename);

                //string suffix = filename.Substring(filename.LastIndexOf(".") + 1, filename.Length - filename.LastIndexOf(".") - 1);
                if (suffix == ".xls")
                {
                    workBook = new HSSFWorkbook();
                }
                else
                    workBook = new XSSFWorkbook();

                for (int i = 0; i < dataSet.Tables.Count; i++)
                {
                    string sheetName = "";
                    //如果是生成表（NPL和SML），则sheet的名字改成“各支行汇总”
                    //if (dataSet.Tables[i].TableName.Contains("各支行汇总"))
                    //{

                    //if (dataSet.Tables[i].TableName.Contains("不良贷款"))
                    //    sheetName = "各支行汇总" + "不良贷款";
                    //else if (dataSet.Tables[i].TableName.Contains("关注类贷款"))
                    //    sheetName = "各支行汇总" + "关注类贷款";
                    //}
                    sheetName = dataSet.Tables[i].TableName;

                    ISheet sheet = workBook.CreateSheet(sheetName);
                    SheetFromDT(sheet, dataSet.Tables[i]);
                }
                workBook.Write(ms);

                try
                {
                    SaveToFile(ms, filename);
                    ms.Flush();
                    return true;
                }
                catch
                {
                    ms.Flush();
                    throw;
                }
            }
        }

        private static void SheetFromDT(ISheet sheet, DataTable dataTable)
        {
            IRow headerRow = sheet.CreateRow(0);
            //表头
            foreach (DataColumn column in dataTable.Columns)
                headerRow.CreateCell(column.Ordinal).SetCellValue(column.ColumnName);//If Caption not set, returns the ColumnName value

            int rowIndex = 1;
            foreach (DataRow row in dataTable.Rows)
            {
                IRow dataRow = sheet.CreateRow(rowIndex);
                foreach (DataColumn column in dataTable.Columns)
                {
                    dataRow.CreateCell(column.Ordinal).SetCellValue(row[column].ToString());
                }
                rowIndex++;
            }
        }

        private static void SaveToFile(MemoryStream ms, string fileName)
        {
            using (FileStream fs = new FileStream(fileName, FileMode.Create, FileAccess.Write))
            {
                byte[] data = ms.ToArray();         //转为字节数组 
                fs.Write(data, 0, data.Length);     //保存为Excel文件
                fs.Flush();
                data = null;
            }
        }
        #endregion

        //MD5加密
        public static string ComputeMD5Hash(string userPassword)
        {
            string strMD5Hash = "";

            MD5 md5 = new MD5CryptoServiceProvider();

            byte[] byteSource = System.Text.Encoding.UTF8.GetBytes(userPassword);

            byte[] byteMD5Hash = md5.ComputeHash(byteSource);

            for (int i = 0; i < byteMD5Hash.Length; i++)
            {
                strMD5Hash += byteMD5Hash[i].ToString("X2");
            }

            return strMD5Hash;
        }

        /// <summary>
        /// 更新数据表
        /// </summary>
        /// <param name="dtName">待更新的数据表名称</param>
        /// <param name="dbName">数据库名称</param>
        /// <returns>是否更新成功</returns>
        public static int UpdateDB(string strUpdate, string dbName)
        {
            using (SqlConnection conn = new SqlConnection("server=" + Form1.serverIP + ";database=" + dbName + ";uid="+Db_UserName+";pwd="+Db_PassWord))
            {
                try
                {
                    SqlCommand comm = new SqlCommand(strUpdate, conn);
                    if (conn.State == ConnectionState.Closed)
                        conn.Open();

                    int i =
                        comm.ExecuteNonQuery();
                    if (i < 1) return 0;
                    return i;
                }

                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message + "\n");
                    return -1;
                }
            }
        }

        /// <summary>
        /// 更新审核状态
        /// </summary>
        /// <param name="dtName">待更新的数据表名称</param>
        /// <param name="dbName">数据库名称</param>
        /// <returns>是否更新成功</returns>
        public static int UpdateState(int auditState,string dtName, string dbName )
        {
            //Control ctl = Form1.FromHandle();

            using (SqlConnection conn = new SqlConnection("server=" + Form1.serverIP + ";database=" + dbName + ";uid="+Db_UserName+";pwd="+Db_PassWord))
            {
                try
                {
                    string str_InsertUpdate = string.Format(@"INSERT INTO {0} (county, reportDate, state) VALUES ({1},{2},{3}) DUPLICATE KEY UPDATE state = {3}",
                        dtName, Form1.userInfo.Substring(0, 2), Form1.auditDate, auditState);
                    //string strUpdate = string.Format(@"update {0} set state='{1}' where county='{2}' and reportDate='{3}'", dtName, auditState, Form_Audit.userInfo.Substring(0,2), Form_Audit.auditDate);
                    SqlCommand comm = new SqlCommand(str_InsertUpdate, conn);
                    if (conn.State == ConnectionState.Closed)
                        conn.Open();

                    int i =
                        comm.ExecuteNonQuery();
                    if (i < 1) return 0;
                    return i;
                }

                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message + "\n");
                    return -1;
                }
            }
        }

        public static bool ReadServerInfoFromXML(string xmlFileName)
        {
            if (!File.Exists(xmlFileName))
            {
                return true;//不存在文件，也返回true，因为本地没有，可以随后生成
            }
            try
            {
                //创建一个XmlTextReader对象，读取XML数据
                XmlTextReader xmlReader = new XmlTextReader(xmlFileName);

                while (xmlReader.Read())
                {
                    if (xmlReader.Name.Equals("ServerIP") == true)
                    {
                        Form1.serverIP = xmlReader.ReadString().Trim();
                    }
                }
                xmlReader.Close();
            }
            catch (Exception ex)
            {
                return false;

            }
            return true;
        }

        //XmlTextWriter默认是覆盖以前的文件,如果此文件名不存在,它将创建此文件
        public static bool WriteServerInfoToXML(string xmlFileName, string elementString)
        {
            try
            {
                XmlTextWriter xmlWriter = new XmlTextWriter(xmlFileName, Encoding.UTF8);
                xmlWriter.Formatting = Formatting.Indented;
                //写入根元素
                xmlWriter.WriteStartElement("Server");
                //加入子元素
                xmlWriter.WriteElementString("ServerIP", elementString);

                //关闭根元素，并书写结束标签
                xmlWriter.WriteEndElement();
                //将XML写入文件并且关闭XmlTextWriter
                xmlWriter.Close();
            }
            catch (Exception ex)
            {
                return false;
            }
            return true;
        }

        public static bool ReadUserInfoFromXML(string xmlFileName, string[] elementString)
        {
            if (!File.Exists(xmlFileName))
            {
                return true;//不存在文件，也返回true，因为本地没有，可以随后生成
            }
            try
            {
                //创建一个XmlTextReader对象，读取XML数据
                XmlTextReader xmlReader = new XmlTextReader(xmlFileName);

                while (xmlReader.Read())
                {
                    if (xmlReader.Name.Equals("user") == true)
                    {
                        elementString[0] = xmlReader.ReadString().Trim();
                    }

                    else if (xmlReader.Name.Equals("passMD5") == true)
                    {
                        elementString[1] = xmlReader.ReadString().Trim();
                    }
                }
                xmlReader.Close();
            }
            catch (Exception ex)
            {
                return false;

            }
            return true;
        }

        //XmlTextWriter默认是覆盖以前的文件,如果此文件名不存在,它将创建此文件
        //进行md5加密后存储
        public static bool WriteUserInfoToXML(string xmlFileName, string[] elementString)
        {
            //string pass_MD5Hash = ComputeMD5Hash(loginPass);
            if (elementString.Length != 2)
            {
                return false;
            }
            try
            {
                XmlTextWriter xmlWriter = new XmlTextWriter(xmlFileName, Encoding.UTF8);
                xmlWriter.Formatting = Formatting.Indented;
                //写入根元素
                xmlWriter.WriteStartElement("userData");
                //加入子元素
                xmlWriter.WriteElementString("user", elementString[0]);
                xmlWriter.WriteElementString("passMD5", elementString[1]);

                //关闭根元素，并书写结束标签
                xmlWriter.WriteEndElement();
                //将XML写入文件并且关闭XmlTextWriter
                xmlWriter.Close();
            }
            catch (Exception ex)
            {
                return false;
            }
            return true;
        }
    }
}

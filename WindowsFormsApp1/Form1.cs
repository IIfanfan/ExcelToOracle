using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Data.OracleClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        private DataSet myDs = new DataSet();
        private string boolIsApporve = "N";
        private string boolNoApporve = "0";  //没有认证
        private string boolApporved = "1";  //已经认证

        private string connOA = ConfigurationManager.ConnectionStrings["siplpKq"].ConnectionString;
        private string connectionStringFormat = "Provider=Microsoft.Ace.OleDb.12.0;" + "data source= '{0}';Extended Properties='Excel 12.0; HDR=Yes; IMEX=1'";


        public Form1()
        {
            InitializeComponent();
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            openFileDialog1.InitialDirectory = "F:\\工作文档";
            openFileDialog1.Filter = "Excel files (*.xls)|*.xls|All files (*.*)|*.*";
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                this.txtFilePath.Text = openFileDialog1.FileName;
            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            System.Environment.Exit(0);
        }

        private void ImportData_Click(object sender, EventArgs e)
        {
            string connectString = string.Format(connectionStringFormat, this.txtFilePath.Text);
            try
            {
                myDs.Tables.Clear();
                myDs.Clear();
                OleDbConnection cnnxls = new OleDbConnection(connectString);
                OleDbDataAdapter myDa = new OleDbDataAdapter("select * from [Sheet1$]", cnnxls);
                myDa.Fill(myDs, "c");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            InsertData();

            //  var mytest =  GetData(this.txtFilePath.Text);
        }

        private bool CheckIsNumeric(string columnName)
        {
            string str = ",ISAPPROVE,";   //数据库字段是否认证
            return str.Contains("," + columnName.ToUpper() + ",");
        }
        private bool CheckIsMaterialNum(string columnName)
        {
            string str = ",MATERIALNUM,";   //数据库字段
            return str.Contains("," + columnName.ToUpper() + ",");
        }
        private DataTable GetData(string strPath)
        {
            DataTable dtbl = new DataTable();
            try
            {
                //string strCon = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + strPath + ";" + "Extended Properties=Excel 8.0;";

                string strCon = "Provider=Microsoft.Ace.OleDb.12.0;" + "data source=" + strPath + ";Extended Properties='Excel 12.0; HDR=Yes; IMEX=1'"; //此连接可以操作.xls与.xlsx文件 (支持Excel2003 和 Excel2007 的连接字符串)

                string strSheetName = "";
                using (OleDbConnection con = new OleDbConnection(strCon))
                {
                    con.Open();
                    myDs.Tables.Clear();
                    myDs.Clear();

                }
                String strCmd = "select * from [" + strSheetName + "]";
                OleDbDataAdapter cmd = new OleDbDataAdapter(strCmd, strCon);
                cmd.Fill(dtbl);
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
            return dtbl;
        }
        public static void WriteLog(String msg)
        {
            StreamWriter writer = null;
            try
            {
                writer = File.AppendText(@"D:\FTLDucment\log\a.log");
                writer.WriteLine("{0} {1}", DateTime.Now.ToString("yyyy/MM/dd hh:mm:ss"), msg);
                writer.Flush();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (writer != null)
                {
                    writer.Close();
                }
            }


        }

        private void InsertData()
        {

            int intOk = 0;
            int intFail = 0;

            if (myDs != null && myDs.Tables[0].Rows.Count > 0)
            {
                OracleConnection conn = new OracleConnection(connOA);
                conn.Open();
                OracleCommand com = null;

                #region 组装字段列表
                string insertColumnString = "MATERIALMANAGEID,";
                DataTable dt = myDs.Tables[0];
                int k = 0;
                foreach (DataColumn col in dt.Columns)
                {
                    insertColumnString += string.Format("{0},", col.ColumnName);
                }
                insertColumnString = insertColumnString.Trim(',');

                #endregion

                try
                {
                    int id = 1;  //零时变量用于数据库ID自增
                    foreach (DataRow dr in dt.Rows)
                    {
                        if (dr[0].ToString() == "")
                        {
                            continue;
                        }

                        #region 组装Sql语句
                        string insertValueString = id.ToString() + ",";
                        string updateValueString = "MATERIALMANAGEID=" + id.ToString() + ",";
                        string MaterialNum = dr["MATERIALNUM"].ToString().Replace("<空>", "");

                        #region 拼接Sql字符串

                        for (int i = 0; i < dt.Columns.Count; i++)
                        {
                            string originalValue = dr[i].ToString().Replace("<空>", "");

                            if (!string.IsNullOrEmpty(originalValue))
                            {
                                if (CheckIsNumeric(dt.Columns[i].ColumnName))
                                {
                                    if (originalValue == boolIsApporve)
                                    {
                                        originalValue = boolNoApporve;
                                        insertValueString += string.Format("'{0}',", Convert.ToDecimal(originalValue));
                                        updateValueString += string.Format("{0}='{1}',", dt.Columns[i].ColumnName, Convert.ToDecimal(originalValue));
                                    }
                                    else
                                    {
                                        originalValue = boolApporved;
                                        insertValueString += string.Format("'{0}',", Convert.ToDecimal(originalValue));
                                        updateValueString += string.Format("{0}='{1}',", dt.Columns[i].ColumnName, Convert.ToDecimal(originalValue));
                                    }
                                }
                                else
                                {
                                    if (CheckIsMaterialNum(dt.Columns[i].ColumnName))
                                    {
                                        MaterialNum = originalValue;
                                        insertValueString += string.Format("'{0}',", originalValue);
                                        updateValueString += string.Format("{0}='{1}',", dt.Columns[i].ColumnName, originalValue);
                                    }
                                    else
                                    {
                                        insertValueString += string.Format("'{0}',", originalValue);
                                        updateValueString += string.Format("{0}='{1}',", dt.Columns[i].ColumnName, originalValue);
                                    }
                                }
                            }
                            else
                            {
                                insertValueString += string.Format("NULL,");
                                updateValueString += string.Format("{0}=NULL,", dt.Columns[i].ColumnName);
                            }
                        }
                        insertValueString = insertValueString.Trim(',');
                        updateValueString = updateValueString.Trim(',');
                        #endregion

                        string insertSql = string.Format(@"INSERT INTO BASE_MATERIALMANAGE ({0}) VALUES({1})", insertColumnString, insertValueString);
                        string updateSql = string.Format("Update BASE_MATERIALMANAGE set {0} Where MATERIALNUM='{1}' ", updateValueString, MaterialNum);
                        string checkExistSql = string.Format("Select count(*) from BASE_MATERIALMANAGE where MATERIALNUM='{0}' ", MaterialNum);
                        #endregion

                        #region 写入数据
                        try
                        {
                            com = new OracleCommand();
                            com.Connection = conn;
                            bool succeed = false;
                            com.CommandText = checkExistSql;
                            object objCount = com.ExecuteScalar();
                            bool exist = Convert.ToInt32(objCount) > 0;
                            if (exist)
                            {
                                //需要更新
                                com.CommandText = updateSql;
                                succeed = com.ExecuteNonQuery() > 0;
                            }
                            else
                            {
                                //需要插入
                                com.CommandText = insertSql;
                                succeed = com.ExecuteNonQuery() > 0;
                            }

                            if (succeed)
                            {
                                intOk++;
                            }
                            else
                            {
                                intFail++;
                            }

                        }
                        catch (Exception ex)
                        {

                            intFail++;
                            //break;
                            WriteLog(insertSql);
                        }

                        id++;
                        #endregion
                    }

                    #region 关闭
                    if (conn != null && conn.State != ConnectionState.Closed)
                    {
                        conn.Close();
                    }
                    if (com != null)
                    {
                        com.Dispose();
                    }
                    #endregion
                }
                catch (Exception ex)
                {
                }

                if (intOk > 0 || intFail > 0)
                {
                    string tips = string.Format("数据导入成功：{0}个，失败：{1}个", intOk, intFail);
                    MessageBox.Show(tips);
                }
            }
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Data;
using System.Data.Odbc;
using System.Data.OleDb;
using System.Configuration;
using Word = Microsoft.Office.Interop.Word;

namespace Lowis_Reports_Testing.ObjectLibrary
{
    class Helper
    {
        public enum CompareType
        {
            Equal,
            Contians,
            Tolerance,
            DecimalRound
        }
        private DataTable _dtRep = new DataTable();
        public DataTable dtRep
        {
            get { return _dtRep; }
            set { _dtRep = value; }
        }
        int counter = 1;
        public DataTable dtFromExcelFile(string filepath, string sheetname)
        {
            try
            {
                DataTable dtble = new DataTable();

                OdbcConnection oconn = new OdbcConnection();
                oconn.ConnectionString = ConfigurationManager.ConnectionStrings["ReportLinks"].ToString() + filepath;
                string odbccmdtext = "Select * from [" + sheetname + "$]";
                OdbcCommand ocmd = new OdbcCommand(odbccmdtext, oconn);
                oconn.Open();
                OdbcDataAdapter da = new OdbcDataAdapter(ocmd);
                da.Fill(dtble);
                oconn.Close();
                return dtble;
            }
            catch
            {
                throw new Exception();
            }

        }

        public DataTable dtFromExcelFile(string filepath, string sheetname, string filtercolumnName, string filtervalue)
        {
            try
            {
                DataTable dtble = new DataTable();
                //string strFileName = filepath;
                //string strConn = "Provider=Microsoft.Jet.OLEDB.4.0;";
                //strConn += "Data Source= " + strFileName + "; Extended Properties='Excel 8.0;HDR=No;IMEX=1'";
                this.LogtoTextFile("Inside Get Excel Connection");
                OdbcConnection oconn = new OdbcConnection();
                oconn.ConnectionString = ConfigurationManager.ConnectionStrings["ReportLinks"].ToString() + filepath;
                string odbccmdtext = "Select * from [" + sheetname + "$]  where " + filtercolumnName + "='" + filtervalue + "'";
                OdbcCommand ocmd = new OdbcCommand(odbccmdtext, oconn);
                try
                {
                    oconn.Open();
                }
                catch (Exception ex)
                {
                    this.LogtoTextFile("c error" + ex.Message);
                }
                OdbcDataAdapter da = new OdbcDataAdapter(ocmd);
                da.Fill(dtble);
                oconn.Close();

                //OleDbConnection oconn = new OleDbConnection();
                //oconn.ConnectionString = strConn;
                //string odbccmdtext = "Select * from [" + sheetname + "$]  where " + filtercolumnName + "='" + filtervalue + "'";
                //OleDbCommand ocmd = new OleDbCommand(odbccmdtext, oconn);
                //oconn.Open();
                //OleDbDataAdapter da = new OleDbDataAdapter(ocmd);
                //da.Fill(dtble);
                //oconn.Close();

                this.LogtoTextFile("Out from  Get Excel Connection");
                return dtble;
            }
            catch
            {
                throw new Exception();
            }

        }
        public void AreEqual(string tcnameid, string linkName, string VerifyParameter, string exp, string act, CompareType compareOperator)
        {

            if (counter == 1)
            {
                dtRep.Columns.Add("TestCaseNameORId");
                dtRep.Columns.Add("LinkName");
                dtRep.Columns.Add("VerifyParameter");
                dtRep.Columns.Add("Expected");
                dtRep.Columns.Add("Actual");
                dtRep.Columns.Add("Result");
            }
            DataRow dr = dtRep.NewRow();
            switch (compareOperator.ToString().ToLower())
            #region OperatorsofVerify
            {

                case "equal":
                    {
                        if (exp.Length > 0)
                        {
                            if (exp.ToLower().Trim() == act.ToLower().Trim())
                            {
                                dr["TestCaseNameORId"] = tcnameid;
                                dr["LinkName"] = linkName;
                                dr["VerifyParameter"] = VerifyParameter;
                                dr["Expected"] = exp;
                                dr["Actual"] = trimcustom(act);
                                dr["Result"] = "Pass";
                            }
                            else
                            {
                                dr["TestCaseNameORId"] = tcnameid;
                                dr["LinkName"] = linkName;
                                dr["VerifyParameter"] = VerifyParameter;
                                dr["Expected"] = exp;
                                dr["Actual"] = trimcustom(act);
                                dr["Result"] = "Fail";
                            }
                        }
                        break;
                    }
                case "contains":
                    {
                        if (exp.Length > 0)
                        {
                            if (cleanIntermediateWhiteSpaces(act).ToLower().Contains(cleanIntermediateWhiteSpaces(exp).ToLower()))
                            {
                                dr["TestCaseNameORId"] = tcnameid;
                                dr["LinkName"] = linkName;
                                dr["VerifyParameter"] = VerifyParameter;
                                dr["Expected"] = exp;
                                dr["Actual"] = trimcustom(act);
                                dr["Result"] = "Pass";
                            }
                            else
                            {
                                dr["TestCaseNameORId"] = tcnameid;
                                dr["LinkName"] = linkName;
                                dr["VerifyParameter"] = VerifyParameter;
                                dr["Expected"] = exp;
                                dr["Actual"] = trimcustom(act);
                                dr["Result"] = "Fail";
                            }
                        }
                        break;
                    }

                case "tolerance":
                    {
                        break;
                    }

                case "decimalround":
                    {
                        break;
                    }
                default:
                    {
                        break;
                    }

            }
            #endregion


            if (dr["TestCaseNameORId"].ToString().Length > 0)
            {
                dtRep.Rows.Add(dr);
            }
            counter++;
        }

        public void LogtoFileCSV(DataTable dtin)
        {
            char delm = '\u0022';
            StringBuilder sb = new StringBuilder();
            if (dtin.Rows.Count > 0)
            {
                if (System.IO.File.Exists(ConfigurationManager.AppSettings["logfile"]) == false)
                {
                    //Adding Header Row only once
                    for (int kk = 0; kk < dtin.Columns.Count; kk++)
                    {
                        sb.Append(delm + dtin.Columns[kk].ColumnName.ToString() + delm + ",");

                    }

                    sb.Append(Environment.NewLine);
                }

                for (int i = 0; i < dtin.Rows.Count; i++)
                {
                    for (int kk = 0; kk < dtin.Columns.Count; kk++)
                    {
                        sb.Append(delm + dtin.Rows[i][kk].ToString() + delm + ",");
                    }
                    sb.Append(Environment.NewLine);
                }

                System.IO.File.AppendAllText(ConfigurationManager.AppSettings["logfile"], sb.ToString());
            }
        }

        public void LogTabletoWordFile(DataTable dtin)
        {
            try
            {
                this.LogtoTextFile("Inside LogTabletoWordFile method");
                object oMissing = System.Reflection.Missing.Value;
                object oEndOfDoc = "\\endofdoc";
                Microsoft.Office.Interop.Word._Application objWord;
                Microsoft.Office.Interop.Word._Document objDoc;
                objWord = new Microsoft.Office.Interop.Word.Application();
                objWord.Visible = false;
                objDoc = objWord.Documents.Add(ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing);
                int i = 0;
                int j = 0;
                Microsoft.Office.Interop.Word.Table objTable;
                Microsoft.Office.Interop.Word.Range wrdRng = objDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                objTable = objDoc.Tables.Add(wrdRng, dtin.Rows.Count, dtin.Columns.Count, ref oMissing, ref oMissing);
                objTable.Range.ParagraphFormat.SpaceAfter = 4;
                for (j = 1; j <= dtin.Columns.Count; j++)
                {
                    objTable.Cell(1, j).Range.Text = dtin.Columns[j - 1].ColumnName.ToString();
                    objTable.Cell(1, j).Borders[Word.WdBorderType.wdBorderLeft].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    objTable.Cell(1, j).Borders[Word.WdBorderType.wdBorderRight].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    objTable.Cell(1, j).Borders[Word.WdBorderType.wdBorderTop].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    objTable.Cell(1, j).Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                }
                objTable.Rows.Add();
                for (i = 2; i <= dtin.Rows.Count+1;i++)
                {
                    for (j = 1; j <= dtin.Columns.Count; j++)
                    {
                        objTable.Cell(i, j).Range.Text = dtin.Rows[i - 2][j - 1].ToString();
                        objTable.Cell(i, j).Range.Borders[Word.WdBorderType.wdBorderLeft].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                        objTable.Cell(i, j).Range.Borders[Word.WdBorderType.wdBorderRight].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                        objTable.Cell(i, j).Range.Borders[Word.WdBorderType.wdBorderTop].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                        objTable.Cell(i, j).Range.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    }
                    objTable.Rows.Add();
                }
                objTable.Rows[1].Range.Font.ColorIndex = Word.WdColorIndex.wdBlue;
                objTable.Rows[1].Range.Font.Bold = 1;
                objTable.Rows[1].Range.Font.Italic = 1;
                this.LogtoTextFile("Before Save working fine");
                //objTable.Borders[Word.WdBorderType.wdBorderLeft].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                //objTable.Borders[Word.WdBorderType.wdBorderRight].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                //objTable.Borders[Word.WdBorderType.wdBorderTop].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                //objTable.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                objDoc.SaveAs(ConfigurationManager.AppSettings["wordfile"]);
                objDoc.Close();
            }
            catch(Exception ex)
            {
                this.LogtoTextFile("Got Exception from method [LogTabletoWordFile] " + ex.Message);
            }

        }

        public void UpdateExcelFileColumn(string filepath, string sheetName, string columnName, string columnValue, string filterColName, string filterColValue)
        {
            OdbcConnection oconn = null;
            try
            {
                oconn = new OdbcConnection();
                oconn.ConnectionString = ConfigurationManager.ConnectionStrings["ReportLinks"].ToString() + filepath;
                string odbccmdtext = "Update [" + sheetName + "$] Set " + columnName + "='" + columnValue + "' where " + filterColName + "='" + filterColValue + "' ";
                OdbcCommand ocmd = new OdbcCommand(odbccmdtext, oconn);
                oconn.Open();
                ocmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                this.LogtoTextFile("Exception from update " + ex.Message);

            }
            finally
            {
                oconn.Close();
            }

        }

        private string trimcustom(string inp)
        {
            string op = "";
            char[] chartotrim = { ' ', '\n', '\t' };
            op = inp.Trim(chartotrim);
            string fop = op.Replace('\n', ' ');
            fop = fop.Replace('\r', ' ');
            if (fop.Length > 255)
            {
                // trim charts to 255 only 
                fop = fop.Substring(0, 255);
            }
            return fop;
        }
        public void LogtoTextFile(string msg)
        {
            System.IO.File.AppendAllText(ConfigurationManager.AppSettings["logtextfile"], "[" + System.DateTime.Now.ToString() + "] :" + msg + System.Environment.NewLine);


        }
        private string cleanIntermediateWhiteSpaces(string strinput)
        {

            string pattn = "\\s+";
            Regex re = new Regex(pattn);
            string retstring = re.Replace(strinput, " ");

            string op = "";
            char[] chartotrim = { ' ', '\n', '\t' };
            op = retstring.Trim(chartotrim);
            string fop = op.Replace('\n', ' ');
            fop = fop.Replace('\r', ' ');
            return retstring;

        }

        
    }
}

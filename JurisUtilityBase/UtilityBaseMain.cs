using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Globalization;
using Gizmox.Controls;
using JDataEngine;
using JurisAuthenticator;
using JurisUtilityBase.Properties;
using System.Data.OleDb;

namespace JurisUtilityBase
{
    public partial class UtilityBaseMain : Form
    {
        #region Private  members

        private JurisUtility _jurisUtility;

        #endregion

        #region Public properties

        public string CompanyCode { get; set; }

        public string JurisDbName { get; set; }

        public string JBillsDbName { get; set; }

        private int clisysnbr = 0;

        private int matsysnbr = 0;

        private bool startLooking = false;

        #endregion

        #region Constructor

        public UtilityBaseMain()
        {
            InitializeComponent();
            _jurisUtility = new JurisUtility();
        }

        #endregion

        #region Public methods

        public void LoadCompanies()
        {
            var companies = _jurisUtility.Companies.Cast<object>().Cast<Instance>().ToList();
//            listBoxCompanies.SelectedIndexChanged -= listBoxCompanies_SelectedIndexChanged;
            listBoxCompanies.ValueMember = "Code";
            listBoxCompanies.DisplayMember = "Key";
            listBoxCompanies.DataSource = companies;
//            listBoxCompanies.SelectedIndexChanged += listBoxCompanies_SelectedIndexChanged;
            var defaultCompany = companies.FirstOrDefault(c => c.Default == Instance.JurisDefaultCompany.jdcJuris);
            if (companies.Count > 0)
            {
                listBoxCompanies.SelectedItem = defaultCompany ?? companies[0];
            }
        }

        #endregion

        #region MainForm events

        private void Form1_Load(object sender, EventArgs e)
        {
        }

        private void listBoxCompanies_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_jurisUtility.DbOpen)
            {
                _jurisUtility.CloseDatabase();
            }
            CompanyCode = "Company" + listBoxCompanies.SelectedValue;
            _jurisUtility.SetInstance(CompanyCode);
            JurisDbName = _jurisUtility.Company.DatabaseName;
            JBillsDbName = "JBills" + _jurisUtility.Company.Code;
            _jurisUtility.OpenDatabase();
            if (_jurisUtility.DbOpen)
            {
                ///GetFieldLengths();
            }
            
            comboBoxClient.ClearItems();
            DataSet myRSPC2 = new DataSet();
            string SQLPC2 = "select dbo.jfn_FormatClientCode(clicode)  + '    ' + clireportingname as PC from client where clisysnbr in (select distinct matclinbr from matter) order by dbo.jfn_FormatClientCode(clicode)";
            myRSPC2 = _jurisUtility.RecordsetFromSQL(SQLPC2);

            if (myRSPC2 == null || myRSPC2.Tables.Count == 0 || myRSPC2.Tables[0].Rows.Count == 0)
            {
                MessageBox.Show("There are no Clients. The tool will now exit", "Data Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Close();
            }
            else
            {
                foreach (DataRow dr in myRSPC2.Tables[0].Rows)
                    comboBoxClient.Items.Add(dr["PC"].ToString());
                comboBoxClient.SelectedIndex = 0;
            }
            startLooking = true;




        }



        #endregion

        #region Private methods

        private void DoDaFix()
        {
            matsysnbr = getMatSysNbr(this.comboBoxMatter.GetItemText(this.comboBoxMatter.SelectedItem).Split(' ')[0]);
            if (matsysnbr != 0)
            {
                //get last bill and see if it was unposted
                string sql = "SELECT  cast(max([LHBillNbr]) as int) as highestbill FROM [LedgerHistory] where lhmatter = " + matsysnbr.ToString();
                int lastBill = 0;
                DataSet dds = _jurisUtility.RecordsetFromSQL(sql);
                if (dds == null || dds.Tables.Count == 0 || dds.Tables[0].Rows.Count == 0 || !isValidInt(dds.Tables[0].Rows[0][0].ToString()))
                {
                    MessageBox.Show("There are no bills for that Matter. Select another Matter", "Data Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    foreach (DataRow dr in dds.Tables[0].Rows)
                    {
                        lastBill = Convert.ToInt32(dr[0].ToString());
                    }
                    labelLastBill.Text = "Last Bill: " + lastBill.ToString();
                    sql = "SELECT  * FROM [LedgerHistory] where  lhbillnbr = " + lastBill.ToString() + " and lhtype in ('A', 'B', 'C') and lhmatter = " + matsysnbr.ToString();
                    dds.Clear();
                    dds = _jurisUtility.RecordsetFromSQL(sql);
                    if (dds == null || dds.Tables.Count == 0 || dds.Tables[0].Rows.Count == 0)
                    {
                        MessageBox.Show("The last bill for that Matter was not unposted. Select another Matter", "Data Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    else
                    {
                        //it WAS unposted - display message, get confirmation and make change
                        DialogResult dresult = MessageBox.Show("The last bill: " + lastBill.ToString() + " was unposted. Proceed to update matter?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (dresult == DialogResult.Yes)
                        {
                            //run the fix
                            sql = "SELECT top 2 [ID], MatDateLastBill,MatARLastBill,MatAdjSinceLastBill,MatPaySinceLastBill " +
                                 " FROM[Matter_Log] " +
                                 "  where matsysnbr = " + matsysnbr.ToString() + " and recordtype = 2 " +
                                 "  order by id desc;";
                            dds.Clear();
                            dds = _jurisUtility.RecordsetFromSQL(sql);
                            if (dds == null || dds.Tables.Count == 0 || dds.Tables[0].Rows.Count == 0)
                            {
                                MessageBox.Show("There is no history in the log files. Contact Juris Professional Services", "Data Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                            else
                            {
                                int count = 1;
                                foreach (DataRow dr in dds.Tables[0].Rows)
                                {
                                    if (count != 1)
                                    {
                                        sql = "update matter set MatDateLastBill = '" + dr[1].ToString() + "', MatARLastBill = " + dr[2].ToString() + ", " +
                                            " MatAdjSinceLastBill = " + dr[3].ToString() + ", MatPaySinceLastBill = " + dr[3].ToString() + " where matsysnbr = " + matsysnbr.ToString();
                                        _jurisUtility.ExecuteNonQuery(0, sql);
                                    }
                                    count++;
                                }
                                MessageBox.Show("The process completed without error.", "Completed", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }

                        }
                        else
                        {
                            MessageBox.Show("No changes were made to your data. The tool will now exit.", "Finished", MessageBoxButtons.OK, MessageBoxIcon.None);
                            this.Close();
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("No matsys was found for that matter. Contact Juris Professional Services", "Finished", MessageBoxButtons.OK, MessageBoxIcon.None);
            }
        }

        private bool isValidInt(string test)
        {
            try
            {
                int aa = Convert.ToInt32(test);
                return true;
            }
            catch (Exception vv)
            {
                return false;
            }

        }


        private bool VerifyFirmName()
        {
            //    Dim SQL     As String
            //    Dim rsDB    As ADODB.Recordset
            //
            //    SQL = "SELECT CASE WHEN SpTxtValue LIKE '%firm name%' THEN 'Y' ELSE 'N' END AS Firm FROM SysParam WHERE SpName = 'FirmName'"
            //    Cmd.CommandText = SQL
            //    Set rsDB = Cmd.Execute
            //
            //    If rsDB!Firm = "Y" Then
            return true;
            //    Else
            //        VerifyFirmName = False
            //    End If

        }

        private bool FieldExistsInRS(DataSet ds, string fieldName)
        {

            foreach (DataColumn column in ds.Tables[0].Columns)
            {
                if (column.ColumnName.Equals(fieldName, StringComparison.OrdinalIgnoreCase))
                    return true;
            }
            return false;
        }


        private static bool IsDate(String date)
        {
            try
            {
                DateTime dt = DateTime.Parse(date);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static bool IsNumeric(object Expression)
        {
            double retNum;

            bool isNum = Double.TryParse(Convert.ToString(Expression), System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out retNum);
            return isNum; 
        }

        private void WriteLog(string comment)
        {
            var sql =
                string.Format("Insert Into UtilityLog(ULTimeStamp,ULWkStaUser,ULComment) Values('{0}','{1}', '{2}')",
                    DateTime.Now, GetComputerAndUser(), comment);
            _jurisUtility.ExecuteNonQueryCommand(0, sql);
        }

        private string GetComputerAndUser()
        {
            var computerName = Environment.MachineName;
            var windowsIdentity = System.Security.Principal.WindowsIdentity.GetCurrent();
            var userName = (windowsIdentity != null) ? windowsIdentity.Name : "Unknown";
            return computerName + "/" + userName;
        }

        /// <summary>
        /// Update status bar (text to display and step number of total completed)
        /// </summary>
        /// <param name="status">status text to display</param>
        /// <param name="step">steps completed</param>
        /// <param name="steps">total steps to be done</param>


        private void DeleteLog()
        {
            string AppDir = Path.GetDirectoryName(Application.ExecutablePath);
            string filePathName = Path.Combine(AppDir, "VoucherImportLog.txt");
            if (File.Exists(filePathName + ".ark5"))
            {
                File.Delete(filePathName + ".ark5");
            }
            if (File.Exists(filePathName + ".ark4"))
            {
                File.Copy(filePathName + ".ark4", filePathName + ".ark5");
                File.Delete(filePathName + ".ark4");
            }
            if (File.Exists(filePathName + ".ark3"))
            {
                File.Copy(filePathName + ".ark3", filePathName + ".ark4");
                File.Delete(filePathName + ".ark3");
            }
            if (File.Exists(filePathName + ".ark2"))
            {
                File.Copy(filePathName + ".ark2", filePathName + ".ark3");
                File.Delete(filePathName + ".ark2");
            }
            if (File.Exists(filePathName + ".ark1"))
            {
                File.Copy(filePathName + ".ark1", filePathName + ".ark2");
                File.Delete(filePathName + ".ark1");
            }
            if (File.Exists(filePathName ))
            {
                File.Copy(filePathName, filePathName + ".ark1");
                File.Delete(filePathName);
            }

        }

            

        private void LogFile(string LogLine)
        {
            string AppDir = Path.GetDirectoryName(Application.ExecutablePath);
            string filePathName = Path.Combine(AppDir, "VoucherImportLog.txt");
            using (StreamWriter sw = File.AppendText(filePathName))
            {
                sw.WriteLine(LogLine);
            }	
        }
        #endregion

        private void button1_Click(object sender, EventArgs e)
        {
            DoDaFix();
        }

        private void buttonReport_Click(object sender, EventArgs e)
        {

            System.Environment.Exit(0);
          
        }

        private void comboBoxClient_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (startLooking)
            {
                clisysnbr = getCliSysNbr(this.comboBoxClient.GetItemText(this.comboBoxClient.SelectedItem).Split(' ')[0]);
                if (clisysnbr != 0)
                {
                    comboBoxMatter.Enabled = true;
                    comboBoxMatter.ClearItems();
                    DataSet myRSPC2 = new DataSet();
                    string SQLPC2 = "select dbo.jfn_FormatMatterCode(MatCode)  + '    ' + matreportingname as PC from matter where matclinbr = " + clisysnbr.ToString() + " order by dbo.jfn_FormatMatterCode(MatCode)";
                    myRSPC2 = _jurisUtility.RecordsetFromSQL(SQLPC2);

                    if (myRSPC2 == null || myRSPC2.Tables.Count == 0 || myRSPC2.Tables[0].Rows.Count == 0)
                    {
                        MessageBox.Show("There are no Matters for that Client. Select another Client", "Data Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        this.Close();
                    }
                    else
                    {
                        foreach (DataRow dr in myRSPC2.Tables[0].Rows)
                            comboBoxMatter.Items.Add(dr["PC"].ToString());
                        comboBoxMatter.SelectedIndex = 0;
                    }
                }
                else
                {
                    MessageBox.Show("There are no Matters for that Client. Select another Client", "Data Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

                }
            }
        }

        private int getCliSysNbr(string clicode)
        {
            int clisys = 0;
            string sql = "select clisysnbr from client where dbo.jfn_FormatClientCode(clicode) = '" + clicode + "'";
            DataSet dds = _jurisUtility.RecordsetFromSQL(sql);
            if (dds == null || dds.Tables.Count == 0 || dds.Tables[0].Rows.Count == 0)
            {
                clisys = 0;
            }
            else
            {
                foreach (DataRow dr in dds.Tables[0].Rows)
                {
                    clisys = Convert.ToInt32(dr[0].ToString());
                }
            }
            return clisys;
        }

        private void comboBoxMatter_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private int getMatSysNbr(string matcode)
        {
            int matsys = 0;
            string sql = "select matsysnbr from matter where matclinbr = " + clisysnbr.ToString() + " and dbo.jfn_FormatMatterCode(MatCode) = '" + matcode + "'";
            DataSet dds = _jurisUtility.RecordsetFromSQL(sql);
            if (dds == null || dds.Tables.Count == 0 || dds.Tables[0].Rows.Count == 0)
            {
                matsys = 0;
            }
            else
            {
                foreach (DataRow dr in dds.Tables[0].Rows)
                {
                    matsys = Convert.ToInt32(dr[0].ToString());
                }
            }
            return matsys;
        }
    }
}

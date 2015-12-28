using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.Data;
using System.IO;
using System.Management;
using System.Management.Instrumentation;
using System.Diagnostics;





namespace LicenseCheck
{
    class Program
    {
        static void Main(string[] args)
        {
            string filepath = ConfigurationManager.AppSettings["excelfilepath"];
            string listTo = ConfigurationManager.AppSettings["listto"];
            string listCC = ConfigurationManager.AppSettings["listcc"];
            string strmsgfile = ConfigurationManager.AppSettings["subjecttextfile1"];
            string strmsgfileno = ConfigurationManager.AppSettings["subjecttextfile2"];
            Console.WriteLine("Getting Data from excel");
            Lowis_Reports_Testing.ObjectLibrary.Helper hlpr = new Lowis_Reports_Testing.ObjectLibrary.Helper();
            DataTable dtlictable = hlpr.dtFromExcelFile(filepath, "Sheet1");
            DateTime datesevendaysfromnow = DateTime.Now.AddDays(7.00);
            DataTable dtoutput = new DataTable();
            dtoutput.Columns.Add("MachineOwner");
            dtoutput.Columns.Add("MachineName");
            dtoutput.Columns.Add("ProductName");
            dtoutput.Columns.Add("Status");
            dtoutput.Columns.Add("ExpiryDate");
            StringBuilder sbout = new StringBuilder();
            string[] arrProducts = new string[] { "Matbal", "WellFlo", "PanSystem", "RFC", "Reo", "PVTFlex", "Lowis" };
            foreach (DataRow dr in dtlictable.Rows)
            {
                bool flagemty = false;
                Console.WriteLine("Processing for Machine  {0}", dr["MACHINE NAME"].ToString());
                if (dr["Active"].ToString().ToLower() == "y" )
                {
                    #region checkdateofeachproduct
                    
                    foreach (string product in arrProducts)
                    {
                        DataRow dro = dtoutput.NewRow();
                        string expDateProduct = dr[product].ToString();
                        try
                        {
                            DateTime dtExp = DateTime.Parse(expDateProduct);

                            if (dtExp < datesevendaysfromnow)
                            {
                                dro["MachineOwner"] = dr["MACHINE OWNER"].ToString();
                                dro["MachineName"] = dr["MACHINE NAME"].ToString();
                                dro["ProductName"] = product;
                                dro["Status"] = "Will Expire";
                                dro["ExpiryDate"] = dtExp.ToString("dd-MMM-yyyy");
                                flagemty = true;

                            }

                        }
                        catch(Exception ex)
                        {
                            flagemty = false;
                            Console.WriteLine("Expetion " + ex.Message);
                        }
                        if (flagemty)
                        {
                            dtoutput.Rows.Add(dro);
                        }

                    }
                    #endregion
                    
                }
            }
            hlpr.LogTabletoWordFile(dtoutput);

            TerminateProcessByForce("WINWORD*");
           
            if (dtoutput.Rows.Count > 0)
            {
                SendEmailNotification(ConfigurationManager.AppSettings["wordfile"],listTo,listCC, ReadTextfromTextFile(strmsgfile));
            }
            else
            {
                SendEmailNotification(null, listTo, listCC, ReadTextfromTextFile(strmsgfileno));
            }
            TerminateProcessByForce("WINWORD*");
            TaskKill("WINWORD*");
        }

        public static void SendEmailNotification(string attachmentfilename, string ListTo ,string ListCC, string messagebody)
        {
            if (ConfigurationManager.AppSettings["sendmail"].ToLower() == "y")
            {
                Lowis_Reports_Testing.ObjectLibrary.Helper h2 = new Lowis_Reports_Testing.ObjectLibrary.Helper();
                h2.LogtoTextFile("Send Email Started");
                Console.WriteLine("Sending Email .....");
                try
                {
                    h2.LogtoTextFile("adding receipeint");
                    System.Net.Mail.MailMessage message = new System.Net.Mail.MailMessage();
                    h2.LogtoTextFile("Email list obtained from Top " + ListTo);
                    string[] recipients = ListTo.Split(';');
                    h2.LogtoTextFile(recipients[0].ToString());
                    foreach (string recipient in recipients)
                    {
                        if (recipient.Length > 0)
                        {
                            message.To.Add(recipient);
                        }
                    }
                    string[] recipientscc = ListCC.Split(';');
                    h2.LogtoTextFile(recipients[0].ToString());
                    foreach (string irecipient in recipientscc)
                    {
                        if (irecipient.Length > 0)
                        {
                            message.CC.Add(irecipient);
                        }
                    }
                    h2.LogtoTextFile("added receipient");
                    message.From = new System.Net.Mail.MailAddress("ashok.krishna@me.weatherford.com");
                    System.Net.Mail.SmtpClient smtp = new System.Net.Mail.SmtpClient("mail2.weatherford.com");
                    smtp.Port = 25;
                    message.Subject = "License";
                    message.Body = messagebody;
                    if (attachmentfilename.Length > 0)
                    {
                        var attachment = new System.Net.Mail.Attachment(attachmentfilename);
                        message.Attachments.Add(attachment);
                    }
                    smtp.Send(message);
                }
                catch (Exception ex)
                {
                    h2.LogtoTextFile("Error in Sending Mails.." + ex.Message);
                    //  throw new Exception("Error in Sending Mails.." + ex.Message);
                }
                h2.LogtoTextFile("Send Email Completed");
                Console.WriteLine("Sent Email .....");
            }
        }

        public static string ReadTextfromTextFile(string filepath)
        {
            StringBuilder sb = new StringBuilder();
            string line = "";
             StreamReader sr = new StreamReader(filepath);
            while ( (line = sr.ReadLine()) != null )
            {
                sb.Append(line);
                sb.Append(Environment.NewLine);
            }
            sr.Close();
            return sb.ToString();
        }

        private static void TerminateProcessByForce(string strprocess)
        {
            try
            {
                //Assign the name of the process you want to kill on the remote machine
                string processName = strprocess;

                //Assign the user name and password of the account to ConnectionOptions object
                //which have administrative privilege on the remote machine.
                ConnectionOptions connectoptions = new ConnectionOptions();
                //   connectoptions.Username = @"YourDomainName\UserName";
                //  connectoptions.Password = "User Password";

                //IP Address of the remote machine
                string ipAddress = "127.0.0.1";
                ManagementScope scope = new ManagementScope(@"\\" + ipAddress + @"\root\cimv2", connectoptions);

                //Define the WMI query to be executed on the remote machine
                SelectQuery query = new SelectQuery("select * from Win32_process where name = '" + processName + "'");

                using (ManagementObjectSearcher searcher = new
                            ManagementObjectSearcher(scope, query))
                {
                    foreach (ManagementObject process in searcher.Get())
                    {

                        process.InvokeMethod("Terminate", null);
                        Console.WriteLine("Found and Terminated " + processName);

                    }
                }

            }
            catch (Exception ex)
            {
                //Log exception in exception log.
                //Logger.WriteEntry(ex.StackTrace);
                Console.WriteLine(ex.StackTrace);

            }
        }

        private static void TaskKill(string proc)
        {

            if (proc != "")
            {
                ProcessStartInfo startInfo = new ProcessStartInfo();
                startInfo.CreateNoWindow = false;
                startInfo.UseShellExecute = false;
                startInfo.FileName = "taskkill.exe";
                startInfo.WindowStyle = ProcessWindowStyle.Hidden;
                startInfo.Arguments = " /f /im "+ "\""+ proc + "\"";
                try
                {
                    using (Process exeProcess = Process.Start(startInfo))
                    {
                        exeProcess.WaitForExit();
                    }
                }
                catch
                {
                    // Log error.
                }
            }
        }

    }
}

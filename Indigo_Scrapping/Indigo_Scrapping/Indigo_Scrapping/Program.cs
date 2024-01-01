using Microsoft.Build.Tasks.Deployment.Bootstrapper;
using Microsoft.VisualBasic.FileIO;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Net.Mail;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
namespace Indigo_Scrapping
{
    class Program
    {

        static DateTime now = DateTime.Today;
        static DateTime dt = DateTime.Now;
        DateTime dateTime = DateTime.UtcNow.Date;
        static string Date1 = DateTime.Today.ToString("dd");
        string Date2 = DateTime.Today.AddDays(-1).ToString("dd MMM ");
        string Month = dt.ToString("MMM");
        DateTime dt1 = DateTime.Now;
        static string Month1 = dt.ToString("MMM");
        static string Year = now.ToString("yyyy");
        // static string BaseURL = @"E:\Amresh Pandey D\All Projects\Indigo_Scrapping\Indigo_Scrapping\" + Year + @"\" + Month1 + @"\" + Date1;
        static string BaseURL = @"E:\All Console Applications\Indigo_Scrapping\Indigo_Scrapping\Indigo_Scrapping\" + Year + @"\" + Month1 + @"\" + Date1;
        // static string Downloadpath = @"E:\Amresh Pandey D\All Projects\Indigo_Scrapping\Indigo_Scrapping\2023\Sep\20";
        static void Main(string[] args)
        {
            //  DataTable a = GetDataTableFromCSVFile();
            //var task = ExecuteScheduler();
            //task.Wait();
            CreateFolder();
            try
            {
                Program program = new Program();
                var Date2 = DateTime.Now.Date;
                DateTime date1 = new DateTime();
                Console.WriteLine(date1.ToString());
                string CreatedOn = "";
                string cs = ConfigurationManager.ConnectionStrings["DSRERP"].ConnectionString;
                SqlConnection con = new SqlConnection(cs);
                string filecheckdate = "select top 1 convert(date,Created_datetime) CreatedOn from LCC_Sales_UAT where convert(date,Created_datetime) = CONVERT(DATE, GETDATE())";
                SqlCommand cmd = new SqlCommand(filecheckdate, con);
                con.Open();
                SqlDataReader dr = cmd.ExecuteReader();
                while (dr.Read())
                {
                    CreatedOn = dr["CreatedOn"].ToString();
                    Console.WriteLine("CREATED ON " + CreatedOn);
                }
                con.Close();
                if (!string.IsNullOrEmpty(CreatedOn) && DateTime.Parse(CreatedOn).Date == Date2)
                {
                    program.FileAlreadyExist();
                }
                else
                {
                    string Email = "";
                    string Password = "";
                    string cs1 = ConfigurationManager.ConnectionStrings["CRED"].ConnectionString;
                    SqlConnection con1 = new SqlConnection(cs1);
                    string getCredential = "select * from TBL_HotfileCred where id = 5 and flag = 1";
                    SqlCommand cmd1 = new SqlCommand(getCredential, con1);
                    con1.Open();
                    SqlDataReader dr1 = cmd1.ExecuteReader();
                    while (dr1.Read())
                    {
                        Email = dr1["email"].ToString();
                        Password = dr1["password"].ToString();
                    }
                    con.Close();
                    var Driver = LoginPage(Email, Password);
                    GetReport(Driver);
                }
            }
            catch (Exception ex)
            {
                ErrorLog(ex);
            }
        }
        public static IWebDriver LoginPage(string email, string password)
        {
            var chromeOptions = new ChromeOptions();
            chromeOptions.AddUserProfilePreference("download.default_directory", BaseURL);
            // chromeOptions.AddUserProfilePreference("download.default_directory", Downloadpath);
            chromeOptions.AddUserProfilePreference("intl.accept_languages", "nl");
            chromeOptions.AddUserProfilePreference("disable-popup-blocking", "true");
            chromeOptions.AddArgument("--disable-blink-features=AutomationControlled");
            chromeOptions.AddUserProfilePreference("credentials_enable_service", "false");
            chromeOptions.AddArguments("start-maximized");
            chromeOptions.AddExcludedArgument("enable-automation");
            chromeOptions.AddExcludedArguments("useAutomationExtension", "false");
            chromeOptions.AddUserProfilePreference("profile.password_manager_enabled", "false");
            chromeOptions.AddArguments("headless=new");
            IWebDriver driver = null;
            try
            {
                driver = new ChromeDriver(chromeOptions);
                driver.Navigate().GoToUrl("https://www.goindigo.in/agent.html?linkNav=partner-login_header");
                Console.WriteLine("enter into login page");
                Thread.Sleep(3000);
                //var username = driver.FindElement(By.XPath("/html/body/main/div[4]/div/div/div[2]/div/div/div/div/div/div/div/div/div/div/div/div/div[2]/div[2]/form/div[1]/div/div[1]/input"));
                var username = driver.FindElement(By.XPath("/html/body/main/div/div/div/div[2]/div/div/div/div[2]/div[2]/div/div/div/div/div/div/div/div/div[2]/div[2]/form/div[1]/div/div[1]/input"));
                username.SendKeys(email);
                //var passwordfield = driver.FindElement(By.XPath("/html/body/main/div[4]/div/div/div[2]/div/div/div/div/div/div/div/div/div/div/div/div/div[2]/div[2]/form/div[2]/div/div[1]/input"));
                var passwordfield = driver.FindElement(By.XPath("/html/body/main/div/div/div/div[2]/div/div/div/div[2]/div[2]/div/div/div/div/div/div/div/div/div[2]/div[2]/form/div[2]/div/div[1]/input"));
                Console.WriteLine(password);
                passwordfield.SendKeys(password);
                //var loginButton = driver.FindElement(By.XPath("/html/body/main/div[4]/div/div/div[2]/div/div/div/div/div/div/div/div/div/div/div/div/div[2]/div[2]/form/div[4]/button"));
                var loginButton = driver.FindElement(By.XPath("/html/body/main/div/div/div/div[2]/div/div/div/div[2]/div[2]/div/div/div/div/div/div/div/div/div[2]/div[2]/form/div[4]/button"));
                loginButton.Click();
                Thread.Sleep(8000);
                var agentPortal = driver.FindElement(By.XPath("/html/body/main/div/div/div/main/div/div/div/div/div/div[1]/div/div/div/section/div/div[1]/a[3]/button"));
                agentPortal.Click();
                Thread.Sleep(30000);
            }
            catch (Exception ex)
            {
                ErrorLog(ex);
            }
            return driver;
        }
        public static void GetReport(IWebDriver driver1)
        {
            driver1.SwitchTo().NewWindow(WindowType.Tab).Navigate().GoToUrl("https://6eagentportal.goindigo.in/agency-dashboard");
            Thread.Sleep(5000);
            try
            {
                var myReport = driver1.FindElement(By.XPath("/html/body/app-root/div/app-header/header/nav/div/div/div[2]/ul/li[5]/a"));
                myReport.Click();
                Thread.Sleep(5000);
            }
            catch (Exception e)
            {
                ErrorLog(e);
            }
            Thread.Sleep(5000);
            var PreviousDate = DateTime.Today.AddDays(-1).ToString("ddMMMyyyy");
            var actualName = "TransactionReport_All_" + PreviousDate + "_" + PreviousDate + "_";
            IWebElement tableElement = driver1.FindElement(By.XPath("/html/body/app-root/div/app-transaction-report/div/form/div/div[2]/div/table"));
            IList<IWebElement> tableRow = tableElement.FindElements(By.TagName("tr"));
            IList<IWebElement> rowTD;
            int i = 1;
            bool isExist = false;
            Thread.Sleep(5000);
            try
            {
                foreach (IWebElement row in tableRow)
                {
                    rowTD = row.FindElements(By.TagName("td"));
                    if (rowTD.Count > 0)
                    {
                        var ReportDate = rowTD[1].Text;
                        var newName = ReportDate.Remove(42, 18);
                        if (newName.Contains(actualName))
                        {
                            var test = "/html/body/app-root/div/app-transaction-report/div/form/div/div[2]/div/table/tbody/tr[" + i + "]/td[5]/button";
                            var Download = driver1.FindElement(By.XPath("/html/body/app-root/div/app-transaction-report/div/form/div/div[2]/div/table/tbody/tr[" + i + "]/td[5]/button"));
                            try
                            {
                                Download.Click();
                                Thread.Sleep(10000);
                            }
                            catch (Exception e)
                            {

                            }
                            try
                            {
                                Download.Click();
                                Thread.Sleep(10000);
                            }
                            catch (Exception e)
                            {

                            }
                            isExist = true;
                        }
                        i++;
                    }
                }

            }
            catch (StaleElementReferenceException s)
            {

            }

            if (isExist)
            {
                DataTable a = GetDataTableFromCSVFile();
                Thread.Sleep(2000);
                driver1.Dispose();
            }
            else
            {
                GenerateReport(driver1);
            }
        }
        public static void GenerateReport(IWebDriver driver2)
        {
            var fromDate = driver2.FindElement(By.Id("DateFrom")).GetAttribute("value");
            var toDate = driver2.FindElement(By.Id("DateTo")).GetAttribute("value");
            var PreviousDate = DateTime.Today.AddDays(-1).ToString("dd MM yyyy");
            try
            {
                if (fromDate == PreviousDate)
                {
                    var generateReport = driver2.FindElement(By.XPath("/html/body/app-root/div/app-transaction-report/div/form/div/div[1]/div[20]/button"));
                    generateReport.Click();
                }
            }
            catch (Exception e)
            {
                ErrorLog(e);
            }
            finally
            {
                driver2.Dispose();
            }

        }
        public static DataTable GetDataTableFromCSVFile()
        {
            DataTable csvData = new DataTable();
            try
            {
                var PreviousDate = DateTime.Today.AddDays(-1).ToString("ddMMMyyyy");
                var upDatedName = "TransactionReport_All_" + PreviousDate + "_" + PreviousDate + "_";
                // var  upDatedName = "TransactionReport_All_19Sep2023_19Sep2023_20230920125314";
                string fileName = @"^" + upDatedName;
                var matches = Directory
                  .GetFiles(BaseURL)
                  .Where(path => System.Text.RegularExpressions.Regex.IsMatch(Path.GetFileName(path), fileName));
                foreach (var item in matches)
                {
                    var Path = item;
                    var Filename = Path.Remove(0, 91);
                    if (Filename.StartsWith(upDatedName))
                    {
                        try
                        {
                            using (TextFieldParser csvReader = new TextFieldParser(/*@"E:\TransactionReport_All_30Dec2023_30Dec2023_20240101110357.csv"*/Path))
                            {
                                csvReader.SetDelimiters(new string[] { "," });
                                csvReader.HasFieldsEnclosedInQuotes = true;
                                string[] colFields = csvReader.ReadFields();
                                Dictionary<string, string> Columnmapping = getDictionary();
                                Dictionary<string, string> Test = new Dictionary<string, string>();
                                foreach (string column in colFields)
                                {
                                    string newCol = "";
                                    DataColumn datecolumn = new DataColumn();
                                    if (Columnmapping.ContainsValue(column))
                                    {
                                        newCol = Columnmapping.FirstOrDefault(x => x.Value == column).Key;
                                        csvData.Columns.Add(newCol);
                                    }
                                    else
                                    {
                                        csvData.Columns.Add(column);
                                    }
                                }
                                while (!csvReader.EndOfData)
                                {
                                    string[] fieldData = csvReader.ReadFields();
                                    //Making empty value as null
                                    for (int i = 0; i < fieldData.Length; i++)
                                    {
                                        if (fieldData[i] == "")
                                        {
                                            fieldData[i] = null;
                                        }
                                    }
                                    csvData.Rows.Add(fieldData);
                                }
                            }
                        }
                        catch (Exception e)
                        {
                            ErrorLog(e);
                        }
                        string[] selectcol = { "Process_date", "Name", "PaymentDtm", "PaymentMethodCode", "PaymentAmount", "PaymentNumber", "BookingDate", "RecordLocator", "IATA_Code", "BookingPromoCode", "ReceivedBy",
                                               "EmailAddress","HomePhone","GDS_recordcode","GDS_recordlocator","GDS_BookingSystemCode","Total_amt","Base_amount" , "SourceAgentCode", "International", "CurrencyCode", "PaxCount", "Name1",
                                               "First_Leg","First_Leg_Dep_Date","Second_Leg","Second_Leg_Dep_Date","Third_Leg","Third_Leg_Dep_Date","Fourth_Leg","Fourth_Leg_Dep_Date"/*,"Fifth_Leg","Fifth_Leg_Dep_Date","Sixth_Leg","Sixth_Leg_Dep_Date"*/
                            };
                        DataTable filterdataatable = SelectColumns(csvData, selectcol);
                        InsertDataIntoSQLServerUsingSQLBulkCopy(filterdataatable);
                    }
                }

            }
            catch (Exception ex)
            {
                return null;
            }
            return csvData;
        }
        static void InsertDataIntoSQLServerUsingSQLBulkCopy(DataTable csvFileData)
        {
            var Date2 = DateTime.Now.Date;
            string CreatedOn = "";
            string cs = ConfigurationManager.ConnectionStrings["DSRERP"].ConnectionString;
            SqlConnection con = new SqlConnection(cs);
            string filecheckdate = "select top 1 convert(date,Created_datetime) CreatedOn from LCC_Sales_UAT where convert(date,Created_datetime) = CONVERT(DATE, GETDATE())";
            SqlCommand cmd = new SqlCommand(filecheckdate, con);
            con.Open();
            SqlDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                CreatedOn = dr["CreatedOn"].ToString();
                Console.WriteLine("CREATED ON " + CreatedOn);
            }
            con.Close();
            if (!string.IsNullOrEmpty(CreatedOn) && DateTime.Parse(CreatedOn).Date == Date2)
            {
            }
            else
            {
                {
                    con.Open();
                    using (SqlBulkCopy s = new SqlBulkCopy(con))
                    {
                        try
                        {
                            s.DestinationTableName = "LCC_Sales_UAT";
                            foreach (var column in csvFileData.Columns)
                                s.ColumnMappings.Add(column.ToString(), column.ToString());
                            s.WriteToServer(csvFileData);

                        }
                        catch (Exception e)
                        {
                            ErrorLog(e);
                            return;
                        }

                    }
                }
                FileDownload();



            }
        }
        static DataTable SelectColumns(DataTable originalTable, string[] columnNames)
        {
            // Create a new DataTable with the selected columns
            DataTable selectedTable = new DataTable();
            try
            {
                foreach (string columnName in columnNames)
                {
                    selectedTable.Columns.Add(columnName, originalTable.Columns[columnName].DataType);
                }
                foreach (DataRow originalRow in originalTable.Rows)
                {
                    DataRow newRow = selectedTable.NewRow();
                    foreach (string columnName in columnNames)
                    {
                        newRow[columnName] = originalRow[columnName];
                    }
                    selectedTable.Rows.Add(newRow);
                }
            }
            catch (Exception e)
            {
            }
            return selectedTable;
        }
        public static Dictionary<string, string> getDictionary()
        {
            Dictionary<string, string> Columnmapping = new Dictionary<string, string>() {
                        { "Process_date" ,"Transaction Date"},
                        //{ "Name" ,"Name"  },
                        //{ "PaymentDtm" ,"PaymentDtm"  },
                        //{ "PaymentMethodCode" ,"PaymentMethodCode"  },
                        //{ "PaymentAmount" ,"PaymentAmount"  },
                        //{ "PaymentNumber" ,"PaymentNumber"  },
                        //{ "BookingDate" ,"BookingDate"  },
                        //{ "RecordLocator" ,"RecordLocator"  },
                        { "IATA_Code" ,"SourceOrganizationCode"  },
                        //{ "BookingPromoCode" ,"BookingPromoCode"  },
                        //{ "ReceivedBy" ,"ReceivedBy"  },
                        //{ "SourceAgentCode" ,"SourceAgentCode"  },
                        //{ "International" ,"International"  },
                        //{ "CurrencyCode" ,"CurrencyCode"  },
                        //{ "PaxCount" ,"PaxCount"  },
                        //{ "Name1" ,"Name1"  },
                        //{ "EmailAddress" ,"EmailAddress"  },
                        //{ "HomePhone" ,"HomePhone"  },
                        //{ "GDS_recordcode" ,"GDS_recordcode"  },
                        //{ "GDS_recordlocator" ,"GDS_recordlocator"  },
                        //{ "GDS_BookingSystemCode" ,"GDS_BookingSystemCode"  },
                        { "Total_amt" ,"Total"  },
                        { "Base_amount" ,"BaseFare"  },
                        { "First_Leg" ,"First Leg"  },
                        { "First_Leg_Dep_Date" ,"First Leg Dep Date"  },
                        { "Second_Leg" ,"Second Leg"  },
                        { "Second_Leg_Dep_Date" ,"Second Leg Dep Date"  },
                        { "Third_Leg" ,"Third Leg"  },
                        { "Third_Leg_Dep_Date" ,"Third Leg Dep Date"  },
                        { "Fourth_Leg" ,"Fourth Leg"  },
                        { "Fourth_Leg_Dep_Date" ,"Fourth Leg Dep Date"  },
                        { "Fifth_Leg" ,"Fifth Leg"  },
                        { "Fifth_Leg_Dep_Date" ,"Fifth Leg Dep Date"  },
                        { "Sixth_Leg" ,"Sixth Leg"  },
                        { "Sixth_Leg_Dep_Date" ,"Sixth Leg Dep Date"  },
                    };
            return Columnmapping;
        }
        public void FileAlreadyExist()
        {
            try
            {
                string FromMail = "no-reply@riya.travel";
                string displayName = "Test";
                string ToEmail = "developers@riya.travel";
                string MailBody = " Dear User, Indigo File Already Exist";
                string subject = "Indigo File Already Downloaded";
                MailAddress sendFrom = new MailAddress(FromMail, displayName);
                MailAddress sendTo = new MailAddress(ToEmail);
                MailMessage nmsg = new MailMessage(sendFrom, sendTo);
                // MailAddress cc = new MailAddress("mahesh.tatkare@riya.travel");
                MailAddress bcc = new MailAddress("amreshpndy123@gmail.com");
                // nmsg.CC.Add(cc);
                // nmsg.Bcc.Add(bcc);
                nmsg.Subject = subject;
                nmsg.IsBodyHtml = true;
                nmsg.Body = MailBody;
                SmtpClient smtpClient = new SmtpClient();
                smtpClient.UseDefaultCredentials = false;
                smtpClient.Host = "smtp.pepipost.com";
                smtpClient.Port = 25;
                string smtpUserName = "riyatravelpepi@riya.travel";
                string smtpPassword = "7b79dc!db222a";
                smtpClient.Credentials = new NetworkCredential(smtpUserName, smtpPassword);
                smtpClient.EnableSsl = false;
                smtpClient.Send(nmsg);   // for local mail disabled
            }
            catch (Exception ex)
            {
            }
        }
        public static void FileDownload()
        {
            try
            {
                string FromMail = "no-reply@riya.travel";
                string displayName = "Indigo";
                string ToEmail = "developers@riya.travel";
                string MailBody = "Dear User ,file Download Successfully";
                string subject = "Indigo file Download successfully ";
                MailAddress sendFrom = new MailAddress(FromMail, displayName);
                MailAddress sendTo = new MailAddress(ToEmail);
                MailMessage nmsg = new MailMessage(sendFrom, sendTo);
                MailAddress cc = new MailAddress("mahesh.tatkare@riya.travel");
                MailAddress cc1 = new MailAddress("mayur.riya04@gmail.com");
                MailAddress bcc = new MailAddress("amreshpndy123@gmail.com");
                nmsg.CC.Add(cc);
                nmsg.CC.Add(cc1);
                nmsg.Bcc.Add(bcc);
                nmsg.Subject = subject;
                nmsg.IsBodyHtml = true;
                nmsg.Body = MailBody;
                SmtpClient smtpClient = new SmtpClient();
                smtpClient.UseDefaultCredentials = false;
                smtpClient.Host = "smtp.pepipost.com";
                smtpClient.Port = 25;
                string smtpUserName = "riyatravelpepi@riya.travel";
                string smtpPassword = "7b79dc!db222a";
                smtpClient.Credentials = new NetworkCredential(smtpUserName, smtpPassword);
                smtpClient.EnableSsl = false;
                smtpClient.Send(nmsg);   // for local mail disabled             
            }
            catch (Exception ex)
            {
            }
            var task = ExecuteScheduler();
            task.Wait();
        }
        public static void ErrorLog(Exception ex)
        {
            try
            {
                string cs = ConfigurationManager.ConnectionStrings["DSRERP"].ConnectionString;
                SqlConnection con1 = new SqlConnection(cs);
                string Addexception = "insert into Indigo_Errortable(errordetails,stacktrace,Message) values(@errordetails,@stacktrace,@Message)";
                SqlCommand cmd1 = new SqlCommand(Addexception, con1);
                cmd1.Parameters.AddWithValue("@errordetails", ex.InnerException != null ? ex.InnerException.ToString() : "");
                cmd1.Parameters.AddWithValue("@stacktrace", ex.StackTrace);
                cmd1.Parameters.AddWithValue("@message", ex.Message);
                con1.Open();
                int i = cmd1.ExecuteNonQuery();
                con1.Close();



                string FromMail = "no-reply@riya.travel";
                string displayName = "Indigo Exception";
                string ToEmail = "mayur.riya04@gmail.com";
                string MailBody = " Dear User, ";                   
                string subject = " Exception in Scheduler for Indigo Scrapping";
                MailAddress sendFrom = new MailAddress(FromMail, displayName);
                MailAddress sendTo = new MailAddress(ToEmail);
                MailMessage nmsg = new MailMessage(sendFrom, sendTo);               
                //MailAddress bcc = new MailAddress("mayur.riya04@gmail.com");            
                //nmsg.Bcc.Add(bcc);
                nmsg.Subject = subject;
                nmsg.IsBodyHtml = true;
                nmsg.Body = MailBody;
                SmtpClient smtpClient = new SmtpClient();
                smtpClient.UseDefaultCredentials = false;
                smtpClient.Host = "smtp.pepipost.com";
                smtpClient.Port = 25;
                string smtpUserName = "riyatravelpepi@riya.travel";
                string smtpPassword = "7b79dc!db222a";
                smtpClient.Credentials = new NetworkCredential(smtpUserName, smtpPassword);
                smtpClient.EnableSsl = false;
                smtpClient.Send(nmsg);   // for local mail disabled


            }
            catch (Exception ex1)
            {

            }
        }

        public static void CreateFolder()
        {
            try
            {
                DateTime now = DateTime.Today;
                DateTime dt = DateTime.Now;
                DateTime dateTime = DateTime.UtcNow.Date;
                var Date1 = DateTime.Today.ToString("dd");
                var Date2 = DateTime.Today.AddDays(-1).ToString("dd MMM ");
                var Month = dt.ToString("MMM");
                DateTime dt1 = DateTime.Now;
                var Month1 = dt.ToString("MMM");
                var Year = now.ToString("yyyy");
                string folderName = BaseURL /*+ Year + @"\" + Month1 + @"\" + Date1*/;
                System.IO.Directory.CreateDirectory(folderName);
            }
            catch (Exception ex)
            {

            }
        }
        public static async Task ExecuteScheduler()
        {
            var PreviousDate = DateTime.Today.AddDays(-1).ToString("yyyy-MM-dd");
            using (var client = new HttpClient())
            {
                client.BaseAddress = new Uri("https://apiskyview.oneriya.com/");
                // client.BaseAddress = new Uri("https://localhost:44304/");
                client.DefaultRequestHeaders.Accept.Clear();
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                //GET Method
                try
                {
                    //HttpResponseMessage response = await client.GetAsync("api/RttSummary/GetLCCDailySalesSummaryDetails?fromDate=2022-04-11&toDate=2022-04-12&isSpecificDateRange=false");

                    HttpResponseMessage response = await client.GetAsync("api/RttSummary/GetLCCDailySalesSummaryDetails?fromDate=" + PreviousDate + "&toDate=" + PreviousDate + "& isSpecificDateRange=false");

                    if (response.IsSuccessStatusCode)
                    {
                        var resrponse = await response.Content.ReadAsStringAsync();
                    }
                    else
                    {
                        Console.WriteLine("Internal server Error");
                    }
                }
                catch (Exception e)
                {

                    throw;
                }

            }
        }
    }
}






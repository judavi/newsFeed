using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Globalization;
using System.Net;
using System.Net.Http;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System.IO;
using System.Net.Mail;

namespace newsFeed
{
    public class Program
    {
        public string errorLog { get; set; }

        static void Main(string[] args)
        {
            int result = Parsing("http://www.fxstreet.com/economic-calendar/");
            if (result == 1)
                Environment.Exit(0); //Exit codes http://www.symantec.com/connect/articles/windows-system-error-codes-exit-codes-description
            else
                //sendMail();
                Environment.Exit(13); //Data invalid 
        }

        private static int Parsing(string website)
        {
            try
            {
                DateTime sunday = Program.StartOfWeek(DateTime.Now, DayOfWeek.Sunday);
                string week = String.Format("{0:MM}", DateTime.Now) + "-" + String.Format("{0:dd}", sunday) + "-" + DateTime.Now.Year;
                string path = @"C:\\News\Calendar-" + week + ".csv";
                Dictionary<string, int> monthValues = new Dictionary<string, int>();
                bool add = false;
                //Initialize month names
                var culture = CultureInfo.GetCultureInfo("en-US");
                var dateTimeInfo = DateTimeFormatInfo.GetInstance(culture);
                int conter = 1;
                foreach (string name in dateTimeInfo.AbbreviatedMonthNames)
                {
                    monthValues.Add(name.ToUpper(), conter);
                    conter++;
                }

                List<newsRow> rowsForNews = new List<newsRow>();
                // Initialize the Chrome Driver
                using (var driver = new ChromeDriver())
                {
                    // Go to the home page
                    driver.Navigate().GoToUrl("http://www.fxstreet.com/economic-calendar/");


                    TimeSpan sec = new TimeSpan(0, 0, 20);
                    WebDriverWait wait = new WebDriverWait(driver, sec);
                    wait.Until(ExpectedConditions.ElementExists(By.Id("closeroadblock")));
                    var closeAdd = driver.FindElement(By.Id("closeroadblock"));
                    closeAdd.Click();
      
                    wait.Until(ExpectedConditions.ElementExists(By.Id("fxst-calendartable")));

                    var menuOptions = driver.FindElement(By.Id("fxst-calendar-filter-dateshortcuts"));
                    foreach (var item in menuOptions.FindElements(By.TagName("a")))
                    {
                        if (item.GetAttribute("data-typefilter") == "thisweek")
                        {
                            item.Click();
                        }
                    }
                    wait.Until(ExpectedConditions.ElementExists(By.Id("fxst-calendartable")));

                    var table = driver.FindElement(By.Id("fxst-calendartable"));

                    var rows = table.FindElements(By.TagName("tr"));

                    DateTime date = new DateTime();
                    string dateInString = "";
                    foreach (var item in rows)
                    {

                        //First view if is the date
                        if (item.GetAttribute("class") == "fxst-dateRow")
                        {
                            var tstr = item.Text.Split(',');
                            var ttstr = tstr[1].Split(new char[0]);
                            int month = monthValues[ttstr[1]];
                            int day = Convert.ToInt32(ttstr[2]);
                            date = new DateTime(2015, month, day);
                            dateInString = tstr[0].Substring(0, 3) + " " + tstr[1];

                        }
                        else if (item.GetAttribute("class") == "fxst-tr-event fxst-evenRow  fxit-eventrow" | item.GetAttribute("class") == "fxst-tr-event fxst-evenRow fxst-tr-nexteventline fxit-eventrow" | item.GetAttribute("class") == "fxst-tr-event fxst-oddRow  fxit-eventrow")
                        {
                            //list of tr
                            add = true;
                            newsRow nRow = new newsRow();
                            nRow.Date = date;
                            nRow.DateInString = dateInString;
                            var listTd = item.FindElements(By.TagName("td"));
                            int conterTd = 0;
                            foreach (var td in listTd)
                            {
                                //related with the position of tags
                                if (conterTd == 0)
                                {
                                    if (!td.Text.Contains("n/a") && !td.Text.Contains("24h"))
                                    {
                                        var strHour = td.Text.Split(':');
                                        nRow.Date = new DateTime(date.Year, date.Month, date.Day, Convert.ToInt32(strHour[0]), Convert.ToInt32(strHour[1]), 0);
                                        nRow.Hour = strHour[0] + ":" + strHour[1];
                                    }
                                    else
                                    {
                                        nRow.Date = new DateTime(date.Year, date.Month, date.Day, 12, 0, 0);
                                        nRow.Hour = "12:00";
                                    }
                                }
                                else if (conterTd == 3)
                                {
                                    nRow.Currency = td.Text;
                                }
                                else if (conterTd == 4)
                                {
                                    var interior = td.FindElement(By.TagName("a"));
                                    nRow.Event = interior.Text;
                                }
                                else if (conterTd == 5)
                                {
                                    var volat = td.FindElement(By.TagName("span"));
                                    if (volat.GetAttribute("title").Contains("High"))
                                    {
                                        nRow.Importance = "High";
                                    }
                                    else if (volat.GetAttribute("title").Contains("Moderate"))
                                    {
                                        nRow.Importance = "Medium";
                                    }
                                    else
                                    {
                                        add = false;
                                    }
                                }
                                else if (conterTd == 10)
                                {
                                    conterTd = 0;
                                    nRow.Forecast = "0";
                                    nRow.Previous = "0";
                                    nRow.Actual = "0";
                                    nRow.TimeZone = "GMT"; //The feed give the source in GMT
                                    if (add && nRow.Currency != "")
                                        rowsForNews.Add(nRow);
                                }

                                conterTd++;
                            }
                        }

                    }

                    //write in file 
                    //before your loop
                    var csv = new StringBuilder();
                    int totalRows = rowsForNews.Count();
                    string headers = "Date,Time,Time Zone,Currency,Event,Importance,Actual,Forecast,Previous";
                    csv.Append(headers + Environment.NewLine);
                    foreach (var news in rowsForNews)
                    {
                        var newLine = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8}{9}", news.DateInString, news.Hour, news.TimeZone, news.Currency, news.Event, news.Importance, totalRows.ToString(), news.Forecast, news.Previous, Environment.NewLine);
                        csv.Append(newLine);
                    }

                    //after your loop
                    File.WriteAllText(path, csv.ToString());

                    //Console.WriteLine("");
                    return 1;
                    //return testFile(rowsForNews,sunday);

                }

            }
            catch (Exception e)
            {
                //writeError(e.InnerException.ToString());
                Console.Write(e.InnerException.ToString());
                return 3;
            }

        }

        /*
        private static int testFile(List<newsRow> news, DateTime firstDateOfWeek)
        {
            //The criteria to be a good file it's have almost one record for each day
            //it's not a good idea... not all days have a notice with medium or high relevance 
            for (int i = 0; i < 5; i++) //6 days
            {
                if (news.Where(x => x.Date.Day == firstDateOfWeek.AddDays(i).Day).Count() == 0)
                {
                    errorLog = "Insufficient records";
                    return 2;
                }
            }

            return 1; //good file 
        }*/

        public static DateTime StartOfWeek(DateTime dt, DayOfWeek startOfWeek)
        {
            int diff = dt.DayOfWeek - startOfWeek;
            if (diff < 0)
            {
                diff += 7;
            }

            return dt.AddDays(-1 * diff).Date;
        }

        public static void writeError(string e)
        {
            string path = @"C:\\News\errorLog-" + DateTime.Now.ToString()+ ".txt";
            File.WriteAllText(path, e);
        }

        public static void sendMail (string e)
        {
            //Write in log and send error! 
            writeError(e);

            MailMessage mail = new MailMessage();
            SmtpClient SmtpServer = new SmtpClient("smtp.gmail.com");

            mail.From = new MailAddress("pmrbot@gmail.com");
            mail.To.Add("email@gmail.com");
            mail.Subject = "Error generating news feed";
            mail.Body = e;

            SmtpServer.Port = 587;
            SmtpServer.Credentials = new System.Net.NetworkCredential("pmrbot@gmail.com", "vlufmxfunovocrwm");
            SmtpServer.EnableSsl = true;

            SmtpServer.Send(mail);
        }
    }

    public class newsRow
    {
        public DateTime Date { get; set; }
        public string DateInString { get; set; }
        public string Hour { get; set; }
        public string TimeZone { get; set; }
        public string Currency { get; set; }
        public string Event { get; set; }
        public string Importance { get; set; }
        public string Actual { get; set; }
        public string Forecast { get; set; }
        public string Previous { get; set; }
    }
}

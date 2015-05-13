using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Security.Cryptography.X509Certificates;
using System.IO;
using System.Text.RegularExpressions;

namespace EWSClientApp
{
    class Program
    {
        static void Main(string[] args)
        {

            try
            {
                Console.WriteLine("Connecting...");
                // Connect to Exchange Web Services
                ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2013);
                service.Credentials = new WebCredentials("emailaddress", "password");
                service.AutodiscoverUrl("emailaddress", RedirectionCallback);
                //find items in the inbox folder 
                FindItemsResults<Item> findResults = service.FindItems(WellKnownFolderName.Inbox, new ItemView(15));
                foreach (Item item in findResults.Items)
                {
                    int subjectName = getEmailSubject(item.Subject);
                    DateTime dateRec = item.DateTimeReceived;
                    string timeStr = dateRec.ToShortDateString();
                    string reconDate = DateTime.Now.AddDays(-1).ToShortDateString();
                    string appendDate = DateTime.Now.AddDays(-1).ToString("yyyyMMdd", System.Globalization.CultureInfo.GetCultureInfo("en-US"));
                    string csv = ".csv";
                    bool contains = csv.IndexOf("csv", StringComparison.OrdinalIgnoreCase) >= 0;
                    //file paths for recon files (eclipse revenue, eclipse avails, novar revenue)
                    string eclRevFP = "C:\\temp\\ReconData\\Eclipse\\Revenue\\";
                    string eclAvaFP = "C:\\temp\\ReconData\\Eclipse\\Avails\\";
                    string novRevFP = "C:\\temp\\ReconData\\Novar\\Revenue\\";

                    item.Load();
                    if (item.HasAttachments && (item.DateTimeReceived.ToShortDateString().Equals(reconDate)))
                    {
                        if (subjectName >= 0 && subjectName <= 6)
                        {
                            foreach (var i in item.Attachments)
                            {
                                FileAttachment fileAttachment = i as FileAttachment;
                                if (Regex.IsMatch(fileAttachment.Name, csv, RegexOptions.IgnoreCase))
                                {
                                    int fileExtPos = fileAttachment.Name.LastIndexOf(".");
                                    switch (subjectName)
                                    {
                                        case 0:
                                            fileAttachment.Load(novRevFP + fileAttachment.Name.Substring(0, fileExtPos) + "_" + appendDate + ".csv");
                                            Console.WriteLine("FileName: " + fileAttachment.Name);
                                            break;
                                        case 1: case 4:
                                            fileAttachment.Load(eclRevFP + fileAttachment.Name.Substring(0, fileExtPos) + "_" + appendDate + ".csv");
                                            Console.WriteLine("FileName: " + fileAttachment.Name);
                                            break;
                                        case 2: case 3: case 5: case 6:
                                            fileAttachment.Load(eclAvaFP + fileAttachment.Name.Substring(0, fileExtPos) + "_" + appendDate + ".csv");
                                            Console.WriteLine("FileName: " + fileAttachment.Name);
                                            break;
                                        default:
                                            Console.WriteLine("Invalid subject");
                                            break;
                                    }
                                }

                            }
                        }
                        else
                        {
                            Console.WriteLine("Invalid selection, please try again");
                        }
                    }
                }
                Console.WriteLine("All emails have been brought down");
                Console.ReadLine();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
                Console.ReadLine();
            }
        }

        static bool RedirectionCallback(string url)
        {
            return url.ToLower().StartsWith("https://");
        }

        public static int getEmailSubject(String subject)
        {
            string novRev = "Revenue Projection";
            string eclRev = "Revenue Analysis Yearly GROSS at 7:30:00pm";
            string eclAvail1 = "Weekly Avails-bravo-not import - 1 (+7 days future)7:30pm Completed";
            string eclAvail2 = "Weekly Avails-amc-import - 1 (+7 days future) 7:30pm Completed";
            string eclRevTue = "Revenue Analysis Yearly GROSS 3:30:01PM";
            string eclAvail1Tue = "Weekly Avails-bravo-not import - 1 (+7 days future)3:31:00pm Completed";
            string eclAvail2Tue = "Weekly Avails-amc-import - 1 (+7 days future) 3:31:00pm Completed";

            //day 2 is tuesday and we need the earlier recon email on that day 
            //use day 3 below because we run the program on wednesday for tuesday recon
            int day = (int)DateTime.Now.DayOfWeek;


            if (subject.Contains(novRev))
                return 0;
            else if (subject.Contains(eclRev) && day != 3)
                return 1;
            else if (subject.Contains(eclAvail1) && day != 3)
                return 2;
            else if (subject.Contains(eclAvail2) && day != 3)
                return 3;
            else if (subject.Contains(eclRevTue) && day == 3)
                return 4;
            else if (subject.Contains(eclAvail1Tue) && day == 3)
                return 5;
            else if (subject.Contains(eclAvail2Tue) && day == 3)
                return 6;
            else
                return -1;
        }
    }
}

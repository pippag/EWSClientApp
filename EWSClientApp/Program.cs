using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Security.Cryptography.X509Certificates;
using System.IO;

namespace EWSClientApp
{
    class Program
    {
        static void Main(string[] args)
        {

            try
            {
                Console.WriteLine("Trying to connect...");
                // Connect to Exchange Web Services
                ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2013);
                service.Credentials = new WebCredentials("pgrundy@decentrix.net", "");
                service.AutodiscoverUrl("pgrundy@decentrix.net", RedirectionCallback);

                FindItemsResults<Item> findResults = service.FindItems(WellKnownFolderName.Inbox, new ItemView(3));
                foreach (Item item in findResults.Items)
                {
                    Console.WriteLine(item.Subject);
                    int subjectName = getEmailSubject(item.Subject);
                    DateTime dateRec = item.DateTimeReceived;
                    string timeStr = dateRec.ToShortDateString(); 
                    Console.WriteLine(dateRec);
                    Console.WriteLine(timeStr);

                    string appendDate = DateTime.Now.AddDays(-1).ToString("yyyyMMdd", System.Globalization.CultureInfo.GetCultureInfo("en-US")); 

                    item.Load();
                    if (item.HasAttachments)
                    {
                        switch (subjectName)
                        {
                            case 0:
                                foreach (var i in item.Attachments)
                                {
                                    FileAttachment fileAttachment = i as FileAttachment;
                                    if (fileAttachment.Name.Contains(".csv"))
                                    {
                                        int fileExtPos = fileAttachment.Name.LastIndexOf(".");
                                        fileAttachment.Load("C:\\temp\\" + fileAttachment.Name.Substring(0, fileExtPos) + "_" + appendDate + ".csv");
                                        Console.WriteLine("FileName: " + fileAttachment.Name);
                                    }

                                }
                                break;
                            case 1:
                                foreach (var i in item.Attachments)
                                {
                                    FileAttachment fileAttachment = i as FileAttachment;
                                    if (fileAttachment.Name.Contains(".csv"))
                                    {
                                        int fileExtPos = fileAttachment.Name.LastIndexOf(".");
                                        fileAttachment.Load("C:\\temp\\" + fileAttachment.Name.Substring(0, fileExtPos) + "_" + appendDate + ".csv");
                                        Console.WriteLine("FileName: " + fileAttachment.Name);
                                    }
                                }
                                break;
                            case 2:
                                foreach (var i in item.Attachments)
                                {
                                    FileAttachment fileAttachment = i as FileAttachment;
                                    if (fileAttachment.Name.Contains(".csv"))
                                    {
                                        int fileExtPos = fileAttachment.Name.LastIndexOf(".");
                                        fileAttachment.Load("C:\\temp\\" + fileAttachment.Name.Substring(0, fileExtPos) + "_" + appendDate + ".csv");
                                        Console.WriteLine("FileName: " + fileAttachment.Name);
                                    }
                                }
                                break;
                            case 3:
                                foreach (var i in item.Attachments)
                                {
                                    FileAttachment fileAttachment = i as FileAttachment;
                                    if (fileAttachment.Name.Contains(".csv"))
                                    {
                                        int fileExtPos = fileAttachment.Name.LastIndexOf(".");
                                        fileAttachment.Load("C:\\temp\\" + fileAttachment.Name.Substring(0, fileExtPos) + "_" + appendDate + ".csv");
                                        Console.WriteLine("FileName: " + fileAttachment.Name);
                                    }
                                }
                                break;
                            case 4:
                                foreach (var i in item.Attachments)
                                {
                                    FileAttachment fileAttachment = i as FileAttachment;
                                    if (fileAttachment.Name.Contains(".csv"))
                                    {
                                        int fileExtPos = fileAttachment.Name.LastIndexOf(".");
                                        fileAttachment.Load("C:\\temp\\" + fileAttachment.Name.Substring(0, fileExtPos) + "_" + appendDate + ".csv");
                                        Console.WriteLine("FileName: " + fileAttachment.Name);
                                    }
                                }
                                break;
                            case 5:
                                foreach (var i in item.Attachments)
                                {
                                    FileAttachment fileAttachment = i as FileAttachment;
                                    if (fileAttachment.Name.Contains(".csv"))
                                    {
                                        int fileExtPos = fileAttachment.Name.LastIndexOf(".");
                                        fileAttachment.Load("C:\\temp\\" + fileAttachment.Name.Substring(0, fileExtPos) + "_" + appendDate + ".csv");
                                        Console.WriteLine("FileName: " + fileAttachment.Name);
                                    }
                                }
                                break;
                            case 6:
                                foreach (var i in item.Attachments)
                                {
                                    FileAttachment fileAttachment = i as FileAttachment;
                                    if (fileAttachment.Name.Contains(".csv"))
                                    {
                                        int fileExtPos = fileAttachment.Name.LastIndexOf(".");
                                        fileAttachment.Load("C:\\temp\\" + fileAttachment.Name.Substring(0, fileExtPos) + "_" + appendDate + ".csv");
                                        Console.WriteLine("FileName: " + fileAttachment.Name);
                                    }
                                }
                                break;
                            default:
                                Console.WriteLine("Invalid selection, please try again");
                                break;

                        }
                    }
                }

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
            //string novRev = "Revenue Projection - Quarterly";
            string novRev = "test email";
            string eclRev = "Revenue Analysis Yearly GROSS at 7:30:00pm";
            string eclAvail1 = "Weekly Avails-bravo-not import - 1 (+7 days future)7:30pm Completed";
            string eclAvail2 = "Weekly Avails-amc-import - 1 (+7 days future) 7:30pm Completed";
            string eclRevTue = "Revenue Analysis Yearly GROSS 3:30:01PM";
            string eclAvail1Tue = "Weekly Avails-bravo-not import - 1 (+7 days future)3:31:00pm Completed";
            string eclAvail2Tue = "Weekly Avails-amc-import - 1 (+7 days future) 3:31:00pm Completed";
            
            //day 2 is tuesday and we need the earlier recon email on that day
            int day = (int)DateTime.Now.DayOfWeek;


            if (subject.Equals(novRev))
                return 0;
            else if (subject.Equals(eclRev) && day != 2)
                return 1;
            else if (subject.Equals(eclAvail1) && day != 2)
                return 2;
            else if (subject.Equals(eclAvail2) && day != 2)
                return 3;
            else if (subject.Equals(eclRevTue) && day == 2)
                return 4;
            else if (subject.Equals(eclAvail1Tue) && day == 2)
                return 5;
            else if (subject.Equals(eclAvail2Tue) && day == 2)
                return 6;
            else
                return -1;
        }
    }
}

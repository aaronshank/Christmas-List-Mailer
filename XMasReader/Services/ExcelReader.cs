using System;
using System.Collections.Generic;
using System.Linq;
using ComponentPro.Excel;
using ComponentPro.Net.Mail;
using ComponentPro.Net;
using System.Web;
using System.IO;
using System.Configuration;
using System.Collections;

namespace XMasReader.Services
{
    public class ExcelReader
    {
        public void UploadExcel()
        {
            List<ExcelData> sendList = new List<ExcelData>();

            // Excel File on desktop location
            string file = @"REDACTED";

            WorkbookManager manager = new WorkbookManager();
            Workbook workbook = manager.Workbooks.Add(file);
            foreach (Worksheet w in workbook.Worksheets)
            {
                Worksheet sheet = w;

                foreach (IRange r in sheet.Rows)
                {
                    if (r.Row >= 2) // Will skip the row with headers
                    {
                        ExcelData exData = new ExcelData();
                        exData.FirstName = r[r.Row, 1].DisplayText;
                        exData.LastName = r[r.Row, 2].DisplayText;
                        // exData.Company = r[r.Row, 3].DisplayText;
                        // exData.Address1 = r[r.Row, 4].DisplayText;
                        // exData.Address2 = r[r.Row, 5].DisplayText;
                        // exData.City = r[r.Row, 6].DisplayText;
                        // exData.State = r[r.Row, 7].DisplayText;
                        // exData.ZipCode = r[r.Row, 8].DisplayText;
                        exData.Email = r[r.Row, 9].DisplayText;
                        exData.From = r[r.Row, 10].DisplayText;

                        sendList.Add(exData);
                    }
                }
            }
            
            // ForEach loop to run through the List<T> of recipients.
            // Null check for error handling.
            // The nulls didn't affect the program because they were read after reading all the pertinent information, but I dislike having errors thrown.
            foreach (ExcelData d in sendList)
            {
                if (d.From != "")
                {
                    SendEmail(d);
                }
            }
        }

        public static void SendEmail(ExcelData data)
        {
            MailMessage mMessage = new MailMessage();
            mMessage.To = data.Email;
            //mMessage.CC = "REDACTED";
            mMessage.From = data.From;
            mMessage.BodyHtml = File.ReadAllText("REDACTED"); // .html file containing our Christmas Card
            mMessage.BodyHtml = mMessage.BodyHtml.Replace("[FirstName]", data.FirstName); // So each one is customized
            mMessage.Subject = "Happy Holidays " + data.FirstName.Trim() + " " + data.LastName.Trim() + "!";
            
            // Do sumtin with this
            // The managerial overlords have decided that nothing is to be done with this
            // data.firstName
            // data.lastName
            // data.company
            // data.address1
            // data.address2
            // data.city
            // data.state
            // data.zipCode

            // Mailer for testing
            // MailMessage testMessage = new MailMessage();
            // testMessage.To = "REDACTED";
            // testMessage.From = "REDACTED";
            // testMessage.Subject = "test";
            // testMessage.BodyHtml = "test";

            Smtp client = new Smtp();
            string mailHost = ConfigurationManager.AppSettings["mailHost"];
            int mailPort = Convert.ToInt32(ConfigurationManager.AppSettings["mailPort"]);
            client.Connect(mailHost, mailPort);
            client.Send(mMessage);
            client.Disconnect();
        }
    }

    public class ExcelData
    {
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string Company { get; set; }
        public string Address1 { get; set; }
        public string Address2 { get; set; }
        public string City { get; set; }
        public string State { get; set; }
        public string ZipCode { get; set; }
        public string Email { get; set; }
        public string From { get; set; }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Administration;  
using Microsoft.SharePoint;  
using System.Net.Mail;  
  
namespace TaskEmailReminderTimerJob
{  
    public class TaskEmailReminderTimerJob : Microsoft.SharePoint.Administration.SPJobDefinition
    {  
        public TaskEmailReminderTimerJob() : base() { }  
        public TaskEmailReminderTimerJob(string jobName, SPService service, SPServer server, SPJobLockType targetType) : base(jobName, service, server, targetType) { }  
        public TaskEmailReminderTimerJob(string jobName, SPWebApplication webApplication) : base(jobName, webApplication, null, SPJobLockType.ContentDatabase)
        {
            this.Title = "Email Notification Job";
        }  
        public override void Execute(Guid contentDbId)
        {
            string from = string.Empty;
            string smtpAddress = string.Empty;
            string to = "michael.dockray@nerconet";
            string subject = "Email for MonkeyPants";
            string body = "<h1> Monkeys Rock !!!! , Email Sending from Testing My Timer Job</h1>";
            SPSecurity.RunWithElevatedPrivileges(delegate () {
                // get a reference to the current site collection's content database   
                SPWebApplication webApplication = this.Parent as SPWebApplication;
                SPContentDatabase contentDb = webApplication.ContentDatabases[contentDbId];
                
                // get a reference to the "Tasks" list in the RootWeb of the first site collection in the content database   
                SPWeb rootWeb = contentDb.Sites[0].RootWeb;
                
                // Get the DB News Announcements List   
                SPList listjob = rootWeb.Lists.TryGetList("Tasks");
                
                // Get sender address from web application settings   
                from = rootWeb.Site.WebApplication.OutboundMailSenderAddress;
                
                // Get SMTP address from web application settings   
                smtpAddress = rootWeb.Site.WebApplication.OutboundMailServiceInstance.Server.Address;
                
                // Send an email if the news is approved   
                bool emailSent = SendMail(smtpAddress, subject, body, true, from, to, null, null);
                
                if (listjob != null && emailSent)
                {
                    SPListItem newListItem = listjob.Items.Add();
                    newListItem["Title"] = string.Concat("Email Notification Sent at : ", DateTime.Now.ToString());
                    newListItem.Update();
                }
            });
            
        }  
        public bool SendMail(string smtpAddress, string subject, string body, bool isBodyHtml, string from, string to, string cc, string bcc)
        {
            bool mailSent = false;
            SmtpClient smtpClient = null;
            
            try
            {
                // Assign SMTP address   
                smtpClient = new SmtpClient();
                smtpClient.Host = smtpAddress;
                
                //Create an email message   
                MailMessage mailMessage = new MailMessage(from, to, subject, body);
                if (!String.IsNullOrEmpty(cc))
                {
                    MailAddress CCAddress = new MailAddress(cc);
                    mailMessage.CC.Add(CCAddress);
                }
                if (!String.IsNullOrEmpty(bcc))
                {
                    MailAddress BCCAddress = new MailAddress(bcc);
                    mailMessage.Bcc.Add(BCCAddress);
                }
                mailMessage.IsBodyHtml = isBodyHtml;
                
                // Send the email   
                smtpClient.Send(mailMessage);
                mailSent = true;
            }
            catch (Exception)
            {
                mailSent = false;
            }

            return mailSent;
         }  
  
  
    }  
}

using Serilog;
using System;
using System.IO;
using System.Net.Mail;
using System.Text;

namespace Test
{
    public class Email
    {
        public static void SendEMail(string emailsubject, string body, string emailreceiver, string emailCC, string filePath)
        {
            Log.Information("SendEMail() has been initiated.");
            MailMessage objMailMessage = new MailMessage();
            SmtpClient objSmtpClient;// = new SmtpClient();
            string strHost = "relay.unitedshore.com";
            int port = 25;
            
            try
            {
                Log.Information("SendEMail() Variables created. Entered Try/Catch");
                objMailMessage.From = new MailAddress("byotlsupport@uwm.com");
                objMailMessage.To.Add(new MailAddress(emailreceiver));
                
                if (!string.IsNullOrWhiteSpace(emailCC))
                {
                    string[] CCaddresses = emailCC.Split(';');
                    foreach (var cc in CCaddresses)
                    {
                        objMailMessage.CC.Add(new MailAddress(cc.Trim()));
                    }
                }

                objMailMessage.Bcc.Add(new MailAddress("aparankussam@uwm.com"));
                objMailMessage.Bcc.Add(new MailAddress("htedjoputranto@uwm.com"));
                objMailMessage.Bcc.Add(new MailAddress("mshaik@uwm.com"));
                objMailMessage.Bcc.Add(new MailAddress("pmallipeddi@uwm.com"));
                objMailMessage.Bcc.Add(new MailAddress("ssaimathi@uwm.com"));
                objMailMessage.Bcc.Add(new MailAddress("egoatley@uwm.com"));
                objMailMessage.Bcc.Add(new MailAddress("rrajendran@uwm.com"));
                objMailMessage.Bcc.Add(new MailAddress("pbattulle@uwm.com"));
                objMailMessage.Subject = emailsubject;
                objMailMessage.Body = body;
                objMailMessage.IsBodyHtml = true;
                Log.Information("SendEMail() Subject, Body & Addresses Added.");

                byte[] data = GetData(filePath);
                Log.Information("GetData() Successeeded.");
                MemoryStream ms = new MemoryStream(data);
                objMailMessage.Attachments.Add(new Attachment(ms, "BirthdayReport", "docx"));
                Log.Information("Document Attached to Email");

                objSmtpClient = new SmtpClient(strHost, port);
                objSmtpClient.EnableSsl = false;
                objSmtpClient.UseDefaultCredentials = false;
                objSmtpClient.Timeout = 10000;
                objSmtpClient.DeliveryMethod = SmtpDeliveryMethod.Network;
                objSmtpClient.Send(objMailMessage);
                Log.Information("SendEMail() Email sent.");
            }
            catch (Exception ex)
            {
                Log.Information("CATCH - SendEMail() has Failed");
                string errorstring = ex.Message.ToString();
                throw ex;
            }
            finally
            {
                Log.Information("FINALLY - SendEMail() has Failed");
                objMailMessage = null;
                objSmtpClient = null;
            }
        }

        static byte[] GetData(string filePath)
        {
            Log.Information("GetData() Called.");
            byte[] data = Encoding.ASCII.GetBytes(filePath);
            return data;
        }
    }
}

using System;
using System.Net.Mail;

namespace Test
{
    class Email
    {
        public void SendEMail(string emailsubject, string body, string emailreceiver, string emailCC)
        {
            MailMessage objMailMessage = new MailMessage();
            SmtpClient objSmtpClient;// = new SmtpClient();
            string strHost = "relay.unitedshore.com";
            int port = 25;
            
            try
            {
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
                objSmtpClient = new SmtpClient(strHost, port);
                objSmtpClient.EnableSsl = false;
                objSmtpClient.UseDefaultCredentials = false;
                objSmtpClient.Timeout = 10000;
                objSmtpClient.DeliveryMethod = SmtpDeliveryMethod.Network;
                objSmtpClient.Send(objMailMessage);
            }
            catch (Exception ex)
            {
                string errorstring = ex.Message.ToString();
                throw ex;
            }
            finally
            {
                objMailMessage = null;
                objSmtpClient = null;
            }
        }
    }
}

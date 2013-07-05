using System;
using System.Data;
using System.Configuration;
using System.Net.Mail;
using System.Net.Configuration;
using System.Net;
using System.Collections.Generic;
using System.ComponentModel;


public class Sendmail
{
    public delegate void CommandActionHandle(object sender, AsyncCompletedEventArgs e);
    public event CommandActionHandle OnCommandAction;
    public Sendmail(string username, string password, string server, string protocol, int port)
    {
        _UserName = username;
        _Password = password;
        _Host = protocol + "." + server;
        _Port = port;
        _Domain = server;
    }
    static string _FromAddress = "";
    static string _FromName = "";
    static string _Host = "smtp.mail.yahoo.com.vn";// "smtp.gmail.com";
    static int _Port = 465;//587;//25
    static string _UserName = "nguyennamthach000";
    static string _Password = "qawsedrf";
    static string _Domain = "mail.yahoo.com.vn";
    public static string FromName
    {
        get { return Sendmail._FromName; }
        set { Sendmail._FromName = value; }
    }
    public static String FormAddress
    {
        get
        {
            //SmtpSection cfg = (SmtpSection)ConfigurationManager.GetSection("system.net/mailSettings/smtp");
            return _UserName;
        }
        set
        {
            _FromAddress = value;
        }
    }
    public string SendMail(string subject, string body,
    Dictionary<string, string> emails, string attFiles, bool isHtml, bool isSSL)
    {

        try
        {
            using (MailMessage mail = new MailMessage())
            {
                mail.From = new MailAddress(FormAddress, _FromName);
                if (emails.Count > 0)
                {
                    string[] arrAtt = attFiles.Split(';');
                    for (int i = 0; i < arrAtt.Length; i++)
                    {
                        Attachment itemAtt = new Attachment(arrAtt[i]);
                        mail.Attachments.Add(itemAtt);
                    }
                    foreach (KeyValuePair<string, string> it in emails)
                    {
                        MailAddress tomail = new MailAddress(it.Key, it.Value);

                        mail.To.Add(tomail);
                    }
                    mail.Subject = subject;
                    mail.Body = body;
                    mail.IsBodyHtml = isHtml;

                    SmtpClient client = new SmtpClient();
                    NetworkCredential Credential = new NetworkCredential(_UserName, _Password);

                    //client.DeliveryMethod = SmtpDeliveryMethod.Network;
                    client.Credentials = Credential;
                    //client.UseDefaultCredentials = false;
                    client.Host = _Host;
                    client.Port = _Port;
                    client.EnableSsl = isSSL;

                    client.Send(mail);
                }
            }
        }
        catch (SmtpException ex)
        {
            return ex.Message;
        }
        return "";
    }

    public string SendMail(string subject, string body,
    string toMail, string toName, string attFiles, bool isHtml, bool isSSL)
    {
        try
        {
            using (MailMessage mail = new MailMessage())
            {
                mail.From = new MailAddress(FormAddress, _FromName);
                MailAddress tomail = new MailAddress(toMail, toName);
                string[] arrAtt = attFiles.Split(';');
                for (int i = 0; i < arrAtt.Length; i++)
                {
                    Attachment itemAtt = new Attachment(arrAtt[i]);
                    mail.Attachments.Add(itemAtt);
                }
                mail.To.Add(tomail);
                mail.Subject = subject;
                mail.SubjectEncoding = System.Text.Encoding.UTF8;
                mail.Body = body;
                mail.IsBodyHtml = isHtml;


                SmtpClient client = new SmtpClient();
                NetworkCredential a = new NetworkCredential(_UserName, _Password);
                client.DeliveryMethod = SmtpDeliveryMethod.Network;
                //client.UseDefaultCredentials = false;
                client.Credentials = a;
                client.Host = _Host;
                client.Port = _Port;
                client.EnableSsl = isSSL;

                client.Send(mail);
            }
        }
        catch (SmtpException ex)
        {
            return ex.Message;
        }
        return "";
    }
}

﻿using S22.Imap;
using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Windows.Forms;

namespace TestGetAndSendEmail
{
    public partial class Form1 : Form
    {
        #region Private Fields

        private string hostname = "xxx.embratelmail.com.br";
        private string password = "xxx";
        private string port = "666";
        private string portToSend = "111";
        private string username = "xxx@gmail.com";

        #endregion Private Fields

        #region Public Constructors

        public Form1()
        {
            InitializeComponent();
        }

        #endregion Public Constructors



        #region Private Methods

        private void btnRun_Click(object sender, EventArgs e)
        {
            GetImapMailAndSend();
            Close();
        }

        private void GetImapMailAndSend()
        {
            using (ImapClient client = new ImapClient(hostname, Convert.ToInt32(port), username, password, AuthMethod.Login, true))
            {
                IEnumerable<uint> uids = client.Search(SearchCondition.Unseen());
                foreach (uint uid in uids)
                {
                    MailMessage mensagem = client.GetMessage(uid);

                    SendMail(mensagem);
                }
            }
        }

        private void SendMail(MailMessage mensagem)
        {
            using (SmtpClient client = new SmtpClient())
            {
                client.Port = Convert.ToInt32(portToSend);
                client.Host = hostname;
                client.EnableSsl = false;
                client.Timeout = 30000;
                client.DeliveryMethod = SmtpDeliveryMethod.Network;
                client.UseDefaultCredentials = false;
                client.Credentials = new NetworkCredential(username, password);

                string novo_assunto = "Encaminhado: " + mensagem.Subject;
                MailMessage nova_mensagem = new MailMessage(username, username, novo_assunto, mensagem.Body);
                nova_mensagem.IsBodyHtml = true;
                nova_mensagem.ReplyToList.Add(new MailAddress(username));
                nova_mensagem.BodyEncoding = UTF8Encoding.UTF8;
                nova_mensagem.DeliveryNotificationOptions = DeliveryNotificationOptions.OnFailure;

                foreach (Attachment anexo in mensagem.Attachments)
                    nova_mensagem.Attachments.Add(anexo);

                client.Send(nova_mensagem);
            }
        }

        #endregion Private Methods
    }
}
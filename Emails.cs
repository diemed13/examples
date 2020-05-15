
/* Generic email class */
/* Developed by Diego Medeiros */
using System;
using System.Web;
using System.Collections.Generic;
using System.Net.Mail;
using System.Diagnostics.CodeAnalysis;
using System.Text.RegularExpressions;
using System.Net;
using System.IO;

namespace Utils.Mail
{
    public class Emails
    {
        delegate void SendAllMessages(List<MailMessage> messages, object state);
        private static Emails currentClass;
        private List<MailMessage> m_MessagesList;

        public List<MailMessage> MessagesList
        {
            get { return m_MessagesList; }
        }

        public Emails()
        {
            m_MessagesList = new List<MailMessage>();
        }

        /// <summary>
        /// Adiciona e-mail em uma lista para serem enviados.
        /// </summary>
        /// <param name="message"></param>
        public static void Add(MailMessage message)
        {
            if (currentClass == null)
                currentClass = new Emails();

            currentClass.m_MessagesList.Add(message);
        }
        public static void Add(MailAddress address, string Subject, string Body)
        {
            MailMessage message = new MailMessage();
            message.To.Add(address);
            message.Subject = Subject;
            message.Body = Body;
            message.IsBodyHtml = true;
            Add(message);
        }

        /// <summary>
        /// Envia os e-mails que foram adicionados na lista ou então os retira da lista.
        /// </summary>
        public static void Flush(bool Send)
        {
            if (currentClass == null)
                return;

            if (Send)
            {
                SendAllMessages s = new SendAllMessages(SendMessages);
                s.BeginInvoke(currentClass.MessagesList, HttpContext.Current, null, null);
            }

            currentClass = null;
        }

        private static void SendMessages(List<MailMessage> messages, object state)
        {
            try
            {
                HttpContext.Current = state as HttpContext;
                SmtpClient client = new SmtpClient();

                foreach (MailMessage message in messages)
                {
                    client.Send(message);
                    GravarLog.Info(string.Format("E-mail para o(s) destinatário(s) '{0}' foi enviado.", message.To.ToString()));
                }
                client = null;
            }
            catch (SmtpException sEx)
            {
                GravarLog.Erro(sEx);
            }
            catch (Exception ex)
            {
                GravarLog.Erro(ex);
            }
        }

        /// <summary>
        /// Adiciona um remetente a uma mensagem de e-mail se o remetente ainda não está adicionado;
        /// </summary>
        /// <param name="message">A mensagem a ser enviada</param>
        /// <param name="Nome">Nome do Remetente</param>
        /// <param name="Email">E-mail do Remetente</param>
        /// <returns>A mensagem com o remetente adicionado</returns>
        public static MailMessage RemetenteAddMessage(string Nome, string Email, MailMessage message)
        {
            try
            {
                if (message == null)
                    return message;

                MailAddress address = RemetenteNew(Nome, Email);
                if (!message.To.Contains(address))
                    message.To.Add(address);

                return message;
            }
            catch (FormatException ex)
            {
                GravarLog.Erro(ex);
                throw;
            }
            catch (ArgumentException ex)
            {
                GravarLog.Erro(ex);
                throw;
            }
        }

        public static MailAddress RemetenteNew(string Nome, string Email)
        {
            string sNome = Nome == null ? string.Empty : Nome;
            return new MailAddress(Email, sNome);
        }

        public static void EnviarEmail(string Para, string Assunto, string Corpo)
        {
            if (string.IsNullOrEmpty(Para)) return;

            MailMessage message = new MailMessage();
            message.To.Add(Para);
            message.Subject = Assunto;
            message.Body = Corpo;
            message.IsBodyHtml = true;

            EnviarEmail(message);
        }

        public static void EnviarEmail(MailMessage email)
        {
            SmtpClient client = new SmtpClient();
            client.Send(email);
            client = null;
        }

        /// <summary>
        /// Envia e-mail com autenticação e Anexo.
        /// </summary>
        /// <param name="De">E-mail de origem. Para SMTP LocaWeb, utilizar o mesmo e-mail do Usuário.</param>
        /// <param name="Para"></param>
        /// <param name="Assunto"></param>
        /// <param name="Corpo"></param>
        /// <param name="Usuario"></param>
        /// <param name="Senha"></param>
        public static void EnviarEmailAutenticado(string De, string Para, string Assunto, string Corpo, string Usuario, string Senha,
            byte[] Arquivo, string NomeArquivo)
        {
            if (string.IsNullOrEmpty(Para)) return;

            MailMessage message = new MailMessage(De, Para, Assunto, Corpo);
            message.IsBodyHtml = true;

            // Arquivo em anexo
            if (Arquivo.Length > 0)
            {
                MemoryStream MS = new MemoryStream(Arquivo);
                Attachment anexo = new Attachment(MS, NomeArquivo);
                message.Attachments.Add(anexo);
            }

            SmtpClient client = new SmtpClient();
            client.Credentials = new NetworkCredential(Usuario, Senha);
            client.Send(message);
            client = null;
        }

        /// <summary>
        /// Envia e-mail com autenticação.
        /// </summary>
        /// <param name="De">E-mail de origem. Para SMTP LocaWeb, utilizar o mesmo e-mail do Usuário.</param>
        /// <param name="Para"></param>
        /// <param name="Assunto"></param>
        /// <param name="Corpo"></param>
        /// <param name="Usuario"></param>
        /// <param name="Senha"></param>
        public static void EnviarEmailAutenticado(string De, string Para, string Assunto, string Corpo, string Usuario, string Senha, int PortaSmtp, string Smtp)
        {
            if (string.IsNullOrEmpty(Para)) return;
            
            MailMessage message = new MailMessage(De, Para, Assunto, Corpo);
            message.IsBodyHtml = true;
            
            SmtpClient client = new SmtpClient();
            client.Credentials = new NetworkCredential(Usuario, Senha);
            client.Host = Smtp;
            client.Port = PortaSmtp;
            client.Send(message);
            client = null;
        }

        public static bool isEmail(string inputEmail)
        {
            if (string.IsNullOrEmpty(inputEmail)) return false;

            string strRegex = @"^([a-zA-Z0-9_\-\.]+)@((\[[0-9]{1,3}" +
                  @"\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([a-zA-Z0-9\-]+\" +
                  @".)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)$";
            
            Regex re = new Regex(strRegex);
            if (re.IsMatch(inputEmail))
                return (true);
            else
                return (false);
        }
    }
}

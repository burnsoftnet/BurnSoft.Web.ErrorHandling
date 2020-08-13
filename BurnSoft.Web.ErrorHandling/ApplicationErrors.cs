using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Configuration;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace BurnSoft.Web.ErrorHandling
{
    public class ApplicationErrors
    {
        
        public static void SendHtmlError(Exception Ex, string to, string from, string smtpServer,string smtpUser, string smtpPwd )
        {
            var strTo = new MailAddress(to);
            var myFrom = new MailAddress(from);
            var Message = new MailMessage(myFrom, strTo);
            Message.IsBodyHtml = true;
            Message.Subject = HttpContext.Current.Request.ServerVariables["HTTP_HOST"] + " - An Application Error Has occured!";
            Message.Body = GetHTMLError(Ex);
            var Client = new SmtpClient(smtpServer);
            if (smtpPwd?.Length > 0)
            {
                Client.Credentials = new NetworkCredential(smtpUser, smtpPwd);
            }

            Client.Send(Message);
        }

        /// <summary>
        /// Returns HTML an formatted error message.
        /// </summary>
        /// <param name="Ex">The ex.</param>
        /// <returns>System.String.</returns>
        public static string GetHTMLError(Exception Ex)
        {
            string Heading;
            string MyHTML;
            var Error_Info = new NameValueCollection();
            Heading = "<TABLE BORDER=\"0\" WIDTH=\"100%\" CELLPADDING=\"1\" CELLSPACING=\"0\"><TR><TD bgcolor=\"RoyalBlue\" COLSPAN=\"2\"><FONT face=\"Arial\" color=\"white\"><B> <!--HEADER--></B></FONT></TD></TR></TABLE>";
            MyHTML = "<FONT face=\"Arial\" size=\"4\" color=\"red\">Error - " + Ex.Message + "</FONT><BR><BR>";
            Error_Info.Add("Message", CleanHTML(Ex.Message));
            HttpContext.Current.Session["ERRORMESSAGE"] = CleanHTML(Ex.Message);
            Error_Info.Add("Source", CleanHTML(Ex.Source));
            Error_Info.Add("TargetSite", CleanHTML(Ex.TargetSite.ToString()));
            Error_Info.Add("StackTrace", CleanHTML(Ex.StackTrace));
            MyHTML += Heading.Replace("<!--HEADER-->", "Error Information");
            MyHTML += CollectionToHtmlTable(Error_Info);
            // // QueryString Collection
            MyHTML += "<BR><BR>" + Heading.Replace("<!--HEADER-->", "QueryString Collection");
            MyHTML += CollectionToHtmlTable(HttpContext.Current.Request.QueryString);
            // // Form Collection
            MyHTML += "<BR><BR>" + Heading.Replace("<!--HEADER-->", "Form Collection");
            MyHTML += CollectionToHtmlTable(HttpContext.Current.Request.Form);
            // // Cookies Collection
            MyHTML += "<BR><BR>" + Heading.Replace("<!--HEADER-->", "Cookies Collection");
            MyHTML += CollectionToHtmlTable(HttpContext.Current.Request.Cookies);
            // // Session Variables
            MyHTML += "<BR><BR>" + Heading.Replace("<!--HEADER-->", "Session Variables");
            MyHTML += CollectionToHtmlTable(HttpContext.Current.Session);
            // // Server Variables
            MyHTML += "<BR><BR>" + Heading.Replace("<!--HEADER-->", "Server Variables");
            MyHTML += CollectionToHtmlTable(HttpContext.Current.Request.ServerVariables);
            return MyHTML;
        }

        /// <summary>
        /// Converts to htmltable.
        /// </summary>
        /// <param name="Collection">The collection.</param>
        /// <returns>System.String.</returns>
        private static string CollectionToHtmlTable(NameValueCollection Collection)
        {
            string TD;
            string MyHTML;
            int i;
            TD = "<TD><FONT face=\"Arial\" size=\"2\"><!--VALUE--></FONT></TD>";
            MyHTML = "<TABLE width=\"100%\">" + " <TR bgcolor=\"#C0C0C0\">" + TD.Replace("<!--VALUE-->", " <B>Name</B>") + " " + TD.Replace("<!--VALUE-->", " <B>Value</B>") + "</TR>";
            // No Body? -> N/A
            if (Collection.Count <= 0)
            {
                Collection = new NameValueCollection();
                Collection.Add("N/A", "");
            }
            else
            {
                // Table Body
                var loopTo = Collection.Count - 1;
                for (i = 0; i <= loopTo; i++)
                    MyHTML += "<TR valign=\"top\" bgcolor=\"#EEEEEE\">" + TD.Replace("<!--VALUE-->", Collection.Keys[i]) + " " + TD.Replace("<!--VALUE-->", Collection[i]) + "</TR> ";
            }
            // Table Footer
            return MyHTML + "</TABLE>";
        }

        /// <summary>
        /// Converts HttpCookieCollection to NameValueCollection
        /// </summary>
        /// <param name="Collection">The collection.</param>
        /// <returns>System.String.</returns>
        private static string CollectionToHtmlTable(HttpCookieCollection Collection)
        {
            var NVC = new NameValueCollection();
            int i;
            string Value;
            try
            {
                if (Collection.Count > 0)
                {
                    var loopTo = Collection.Count - 1;
                    for (i = 0; i <= loopTo; i++)
                        NVC.Add($"{i}", Collection[i].Value);
                }

                Value = CollectionToHtmlTable(NVC);
                return Value;
            }
            catch (Exception MyError)
            {
                return MyError.ToString();
            }
        }

        /// <summary>
        /// Converts HttpSessionState to NameValueCollection
        /// </summary>
        /// <param name="Collection">The collection.</param>
        /// <returns>System.String.</returns>
        private static string CollectionToHtmlTable(System.Web.SessionState.HttpSessionState Collection)
        {
            var NVC = new NameValueCollection();
            int i;
            string Value;
            if (Collection.Count > 0)
            {
                var loopTo = Collection.Count - 1;
                for (i = 0; i <= loopTo; i++)
                    NVC.Add($"{i}", Collection[i].ToString());
            }

            Value = CollectionToHtmlTable(NVC);
            return Value;
        }
        /// <summary>
        /// Cleans the HTML.
        /// </summary>
        /// <param name="HTML">The HTML.</param>
        /// <returns>System.String.</returns>
        private static string CleanHTML(string HTML)
        {
            if (HTML.Length != 0)
            {
                HTML.Replace("<", "<").Replace(@"\r\n", "<BR>").Replace("&", "&").Replace(" ", " ");
            }
            else
            {
                HTML = "";
            }

            return HTML;
        }



    }
}

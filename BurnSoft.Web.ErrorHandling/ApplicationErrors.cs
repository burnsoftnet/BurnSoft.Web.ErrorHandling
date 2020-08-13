using System;
using System.Collections.Specialized;
using System.Net;
using System.Net.Mail;
using System.Web;
// ReSharper disable RedundantAssignment
// ReSharper disable PossibleNullReferenceException
// ReSharper disable ReturnValueOfPureMethodIsNotUsed
// ReSharper disable UnusedMember.Global

namespace BurnSoft.Web.ErrorHandling
{
    /// <summary>
    /// Simple Library that can help send error messages from your website to you or your support email about any application exception errors that have occurred in a pretty HTML report.
    /// Somethings you might need to get the exact exception that occurred as well as any Session information that you can use to recreate or diagnose the issue depending on how you currently use your session variables.
    /// </summary>
    public class ApplicationErrors
    {
        /// <summary>
        /// Send the HTML report of the exception that occured to support or the developer
        /// </summary>
        /// <param name="ex">Exception that you want to send</param>
        /// <param name="to">Send To</param>
        /// <param name="from">Send Form</param>
        /// <param name="smtpServer">SMTP Server</param>
        /// <param name="smtpUser">SMTP User login</param>
        /// <param name="smtpPwd">SMTP User Name password</param>
        /// <param name="port">SMTP Server Port</param>
        /// <example>
        ///  void Application_Error(object sender, EventArgs e) <br/>
        /// { <br/>
        ///    Exception myErr = Server.GetLastError().GetBaseException(); <br/>
        ///    BurnSoft.Web.ErrorHandling.ApplicationErrors.SendHtmlError(myErr,"tosomeone@test.com", "siteSupport@test.com", "smtp.test.com", "siteSupport@test.com", "12345"); <br/>
        /// } <br/>
        /// <br/>
        /// void Session_Start(Object sender, EventArgs e) <br/>
        /// { <br/>
        ///    Session["ERRORMESSAGE"] = ""; <br/>
        /// } <br/>
        /// </example>
        public static void SendHtmlError(Exception ex, string to, string from, string smtpServer,string smtpUser, string smtpPwd, int port = 25 )
        {
            var strTo = new MailAddress(to);
            var myFrom = new MailAddress(from);
            var message = new MailMessage(myFrom, strTo)
            {
                IsBodyHtml = true,
                Subject = GetSubject(),
                Body = GetHtmlError(ex)
            };
            var client = new SmtpClient(smtpServer);
            if (smtpPwd?.Length > 0)
            {
                client.Credentials = new NetworkCredential(smtpUser, smtpPwd);
            }

            client.Port = port;
            client.Send(message);
        }
        /// <summary>
        /// Gets the subject.  This is exposed to the public in case you need to use your own type of smtp server
        /// </summary>
        /// <returns>System.String.</returns>
        public static string GetSubject()
        {
            return $"{HttpContext.Current.Request.ServerVariables["HTTP_HOST"]} - An Application Error Has occured!";
        }

        /// <summary>
        /// Returns HTML an formatted error message.
        /// This is exposed to the public in case you need to use your own type of smtp server
        /// </summary>
        /// <param name="ex">The ex.</param>
        /// <returns>System.String.</returns>
        public static string GetHtmlError(Exception ex)
        {
            var errorInfo = new NameValueCollection();
            var heading = "<TABLE BORDER=\"0\" WIDTH=\"100%\" CELLPADDING=\"1\" CELLSPACING=\"0\"><TR><TD bgcolor=\"RoyalBlue\" COLSPAN=\"2\"><FONT face=\"Arial\" color=\"white\"><B> <!--HEADER--></B></FONT></TD></TR></TABLE>";
            var myHtml = "<FONT face=\"Arial\" size=\"4\" color=\"red\">Error - " + ex.Message + "</FONT><BR><BR>";
            errorInfo.Add("Message", CleanHtml(ex.Message));
            errorInfo.Add("Source", CleanHtml(ex.Source));
            errorInfo.Add("TargetSite", CleanHtml(ex.TargetSite.ToString()));
            errorInfo.Add("StackTrace", CleanHtml(ex.StackTrace));
            myHtml += heading.Replace("<!--HEADER-->", "Error Information");
            myHtml += CollectionToHtmlTable(errorInfo);
            // // QueryString Collection
            myHtml += "<BR><BR>" + heading.Replace("<!--HEADER-->", "QueryString Collection");
            myHtml += CollectionToHtmlTable(HttpContext.Current.Request.QueryString);
            // // Form Collection
            myHtml += "<BR><BR>" + heading.Replace("<!--HEADER-->", "Form Collection");
            myHtml += CollectionToHtmlTable(HttpContext.Current.Request.Form);
            // // Cookies Collection
            myHtml += "<BR><BR>" + heading.Replace("<!--HEADER-->", "Cookies Collection");
            myHtml += CollectionToHtmlTable(HttpContext.Current.Request.Cookies);
            // // Session Variables
            myHtml += "<BR><BR>" + heading.Replace("<!--HEADER-->", "Session Variables");
            myHtml += CollectionToHtmlTable(HttpContext.Current.Session);
            // // Server Variables
            myHtml += "<BR><BR>" + heading.Replace("<!--HEADER-->", "Server Variables");
            myHtml += CollectionToHtmlTable(HttpContext.Current.Request.ServerVariables);
            return myHtml;
        }

        /// <summary>
        /// Converts to html table.
        /// </summary>
        /// <param name="collection">The collection.</param>
        /// <returns>System.String.</returns>
        private static string CollectionToHtmlTable(NameValueCollection collection)
        {
            var td = "<TD><FONT face=\"Arial\" size=\"2\"><!--VALUE--></FONT></TD>";
            var myHtml = "<TABLE width=\"100%\">" + " <TR bgcolor=\"#C0C0C0\">" + td.Replace("<!--VALUE-->", " <B>Name</B>") + " " + td.Replace("<!--VALUE-->", " <B>Value</B>") + "</TR>";
            // No Body? -> N/A
            if (collection.Count <= 0)
            {
                collection = new NameValueCollection {{"N/A", ""}};
            }
            else
            {
                // Table Body
                var loopTo = collection.Count - 1;
                int i;
                for (i = 0; i <= loopTo; i++)
                    myHtml += "<TR valign=\"top\" bgcolor=\"#EEEEEE\">" + td.Replace("<!--VALUE-->", collection.Keys[i]) + " " + td.Replace("<!--VALUE-->", collection[i]) + "</TR> ";
            }
            // Table Footer
            return myHtml + "</TABLE>";
        }

        /// <summary>
        /// Converts HttpCookieCollection to NameValueCollection
        /// </summary>
        /// <param name="collection">The collection.</param>
        /// <returns>System.String.</returns>
        private static string CollectionToHtmlTable(HttpCookieCollection collection)
        {
            var nvc = new NameValueCollection();
            try
            {
                if (collection != null)
                {
                    if (collection.Count > 0)
                    {
                        var loopTo = collection.Count - 1;
                        int i;
                        for (i = 0; i <= loopTo; i++)
                            nvc.Add($"{i}", collection[i].Value);
                    }
                }

                var value = CollectionToHtmlTable(nvc);
                return value;
            }
            catch (Exception myError)
            {
                return myError.ToString();
            }
        }

        /// <summary>
        /// Converts HttpSessionState to NameValueCollection
        /// </summary>
        /// <param name="collection">The collection.</param>
        /// <returns>System.String.</returns>
        private static string CollectionToHtmlTable(System.Web.SessionState.HttpSessionState collection)
        {
            var nvc = new NameValueCollection();
            if (collection != null)
            {
                if (collection.Count > 0)
                {
                    var loopTo = collection.Count - 1;
                    int i;
                    for (i = 0; i <= loopTo; i++)
                        nvc.Add($"{i}", collection[i].ToString());
                }
            }

            var value = CollectionToHtmlTable(nvc);
            return value;
        }
        /// <summary>
        /// Cleans the HTML.
        /// </summary>
        /// <param name="html">The HTML.</param>
        /// <returns>System.String.</returns>
        private static string CleanHtml(string html)
        {
            if (html.Length != 0)
            {
                html.Replace("<", "<").Replace(@"\r\n", "<BR>").Replace("&", "&").Replace(" ", " ");
            }
            else
            {
                html = "";
            }

            return html;
        }



    }
}

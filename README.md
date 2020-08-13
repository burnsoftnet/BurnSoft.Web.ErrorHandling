# BurnSoft.Web.ErrorHandling
Simple Library that can help send error messages from your website to you or your support email about any application exception errors that have occurred in a pretty HTML report.

Somethings you might need to get the exact exception that occurred as well as any Session information that you can use to  recreate or diagnose the issue depending on how you currently use your session variables.

## Example Picture
![](SampleeMail.png)

## How To Use

1. Install nugget package
2. Open your **Global.asax.cs** file.
3. If you do not have it add the *Application_Error* function
4. Call the function *SendHtmlError* in the Application_Error section.
5. Add *Session["ERRORMESSAGE"] = "";* in your **Session_Start** function.

Example:

        void Application_Error(object sender, EventArgs e)
        {
            Exception myErr = Server.GetLastError().GetBaseException();
            BurnSoft.Web.ErrorHandling.ApplicationErrors.SendHtmlError(myErr,"tosomeone@test.com", "siteSupport@test.com", "smtp.test.com", "siteSupport@test.com", "12345");
        }

        void Session_Start(Object sender, EventArgs e)
        {
            Session["ERRORMESSAGE"] = "";
        }
        
So now when you application throws an exception it will catch it and email it to you.
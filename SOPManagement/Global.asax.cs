using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Optimization;
using System.Web.Routing;
using SOPManagement.Models;

namespace SOPManagement
{
    public class MvcApplication : System.Web.HttpApplication
    {
        protected void Application_Start()
        {
            AreaRegistration.RegisterAllAreas();
            FilterConfig.RegisterGlobalFilters(GlobalFilters.Filters);
            RouteConfig.RegisterRoutes(RouteTable.Routes);
            BundleConfig.RegisterBundles(BundleTable.Bundles);

            

        }

        protected void Session_Start(Object sender, EventArgs e)
        {
            //this fires each time a new session is created. for first time then after session exipres 
            //so if we put session value here we cannot chekc it for session time out.
            //you can write code to store session start time of a user in sql table


            //HttpContext.Current.Session["UserFullName"] = Utility.GetLoggedInUserFullName();

        }

        protected void Session_End(Object sender, EventArgs e)
        {
            //you can write code to store session end time of a user in sql table
            //so you can capture how long the user 

            //  Session.Abandon();

            // HttpContext.Application["End"] = "yes";

         //   Application["End"] = "yes";

        }


    }
}

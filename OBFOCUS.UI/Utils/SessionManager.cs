using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using OBFOCUS.Models;

namespace OBFOCUS.UI.Utils
{
    public class SessionManager
    {
        public static UserProfile SessionUserProfile
        {
            get
            {
                if (HttpContext.Current.Session["UserProfile"] != null)
                    return (UserProfile)HttpContext.Current.Session["UserProfile"];
                else
                    return null;
            }
            set
            {
                HttpContext.Current.Session["UserProfile"] = value;
            }
        }
    }
}
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace OBFOCUS.UI.Utils
{
    public class CommonUtil
    {
        public static string GetImagePath(string imgSrc)
        {
            string basePath = "../Content/Images/";
            return basePath + imgSrc;
        }
    }
}
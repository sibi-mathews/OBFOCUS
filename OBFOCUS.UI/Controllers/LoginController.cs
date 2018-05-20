
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using OBFOCUS.UI.Models;

namespace OBFOCUS.UI.Controllers
{
    [ValidateAntiForgeryToken]
    public class LoginController : Controller
    {
        // GET: Login
        public ActionResult Login(LoginModel model)
        {
            return View();
        }
    }
}
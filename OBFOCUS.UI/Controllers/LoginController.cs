using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Threading.Tasks;
using System.Web.Mvc;
using OBFOCUS.UI.ServiceAccessor;
using OBFOCUS.UI.Models;
using System.Web.Security;

namespace OBFOCUS.UI.Controllers
{
    public class LoginController : Controller
    {
        [AllowAnonymous]
        public ActionResult Login()
        {
            return View();
        }

        [HttpPost]
        [AllowAnonymous]
        [ValidateAntiForgeryToken]
        public async Task<ActionResult> Login(LoginModel loginModelObj)
        {
            if (ModelState.IsValid)
            {
                try
                {
                    LoginClient client = new LoginClient();
                    loginModelObj = await client.Authenticate(loginModelObj);
                    if (loginModelObj == null)
                    {
                        ModelState.AddModelError("authentication", "Invalid username/password");
                    }
                    else
                    {
                        FormsAuthentication.SetAuthCookie(loginModelObj.UserName, loginModelObj.RememberMe);
                        return RedirectToAction("Index", "Home");
                    }
                }
                catch (Exception ex)
                {
                    ModelState.AddModelError("authentication", ex.Message);
                }
            }
            
            return View();
        }
    }
}
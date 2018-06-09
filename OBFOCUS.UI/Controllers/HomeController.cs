using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using OBFOCUS.UI.Utils;
using OBFOCUS.UI.Models;
using OBFOCUS.Models;

namespace OBFOCUS.UI.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        [ChildActionOnly]
        public ActionResult Navigation()
        {
            List<NavigationViewModel> navigation = new List<NavigationViewModel>();
            if (SessionManager.SessionUserProfile != null)
            {
                NavigationViewModel navModel = new NavigationViewModel();
                navigation = navModel.LoadTreeView(SessionManager.SessionUserProfile.UserRole.Role);
            }

            return PartialView("_Navigation", navigation);
        }
    }
}
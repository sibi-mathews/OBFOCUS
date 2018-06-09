using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using OBFOCUS.Models;
using OBFOCUS.DO;

namespace OBFOCUS.ServiceAPI
{
    public class LoginController : ApiController
    {
        [HttpPost, ActionName("Authenticate")]
        public HttpResponseMessage Authenticate([FromBody]UserProfile userProfile)
        {
            LoginDo loginDO = new LoginDo();
            HttpResponseMessage response;
            try
            {
                userProfile = loginDO.Authenticate(userProfile);
                if (userProfile.IsAuthenticated)
                {
                    UserRole userRole = loginDO.GetUserRole(userProfile.UserName);
                    if(userRole != null)
                    {
                        userProfile.UserRole = userRole;
                    }
                    response = Request.CreateResponse<UserProfile>(HttpStatusCode.OK, userProfile);
                }
                else
                {
                    response = Request.CreateResponse<UserProfile>(HttpStatusCode.Unauthorized, userProfile);
                }
            }
            catch(Exception ex)
            {
                response = Request.CreateErrorResponse(HttpStatusCode.InternalServerError, ex.Message);
            }
            return response;
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Net.Http;
using System.Text;
using OBFOCUS.UI.Models;
using OBFOCUS.Models;
using Newtonsoft.Json;
using OBFOCUS.UI.Utils;
using System.Threading.Tasks;

namespace OBFOCUS.UI.ServiceAccessor
{
    public class LoginClient
    {
        public async Task<LoginModel> Authenticate(LoginModel loginObj)
        {
            try
            {
                if (loginObj != null)
                {
                    UserProfile userProfile = new UserProfile();
                    userProfile.UserName = loginObj.UserName;
                    userProfile.Password = loginObj.Password;

                    var jsonRequest = JsonConvert.SerializeObject(userProfile);
                    var requestContent = new StringContent(jsonRequest, Encoding.UTF8, "application/json");

                    HttpResponseMessage response = HttpClientService.WebApiClient.PostAsync("api/Login/Authenticate", requestContent).Result;
                    if (response.IsSuccessStatusCode)
                    {
                        string responseStringContent = await response.Content.ReadAsStringAsync();
                        UserProfile result = JsonConvert.DeserializeObject<UserProfile>(responseStringContent);
                        if (result.IsAuthenticated)
                        {
                            userProfile.SessionId = HttpContext.Current.Session.SessionID;
                            SessionManager.SessionUserProfile = userProfile;
                        }
                        else
                        {
                            loginObj = null;
                            throw new Exception("Invalid username/password");
                        }
                    }
                    else
                    {
                        loginObj = null;
                    }
                }
            }
            catch(Exception ex)
            {
                throw ex;
            }
            return loginObj;
        }
    }
}
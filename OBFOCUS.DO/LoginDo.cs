using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OBFOCUS.Models;
using OBFOCUS.DAL;

namespace OBFOCUS.DO
{
    public class LoginDo
    {
        public UserProfile Authenticate(UserProfile userProfile)
        {
            try
            {
                dalLogins login = new dalLogins();
                userProfile = login.Authenticate(ref userProfile);
            }
            catch(Exception ex)
            {
                throw;
            }
            return userProfile;
        }
    }
}

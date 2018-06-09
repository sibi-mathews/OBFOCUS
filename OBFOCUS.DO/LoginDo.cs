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
        /// <summary>
        /// 
        /// </summary>
        /// <param name="userProfile"></param>
        /// <returns></returns>
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


        /// <summary>
        /// 
        /// </summary>
        /// <param name="username"></param>
        /// <returns></returns>
        public UserRole GetUserRole(string username)
        {
            UserRole userRole = null;
            try
            {
                dalLogins login = new dalLogins();
                Globals.UserName = username;
                bool isSuccess = login.GetRole();
                if (isSuccess)
                {
                    userRole = new UserRole();
                    userRole.Role = Globals.UserRole;
                    userRole.LimPhysicianID = Globals.LimPhysicianID;
                    userRole.UserExaminerID = Globals.UserExaminerID;
                }
            }
            catch (Exception ex)
            {
                throw;
            }
            return userRole;
        }
    }
}

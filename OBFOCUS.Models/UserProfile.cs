/******************************************************************************
*
* Name:        UserProfile.cs
*
* Description: Model to hold the user information
*
*-----------------------------------------------------------------------------
*                      CHANGE HISTORY
*   Change No:   Date:          Author:   Description:
*   _________    ___________    ______    ____________________________________
*      
* 
******************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OBFOCUS.Models
{
    public class UserProfile
    {
        private string _userName = string.Empty;
        private string _password = string.Empty;
        private bool _isAuthenticated = false;
        private string _sessionId = string.Empty;
        private UserRole _userRole = null;

        public string UserName
        {
            get { return _userName; }
            set { _userName = value; }
        }

        public string Password
        {
            get { return _password; }
            set { _password = value; }
        }

        public bool IsAuthenticated
        {
            get { return _isAuthenticated; }
            set { _isAuthenticated = value; }
        }

        public string SessionId
        {
            get { return _sessionId; }
            set { _sessionId = value; }
        }

        public UserRole UserRole
        {
            get { return _userRole; }
            set { _userRole = value; }
        }
    }
}

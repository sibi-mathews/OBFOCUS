/******************************************************************************
*
* Name:        UserRole.cs
*
* Description: Model to hold the user Role information
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
    public class UserRole
    {
        private string _role = string.Empty;
        private int? _limPhysicianID = null;
        private int? _userExaminerID = null;

        public string Role
        {
            get { return _role; }
            set { _role = value; }
        }

        public int? LimPhysicianID
        {
            get { return _limPhysicianID; }
            set { _limPhysicianID = value; }
        }

        public int? UserExaminerID
        {
            get { return _userExaminerID; }
            set { _userExaminerID = value; }
        }
    }
}

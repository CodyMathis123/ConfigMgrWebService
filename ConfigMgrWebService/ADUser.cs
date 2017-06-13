using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ConfigMgrWebService
{
    public class ADUser
    {
        public ADUser(string distinguishedName, string displayName, string userName, string email)
        {
            this.distinguishedName = distinguishedName;
            this.displayName = displayName;
            this.userName = userName;
            this.email = email;
        }
        public string distinguishedName { set; get; }
        public string displayName { set; get; }
        public string userName { set; get; }
        public string email { set; get; }

        private ADUser() { }
    }
}

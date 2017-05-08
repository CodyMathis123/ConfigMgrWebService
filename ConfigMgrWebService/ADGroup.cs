using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ConfigMgrWebService
{
    public class ADGroup
    {
        public ADGroup(string name, string description, string distinguishedName)
        {
            this.name = name;
            this.description = description;
            this.distinguishedName = distinguishedName;
        }
        public string name { set; get; }
        public string description { set; get; }
        public string distinguishedName { set; get; }

        private ADGroup() { }
    }
}

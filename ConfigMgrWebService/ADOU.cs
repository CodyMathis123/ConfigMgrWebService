using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ConfigMgrWebService
{
    public class ADOU
    {
        public ADOU(List<string> children, string name, string distinguishedName)
        {
            this.children = children;
            this.name = name;
            this.distinguishedName = distinguishedName;
        }
        public List<string> children { set; get; }
        public string name { set; get; }
        public string distinguishedName { set; get; }

        private ADOU() { }
    }
 
}
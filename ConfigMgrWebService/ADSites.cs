using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ConfigMgrWebService
{
    public class ADSites
    {
        public ADSites(string Name, List<string> domainControllers, List<string> subnets)
        {
            this.Name = Name;
            this.domainControllers = domainControllers;
            this.subnets = subnets;
        }
        public string Name { set; get; }
        public List<string> domainControllers { set; get; }
        public List<string> subnets { set; get; }

        private ADSites() { }
    }
}
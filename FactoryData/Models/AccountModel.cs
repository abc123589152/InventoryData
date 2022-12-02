using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace FactoryData.Models
{
    public class AccountModel
    {
        public string ID { get; set; }
        public string UserName { get; set; }
        public string UserPassWord { get; set; }
        public string Permit { get; set; }
        public string Remarks { get; set; }
        public string AddDatetime { get; set; }
        public string EditDatetime { get; set; }
    }
}
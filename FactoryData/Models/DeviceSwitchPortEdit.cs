using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace FactoryData.Models
{
    public class DeviceSwitchPortEdit
    {
        public string ID { set; get; }
        public string Fab { get; set; }
        public string SwitchName { set; get; }
        public string SwitchIP { set; get; }
        public string SwitchMAC { set; get; }
        public string SwitchPort { set; get; }
        public string DeviceName { set; get; }
        public string DeviceType { set; get; }
        public string RackPeople { set; get; }
        public string Remarks { set; get; }
        public string AddDatetime { set; get; }
        public string EditDatetime { set; get; }
    }
}
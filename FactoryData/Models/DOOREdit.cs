using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace FactoryData.Models
{
    public class DOOREdit
    {
        public string ID { get; set; }
        public string Fab { get; set; }
        public string DoorName { get; set; }
        public string DoorIstarControlName { get; set; }
        public string DoorConnectType { get; set; }
        public string DoorReaderType { get; set; }
        public string DoorAcmNumber { get; set; }
        public string DoorReaderPort { get; set; }
        public string AddVendor { get; set; }
        public string AddDatetime { get; set; }
    }
}
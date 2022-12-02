using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace FactoryData.Models
{
    public class InsertCCTV
    {
        public string ID { get; set; }
        public string Fab { get; set; }
        public string CCTVNumber { set; get; }
        public string CCTVName { get; set; }//CCTV名稱
        public string CCTVIP { get; set; }//CCTV IP位址
        public string CCTVMAC { get; set; }//CCTV MAC位址
        public string CCTVBrand { get; set; }//CCTV品牌
        public string CCTVModel { get; set; }//CCTV型號
        public string CCTVType { get; set; }//CCTV型態
        public string CCTVSwitchIp { get; set; }//CCTV插在的Switch ip位址
        public string CCTVSwitchPort { get; set; }
        public string AddVendor { get; set; }
        public string AddDatetime { get; set; }
        public string EditDatetime { get; set; }
        public string Remarks { get; set; }
    }
}
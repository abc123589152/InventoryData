using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace FactoryData.Models
{
    public class ReportModel
    {
        public string ID { set; get; }
        public string Title { set; get; }//標題
        public string Type { set; get; }//類型
        public string Content { set; get; }//內容
        public string RackPeople { set; get; }//記錄人 
        public string RackVendor { set; get; }//記錄廠商
        public string ImageName { set; get; }//照片名稱
        public string FileName { set; get; }
        public string AddDatetime { set; get; }
        public string EditDatetime { set; get; }
    }
}
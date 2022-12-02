using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace FactoryData.Models
{
    public class PermitionModel
    {
		public string ID{ get; set; }
		public string GroupName{ get; set; }
		public string Device{ get; set; }
		public string Fab{ get; set; }
		public string CCTV{ get; set; }
		public string DOOR{ get; set; }
		public string Report{ get; set; }
		public string OrnerReport { get; set; }
		public string Account{ get; set; }
		public string Permit { get; set; }
		public string SMSCCTV { get; set; }
		public string SMSReader { get; set; }
		public string SMSAlarmSystem { get; set; }
		public string SMSBarriergate { get; set; }
		public string SMSGate { get; set; }
		public string SMSMakeCard { get; set; }
		public string RackPeople { get; set; }
		public string AddDatetime { get; set; }
		public string EditDatetime { get; set; }
	}
}
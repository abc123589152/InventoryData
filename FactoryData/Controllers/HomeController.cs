using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Data;
using System.Data.SqlClient;
using FactoryData.Models;
using System.IO;
using PagedList;
using System.Web.Security;
using System.Net.Http.Headers;
using System.Runtime.Remoting.Messaging;
using Microsoft.Ajax.Utilities;
using System.Security.Cryptography;
using System.Text;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.SS.Util;

namespace FactoryData.Controllers
{
    [Authorize]
    public class HomeController : Controller
    {
        SqlConnection sc = new SqlConnection();
        //string sql = @"server=COSDF15P6VSS71\SQLEXPRESS;initial catalog=FactoryData;user id=sa;password=iamthegad123";//實體server
        //string sql = @"server=COSDF15P6VSS21\SQLEXPRESS;initial catalog=FactoryData;user id=sa;password=iamthegad123";//寫程式的Server
        //string sql = @"server=192.168.1.36;initial catalog=FactoryData;user id=sa;password=iamthegad123";//公司筆電
        string sql = @"server=192.168.1.36;initial catalog=FactoryData;user id=sa;password=iamthegad123";//公司筆電
        string Selectstr = "USE FactoryData select *from DeviceSwitchPort";
        SMSCCTVModel SMSCCTV = new SMSCCTVModel();
        public ActionResult Index(string FirstDate, string LastDate)
        {
            string SelectReportCount = "Select Count(ID) AS CountReport from Report ";
            string SelectCCTVCount = "Select Count(ID) AS CountCCTV from NEWCCTVDATA ";
            string SelectDoorCount = "Select Count(ID) AS CountDoor from NEWDOORDATA ";
            string SelectF15ASMSCCTV = "Select Count(Phase) as F15APhaseCount from SMSCCTV WHERE Phase between 'P1' AND 'P4'";//F15A所有CCTV數量
            string SelectF15BSMSCCTV = "Select Count(Phase) as F15BPhaseCount from SMSCCTV WHERE Phase between 'P5' AND 'P7'";//F15B所有CCTV數量
            string SelectAP05SMSCCTV = "Select Count(Phase) as AP05PhaseCount from SMSCCTV WHERE Phase ='AP5P1'";//AP05所有CCTV數量
            string SelectSMSCCTVALLCount = "Select Count(ID) AS AllCCTV from SMSCCTV";
            string SelectF15ASMSCCTVDownCount = "Select COUNT(Phase) as F15APhaseCount from SMSCCTV WHERE Phase between 'P1' AND 'P4' AND Machinestate = 'Down' AND DownDate between '" + FirstDate + "' AND '"+ LastDate + "'";//F15A CCTVDown數量
            string SelectF15BSMSCCTVDownCount = "Select Count(Phase) AS F15BPhaseCount from SMSCCTV Where  Phase BETWEEN 'P5' AND 'P7' AND Machinestate = 'Down' AND DownDate between '" + FirstDate + "' AND '"+ LastDate + "'";//F15B CCTVDown數量
            string SelectAP05SMSCCTVDownCount = "Select Count(Phase) AS AP05PhaseCount from SMSCCTV Where  Phase like 'AP5P1' AND Machinestate = 'Down'AND DownDate between '" + FirstDate + "' AND '" + LastDate + "'";//FAP05 CCTVDown數量

            DataTable dr = da(SelectReportCount);
            DataTable dc = da(SelectCCTVCount);
            DataTable door = da(SelectDoorCount);
            DataTable F15ADownCCTV = da(SelectF15ASMSCCTVDownCount);
            DataTable F15BDownCCTV = da(SelectF15BSMSCCTVDownCount);
            DataTable AP05DownCCTV = da(SelectAP05SMSCCTVDownCount);
            DataTable F15ACCTV = da(SelectF15ASMSCCTV);
            DataTable F15BCCTV = da(SelectF15BSMSCCTV);
            DataTable AP05CCTV = da(SelectAP05SMSCCTV);
            int F15ACCTVCount = int.Parse(F15ACCTV.Rows[0]["F15APhaseCount"].ToString());
            int F15ACCTVDownCount = int.Parse(F15ADownCCTV.Rows[0]["F15APhaseCount"].ToString());
            int F15BCCTVCount = int.Parse(F15BCCTV.Rows[0]["F15BPhaseCount"].ToString());
            int F15BCCTVDownCount = int.Parse(F15BDownCCTV.Rows[0]["F15BPhaseCount"].ToString());
            int AP05CCTVCount = int.Parse(AP05CCTV.Rows[0]["AP05PhaseCount"].ToString());
            int AP05CCTVDownCount = int.Parse(AP05DownCCTV.Rows[0]["AP05PhaseCount"].ToString());
            ViewData["F15ADownCCTV"] = F15ADownCCTV.Rows[0]["F15APhaseCount"];
            ViewData["F15BDownCCTV"] = F15BDownCCTV.Rows[0]["F15BPhaseCount"];
            ViewData["AP05DownCCTV"] = AP05DownCCTV.Rows[0]["AP05PhaseCount"];
            ViewData["F15ACCTV"] = F15ACCTVCount - F15ACCTVDownCount;//所有F15A CCTV減去已下線的
            ViewData["F15BCCTV"] = F15BCCTVCount - F15BCCTVDownCount;//所有F15B CCTV減去已下線的
            ViewData["AP05CCTV"] = AP05CCTVCount - AP05CCTVDownCount;//所有AP05 CCTV減去已下線的
            ViewData["ReportId"] = dr.Rows[0]["CountReport"];
            ViewData["CCTVId"] = dc.Rows[0]["CountCCTV"];
            ViewData["DOORId"] = door.Rows[0]["CountDoor"];
            string SelectF15ASMSReader = "Select Count(Phase) as F15AReaderPhaseCount from SMSReader WHERE Phase between 'P1' AND 'P4'";//F15A所有CCTV數量
            string SelectF15BSMSReader = "Select Count(Phase) as F15BReaderPhaseCount from SMSReader WHERE Phase between 'P5' AND 'P7'";//F15B所有CCTV數量
            string SelectAP05SMSReader = "Select Count(Phase) as AP05ReaderPhaseCount from SMSReader WHERE Phase ='AP5P1'";//AP05所有CCTV數量
            string SelectSMSReaderALLCount = "Select Count(ID) AS AllCCTV from SMSReader";
            string SelectF15ASMSReaderDownCount = "Select COUNT(Phase) as F15AReaderPhaseCount from SMSReader WHERE Phase between 'P1' AND 'P4' AND Machinestate = 'Down' AND DownDate between '" + FirstDate + "' AND '" + LastDate + "'";//F15A CCTVDown數量
            string SelectF15BSMSReaderDownCount = "Select Count(Phase) AS F15BReaderPhaseCount from SMSReader Where  Phase BETWEEN 'P5' AND 'P7' AND Machinestate = 'Down' AND DownDate between '" + FirstDate + "' AND '" + LastDate + "'";//F15B CCTVDown數量
            string SelectAP05SMSReaderDownCount = "Select Count(Phase) AS AP05ReaderPhaseCount from SMSReader Where  Phase like 'AP5P1' AND Machinestate = 'Down'AND DownDate between '" + FirstDate + "' AND '" + LastDate + "'";//FAP05 CCTVDown數量
            DataTable F15ADownReader = da(SelectF15ASMSReaderDownCount);
            DataTable F15BDownReader = da(SelectF15BSMSReaderDownCount);
            DataTable AP05DownReader = da(SelectAP05SMSReaderDownCount);
            DataTable F15AReader = da(SelectF15ASMSReader);
            DataTable F15BReader = da(SelectF15BSMSReader);
            DataTable AP05Reader = da(SelectAP05SMSReader);
            int F15AReaderCount = int.Parse(F15AReader.Rows[0]["F15AReaderPhaseCount"].ToString());
            int F15AReaderDownCount = int.Parse(F15ADownReader.Rows[0]["F15AReaderPhaseCount"].ToString());
            int F15BReaderCount = int.Parse(F15BReader.Rows[0]["F15BReaderPhaseCount"].ToString());
            int F15BReaderDownCount = int.Parse(F15BDownReader.Rows[0]["F15BReaderPhaseCount"].ToString());
            int AP05ReaderCount = int.Parse(AP05Reader.Rows[0]["AP05ReaderPhaseCount"].ToString());
            int AP05ReaderDownCount = int.Parse(AP05DownReader.Rows[0]["AP05ReaderPhaseCount"].ToString());
            ViewData["F15ADownReader"] = F15ADownReader.Rows[0]["F15AReaderPhaseCount"];
            ViewData["F15BDownReader"] = F15BDownReader.Rows[0]["F15BReaderPhaseCount"];
            ViewData["AP05DownReader"] = AP05DownReader.Rows[0]["AP05ReaderPhaseCount"];
            ViewData["F15AReader"] = F15AReaderCount - F15AReaderDownCount;//所有F15A Reader減去已下線的
            ViewData["F15BReader"] = F15BReaderCount - F15BReaderDownCount;//所有F15B Reader減去已下線的
            ViewData["AP05Reader"] = AP05ReaderCount - AP05ReaderDownCount;//所有AP05 Reader減去已下線的
            string SelectF15ASMSAlarmSystem = "Select Count(Phase) as F15AAlarmPhaseCount from SMSAlarmSystem WHERE Phase between 'P1' AND 'P4'";//F15A所有CCTV數量
            string SelectF15BSMSAlarmSystem = "Select Count(Phase) as F15BAlarmPhaseCount from SMSAlarmSystem WHERE Phase between 'P5' AND 'P7'";//F15B所有CCTV數量
            string SelectAP05SMSAlarmSystem = "Select Count(Phase) as AP05AlarmPhaseCount from SMSAlarmSystem WHERE Phase ='AP5P1'";//AP05所有CCTV數量
            string SelectSMSAlarmSystemALLCount = "Select Count(ID) AS AllCCTV from SMSReader";
            string SelectF15ASMSAlarmSystemDownCount = "Select COUNT(Phase) as F15AAlarmPhaseCount from SMSAlarmSystem WHERE Phase between 'P1' AND 'P4' AND Machinestate = 'Down' AND DownDate between '" + FirstDate + "' AND '" + LastDate + "'";//F15A CCTVDown數量
            string SelectF15BSMSAlarmSystemDownCount = "Select Count(Phase) AS F15BAlarmPhaseCount from SMSAlarmSystem Where  Phase BETWEEN 'P5' AND 'P7' AND Machinestate = 'Down' AND DownDate between '" + FirstDate + "' AND '" + LastDate + "'";//F15B CCTVDown數量
            string SelectAP05SMSAlarmSystemDownCount = "Select Count(Phase) AS AP05AlarmPhaseCount from SMSAlarmSystem Where  Phase like 'AP5P1' AND Machinestate = 'Down'AND DownDate between '" + FirstDate + "' AND '" + LastDate + "'";//FAP05 CCTVDown數量
            DataTable F15ADownAlarm = da(SelectF15ASMSAlarmSystemDownCount);
            DataTable F15BDownAlarm = da(SelectF15BSMSAlarmSystemDownCount);
            DataTable AP05DownAlarm = da(SelectAP05SMSAlarmSystemDownCount);
            DataTable F15AAlarm = da(SelectF15ASMSAlarmSystem);
            DataTable F15BAlarm = da(SelectF15BSMSAlarmSystem);
            DataTable AP05Alarm = da(SelectAP05SMSAlarmSystem);
            int F15AAlarmCount = int.Parse(F15AAlarm.Rows[0]["F15AAlarmPhaseCount"].ToString());
            int F15AAlarmDownCount = int.Parse(F15ADownAlarm.Rows[0]["F15AAlarmPhaseCount"].ToString());
            int F15BAlarmCount = int.Parse(F15BAlarm.Rows[0]["F15BAlarmPhaseCount"].ToString());
            int F15BAlarmDownCount = int.Parse(F15BDownAlarm.Rows[0]["F15BAlarmPhaseCount"].ToString());
            int AP05AlarmCount = int.Parse(AP05Alarm.Rows[0]["AP05AlarmPhaseCount"].ToString());
            int AP05AlarmDownCount = int.Parse(AP05DownAlarm.Rows[0]["AP05AlarmPhaseCount"].ToString());
            ViewData["F15ADownAlarm"] = F15ADownAlarm.Rows[0]["F15AAlarmPhaseCount"];
            ViewData["F15BDownAlarm"] = F15BDownAlarm.Rows[0]["F15BAlarmPhaseCount"];
            ViewData["AP05DownAlarm"] = AP05DownAlarm.Rows[0]["AP05AlarmPhaseCount"];
            ViewData["F15AAlarm"] = F15AAlarmCount - F15AAlarmDownCount;//所有F15A Alarm減去已下線的
            ViewData["F15BAlarm"] = F15BAlarmCount - F15BAlarmDownCount;//所有F15B Alarm減去已下線的
            ViewData["AP05Alarm"] = AP05AlarmCount - AP05AlarmDownCount;//所有AP05 Alarm減去已下線的
            string SelectF15ASMSBarriergate = "Select Count(Phase) as F15ABarriergatePhaseCount from SMSBarriergate WHERE Phase between 'P1' AND 'P4'";//F15A所有CCTV數量
            string SelectF15BSMSBarriergate = "Select Count(Phase) as F15BBarriergatePhaseCount from SMSBarriergate WHERE Phase between 'P5' AND 'P7'";//F15B所有CCTV數量
            string SelectAP05SMSBarriergate = "Select Count(Phase) as AP05BarriergatePhaseCount from SMSBarriergate WHERE Phase ='AP5P1'";//AP05所有CCTV數量
            string SelectSMSBarriergateALLCount = "Select Count(ID) AS AllCCTV from SMSReader";
            string SelectF15ASMSBarriergateDownCount = "Select COUNT(Phase) as F15ABarriergatePhaseCount from SMSBarriergate WHERE Phase between 'P1' AND 'P4' AND Machinestate = 'Down' AND DownDate between '" + FirstDate + "' AND '" + LastDate + "'";//F15A CCTVDown數量
            string SelectF15BSMSBarriergateDownCount = "Select Count(Phase) AS F15BBarriergatePhaseCount from SMSBarriergate Where  Phase BETWEEN 'P5' AND 'P7' AND Machinestate = 'Down' AND DownDate between '" + FirstDate + "' AND '" + LastDate + "'";//F15B CCTVDown數量
            string SelectAP05SMSBarriergateDownCount = "Select Count(Phase) AS AP05BarriergatePhaseCount from SMSBarriergate Where  Phase like 'AP5P1' AND Machinestate = 'Down'AND DownDate between '" + FirstDate + "' AND '" + LastDate + "'";//FAP05 CCTVDown數量
            DataTable F15ADownBarriergate = da(SelectF15ASMSBarriergateDownCount);
            DataTable F15BDownBarriergate = da(SelectF15BSMSBarriergateDownCount);
            DataTable AP05DownBarriergate = da(SelectAP05SMSBarriergateDownCount);
            DataTable F15ABarriergate = da(SelectF15ASMSBarriergate);
            DataTable F15BBarriergate = da(SelectF15BSMSBarriergate);
            DataTable AP05Barriergate = da(SelectAP05SMSBarriergate);
            int F15ABarriergateCount = int.Parse(F15ABarriergate.Rows[0]["F15ABarriergatePhaseCount"].ToString());
            int F15ABarriergateDownCount = int.Parse(F15ADownBarriergate.Rows[0]["F15ABarriergatePhaseCount"].ToString());
            int F15BBarriergateCount = int.Parse(F15BBarriergate.Rows[0]["F15BBarriergatePhaseCount"].ToString());
            int F15BBarriergateDownCount = int.Parse(F15BDownBarriergate.Rows[0]["F15BBarriergatePhaseCount"].ToString());
            int AP05BarriergateCount = int.Parse(AP05Barriergate.Rows[0]["AP05BarriergatePhaseCount"].ToString());
            int AP05BarriergateDownCount = int.Parse(AP05DownBarriergate.Rows[0]["AP05BarriergatePhaseCount"].ToString());
            ViewData["F15ADownBarriergate"] = F15ADownBarriergate.Rows[0]["F15ABarriergatePhaseCount"];
            ViewData["F15BDownBarriergate"] = F15BDownBarriergate.Rows[0]["F15BBarriergatePhaseCount"];
            ViewData["AP05DownBarriergate"] = AP05DownBarriergate.Rows[0]["AP05BarriergatePhaseCount"];
            ViewData["F15ABarriergate"] = F15ABarriergateCount - F15ABarriergateDownCount;//所有F15A 柵欄機減去已下線的
            ViewData["F15BBarriergate"] = F15BBarriergateCount - F15BBarriergateDownCount;//所有F15B 柵欄機減去已下線的
            ViewData["AP05Barriergate"] = AP05BarriergateCount - AP05BarriergateDownCount;//所有AP05 柵欄機減去已下線的
            string SelectF15ASMSGate = "Select Count(Phase) as F15AGatePhaseCount from SMSGate WHERE Phase between 'P1' AND 'P4'";//F15A所有CCTV數量
            string SelectF15BSMSGate = "Select Count(Phase) as F15BGatePhaseCount from SMSGate WHERE Phase between 'P5' AND 'P7'";//F15B所有CCTV數量
            string SelectAP05SMSGate = "Select Count(Phase) as AP05GatePhaseCount from SMSGate WHERE Phase ='AP5P1'";//AP05所有CCTV數量
            string SelectSMSGateALLCount = "Select Count(ID) AS AllCCTV from SMSReader";
            string SelectF15ASMSGateDownCount = "Select COUNT(Phase) as F15AGatePhaseCount from SMSGate WHERE Phase between 'P1' AND 'P4' AND Machinestate = 'Down' AND DownDate between '" + FirstDate + "' AND '" + LastDate + "'";//F15A CCTVDown數量
            string SelectF15BSMSGateDownCount = "Select Count(Phase) AS F15BGatePhaseCount from SMSGate Where  Phase BETWEEN 'P5' AND 'P7' AND Machinestate = 'Down' AND DownDate between '" + FirstDate + "' AND '" + LastDate + "'";//F15B CCTVDown數量
            string SelectAP05SMSGateDownCount = "Select Count(Phase) AS AP05GatePhaseCount from SMSGate Where  Phase like 'AP5P1' AND Machinestate = 'Down'AND DownDate between '" + FirstDate + "' AND '" + LastDate + "'";//FAP05 CCTVDown數量
            DataTable F15ADownGate = da(SelectF15ASMSGateDownCount);
            DataTable F15BDownGate = da(SelectF15BSMSGateDownCount);
            DataTable AP05DownGate = da(SelectAP05SMSGateDownCount);
            DataTable F15AGate = da(SelectF15ASMSGate);
            DataTable F15BGate = da(SelectF15BSMSGate);
            DataTable AP05Gate = da(SelectAP05SMSGate);
            int F15AGateCount = int.Parse(F15AGate.Rows[0]["F15AGatePhaseCount"].ToString());
            int F15AGateDownCount = int.Parse(F15ADownGate.Rows[0]["F15AGatePhaseCount"].ToString());
            int F15BGateCount = int.Parse(F15BGate.Rows[0]["F15BGatePhaseCount"].ToString());
            int F15BGateDownCount = int.Parse(F15BDownGate.Rows[0]["F15BGatePhaseCount"].ToString());
            int AP05GateCount = int.Parse(AP05Gate.Rows[0]["AP05GatePhaseCount"].ToString());
            int AP05GateDownCount = int.Parse(AP05DownGate.Rows[0]["AP05GatePhaseCount"].ToString());
            ViewData["F15ADownGate"] = F15ADownGate.Rows[0]["F15AGatePhaseCount"];
            ViewData["F15BDownGate"] = F15BDownGate.Rows[0]["F15BGatePhaseCount"];
            ViewData["AP05DownGate"] = AP05DownGate.Rows[0]["AP05GatePhaseCount"];
            ViewData["F15AGate"] = F15AGateCount - F15AGateDownCount;//所有F15A 通關機減去已下線的
            ViewData["F15BGate"] = F15BGateCount - F15BGateDownCount;//所有F15B 通關機減去已下線的
            ViewData["AP05Gate"] = AP05GateCount - AP05GateDownCount;//所有AP05 通關機減去已下線的
            string SelectF15ASMSMakeCard = "Select Count(Phase) as F15AMakeCardPhaseCount from SMSMakeCard WHERE Phase between 'P1' AND 'P4'";//F15A所有CCTV數量
            string SelectF15BSMSMakeCard = "Select Count(Phase) as F15BMakeCardPhaseCount from SMSMakeCard WHERE Phase between 'P5' AND 'P7'";//F15B所有CCTV數量
            string SelectAP05SMSMakeCard = "Select Count(Phase) as AP05MakeCardPhaseCount from SMSMakeCard WHERE Phase ='AP5P1'";//AP05所有CCTV數量
            string SelectSMSMakeCardALLCount = "Select Count(ID) AS AllCCTV from SMSReader";
            string SelectF15ASMSMakeCardDownCount = "Select COUNT(Phase) as F15AMakeCardPhaseCount from SMSMakeCard WHERE Phase between 'P1' AND 'P4' AND Machinestate = 'Down' AND DownDate between '" + FirstDate + "' AND '" + LastDate + "'";//F15A CCTVDown數量
            string SelectF15BSMSMakeCardDownCount = "Select Count(Phase) AS F15BMakeCardPhaseCount from SMSMakeCard Where  Phase BETWEEN 'P5' AND 'P7' AND Machinestate = 'Down' AND DownDate between '" + FirstDate + "' AND '" + LastDate + "'";//F15B CCTVDown數量
            string SelectAP05SMSMakeCardDownCount = "Select Count(Phase) AS AP05MakeCardPhaseCount from SMSMakeCard Where  Phase like 'AP5P1' AND Machinestate = 'Down'AND DownDate between '" + FirstDate + "' AND '" + LastDate + "'";//FAP05 CCTVDown數量
            DataTable F15ADownMakeCard = da(SelectF15ASMSMakeCardDownCount);
            DataTable F15BDownMakeCard = da(SelectF15BSMSMakeCardDownCount);
            DataTable AP05DownMakeCard = da(SelectAP05SMSMakeCardDownCount);
            DataTable F15AMakeCard = da(SelectF15ASMSMakeCard);
            DataTable F15BMakeCard = da(SelectF15BSMSMakeCard);
            DataTable AP05MakeCard = da(SelectAP05SMSMakeCard);
            int F15AMakeCardCount = int.Parse(F15AMakeCard.Rows[0]["F15AMakeCardPhaseCount"].ToString());
            int F15AMakeCardeDownCount = int.Parse(F15ADownMakeCard.Rows[0]["F15AMakeCardPhaseCount"].ToString());
            int F15BMakeCardCount = int.Parse(F15BMakeCard.Rows[0]["F15BMakeCardPhaseCount"].ToString());
            int F15BMakeCardDownCount = int.Parse(F15BDownMakeCard.Rows[0]["F15BMakeCardPhaseCount"].ToString());
            int AP05MakeCardCount = int.Parse(AP05MakeCard.Rows[0]["AP05MakeCardPhaseCount"].ToString());
            int AP05MakeCardDownCount = int.Parse(AP05DownMakeCard.Rows[0]["AP05MakeCardPhaseCount"].ToString());
            ViewData["F15ADownMakeCard"] = F15ADownMakeCard.Rows[0]["F15AMakeCardPhaseCount"];
            ViewData["F15BDownMakeCard"] = F15BDownMakeCard.Rows[0]["F15BMakeCardPhaseCount"];
            ViewData["AP05DownMakeCard"] = AP05DownMakeCard.Rows[0]["AP05MakeCardPhaseCount"];
            ViewData["F15AMakeCard"] = F15AMakeCardCount - F15AMakeCardeDownCount;//所有F15A 製證機減去已下線的
            ViewData["F15BMakeCard"] = F15BMakeCardCount - F15BMakeCardDownCount;//所有F15B 製證機減去已下線的
            ViewData["AP05MakeCard"] = AP05MakeCardCount - AP05MakeCardDownCount;//所有AP05 製證機減去已下線的
            return View();
        }       
        [HttpPost]
        public ActionResult InsertFab(FabModel insfab)
        {
            string InsertFab = string.Format("insert into Fab(FabName)VALUES('" + insfab.FabName + "')");
            CommandQuerytosql(InsertFab);
            return RedirectToAction("Fab");
        }
        public ActionResult InsertFab()
        {
            return View();
        }       
        [HttpPost]
        public ActionResult EditFab(FabModel EdFab) 
        {
            
            string EditFab = string.Format("update Fab set FabName = '"+EdFab.FabName+"' Where ID ="+EdFab.ID);
            CommandQuerytosql(EditFab);
            return RedirectToAction("Fab");
        }
        public ActionResult EditFab(FabModel EdFab,string id)
        {
            id = EdFab.ID;
            
            List<FabModel> list = new List<FabModel>();
            string SelectFab= string.Format("Select *from Fab Where ID = "+id);
            DataTable dt = da(SelectFab);
            for (int i = 0; i < dt.Rows.Count; i++) 
            {
                FabModel fab = new FabModel();
                fab.ID = dt.Rows[i]["ID"].ToString();
                fab.FabName = dt.Rows[i]["FabName"].ToString();
                list.Add(fab);
            }
            return View(list);
        }
        public ActionResult FabDelete(FabModel DeFab)
        {            
            string DeleteFab = string.Format("Delete from Fab WHERE ID =" + DeFab.ID.Replace("'", "''"));
            CommandQuerytosql(DeleteFab);
            return RedirectToAction("Fab");
        }
        public ActionResult Fab(string searchFab)
        {
            List<FabModel> list = new List<FabModel>();
            string SelectFab = "";
            if (!string.IsNullOrEmpty(searchFab))
            {
                SelectFab += string.Format("Select *from Fab Where FabName like '" + SelectFab + "'");
            }
            else
            {
                SelectFab += string.Format("Select *from Fab");
            }
            ReturnPermit ret = new ReturnPermit();
            string FAB = ret.returnPermitFab(Session["Permit"].ToString(),sql);
            if (FAB == "授權")
            {
                DataTable dt = da(SelectFab);
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    FabModel fab = new FabModel();
                    fab.ID = dt.Rows[i]["ID"].ToString();
                    fab.FabName = dt.Rows[i]["FabName"].ToString();
                    list.Add(fab);
                }
                return View(list);
            }
            else 
            {
                return View("NoAuthorityView");
            }
        }
        public ActionResult InserAccount(PermitionModel per,ReturnPermit ret)
        {
            string selecPer = "Select *from Permition";
            DataTable dt = da(selecPer);
            string Account = ret.returnPermitAccount(Session["Permit"].ToString(),sql);
            if (Account == "授權")
            {
                return View(dt);
            }
            else
            {
                return RedirectToAction("NoAuthorityView");
            }
        }
        [HttpPost]
        public ActionResult InserAccount(AccountModel acou)
        {
            SHA1CryptoServiceProvider sha1 = new SHA1CryptoServiceProvider();
            byte[] password_bytes = Encoding.ASCII.GetBytes(acou.UserPassWord);
            byte[] encrypt_bytes = sha1.ComputeHash(password_bytes);
            string loginpass = Convert.ToBase64String(encrypt_bytes);
            //string erypass = acou.UserPassWord.ToString();
            //string inserStr = "insert into Account(UserName,UserPassWord,AddDatetime,Remarks)" +
            //    "VALUES('" + acou.UserName + "','" + acou.UserPassWord + "','" + acou.AddDatetime + "','" + acou.Remarks + "')";
            string ins = string.Format("insert into Account(UserName, UserPassWord,Permit,AddDatetime, Remarks)" +
                "VALUES('" + acou.UserName + "','" + acou.UserPassWord + "','"+acou.Permit + "','" + acou.AddDatetime + "','" + acou.Remarks + "')");
            CommandQuerytosql(ins);
            return RedirectToAction("AccountDetail");
        }
        public ActionResult AccountDelete(AccountModel acou)
        {
            string Deletestr = string.Format("Delete from Account WHERE ID =" + acou.ID.Replace("'", "''"));
            CommandQuerytosql(Deletestr);
            DataTable dt = da(Selectstr);
            return RedirectToAction("AccountDetail");

        }
        [HttpPost]
        public ActionResult EditPassword(Account account)
        {

            //FactoryDataEntities1 db = new FactoryDataEntities1();
            //account.UserName = Session["UserName"].ToString();
            //var name = db.Account.Where(x => x.UserName == account.UserName && x.UserPassWord == Password).FirstOrDefault();

            string SqlEdit = string.Format("UPDATE Account SET UserPassWord = '" + account.UserPassWord + "' WHERE UserName = '" + Session["UserName"] + "'");
            CommandQuerytosql(SqlEdit);
            return RedirectToAction("Login", "Login");
        }
        public ActionResult AccountDetail(string searchName, int? page, int? pageSize)
        {
            List<AccountModel> list = new List<AccountModel>();
            ViewBag.PageSize = new List<SelectListItem>()
            {
                new SelectListItem() { Value="10", Text= "10" },
                new SelectListItem() { Value="30", Text= "30" },
                new SelectListItem() { Value="50", Text= "50" },
                new SelectListItem() { Value="100", Text= "100" },
                new SelectListItem() { Value="200", Text= "200" },
            };
            int pageNumber = (page ?? 1);
            int pagesize = (pageSize ?? 10);
            ViewBag.psize = pagesize;
            string strsql = "";
            if (!string.IsNullOrEmpty(searchName))
            {
                strsql += "Select *from Account where UserName ='%" + searchName + "%'";
            }
            else
            {
                strsql += "Select *from Account order by ID DESC";
            }
            ReturnPermit ret = new ReturnPermit();
            string Account = ret.returnPermitAccount(Session["Permit"].ToString(), sql);
            if(Account == "授權") 
            { 
            DataTable dt = da(strsql);
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                AccountModel insd = new AccountModel();
                insd.ID = dt.Rows[i]["ID"].ToString();
                insd.UserName = dt.Rows[i]["UserName"].ToString();
                insd.Permit = dt.Rows[i]["Permit"].ToString();
                insd.Remarks = dt.Rows[i]["Remarks"].ToString();
                insd.AddDatetime = dt.Rows[i]["AddDatetime"].ToString();
                list.Add(insd);
            }
                var pagelist = list.ToPagedList(pageNumber, pagesize);       
                return View(pagelist);
            }
            else
            {
                return RedirectToAction("NoAuthorityView");
            }
            //return View(pagelist);
        }
        public ActionResult EditPassword()
        {
            return View();
        }
        [HttpPost]
        public ActionResult AccountEdit(AccountModel account)
        {
            //FactoryDataEntities1 db = new FactoryDataEntities1();
            //account.UserName = Session["UserName"].ToString();
            //var name = db.Account.Where(x => x.UserName == account.UserName && x.UserPassWord == Password).FirstOrDefault();           
            string SqlEdit = string.Format("UPDATE Account SET Permit ='" + account.Permit + "',Remarks ='" + account.Remarks + "',AddDatetime ='" + account.AddDatetime + "',EditDatetime = GETDATE() WHERE ID =" + account.ID);
            CommandQuerytosql(SqlEdit);
            return RedirectToAction("AccountDetail");
        }
        public ActionResult AccountEdit(AccountModel account, string id)
        {
            id = account.ID;
            string selectAccount = "Select *from Account where ID = " + id;
            DataTable dt = da(selectAccount);
            return View(dt);
        }

        public ActionResult Door(string searchDoorNumber, int? page, int? pageSize)
        {
            List<InsertDoor> list = new List<InsertDoor>();
            string strDOOR = "";
            ViewBag.pageSize = new List<SelectListItem>()
            {
                new SelectListItem() { Value="10", Text= "10" },
                new SelectListItem() { Value="30", Text= "30" },
                new SelectListItem() { Value="50", Text= "50" },
                new SelectListItem() { Value="100", Text= "100" },
                new SelectListItem() { Value="200", Text= "200" },
            };
            int pageNumber = (page ?? 1);
            int pagesize = (pageSize ?? 5);
            ViewBag.psize = pagesize;
            if (!string.IsNullOrEmpty(searchDoorNumber))
            {
                strDOOR = "Select *from NEWDOORDATA WHERE DOORNumber like '%" + searchDoorNumber + "%'";
            }
            else
            {
                strDOOR = "select *from NEWDOORDATA order by ID DESC";
            }
            ReturnPermit ret = new ReturnPermit();
            string DoorRe = ret.returnDoorper(Session["Permit"].ToString(), sql);
            if (DoorRe == "授權")
            {
                DataTable dt = da(strDOOR);
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    InsertDoor insd = new InsertDoor();
                    insd.ID = dt.Rows[i]["ID"].ToString();
                    insd.Fab = dt.Rows[i]["Fab"].ToString();
                    insd.DoorNumber = dt.Rows[i]["DoorNumber"].ToString();
                    insd.DoorName = dt.Rows[i]["DoorName"].ToString();
                    insd.DoorIstarControlName = dt.Rows[i]["DoorIstarControlName"].ToString();
                    insd.DoorConnectType = dt.Rows[i]["DoorConnectType"].ToString();
                    insd.DoorReaderType = dt.Rows[i]["DoorReaderType"].ToString();
                    insd.DoorAcmNumber = dt.Rows[i]["DoorAcmNumber"].ToString();
                    insd.DoorReaderPort = dt.Rows[i]["DoorReaderPort"].ToString();
                    insd.AddVendor = dt.Rows[i]["AddVendor"].ToString();
                    insd.AddDatetime = dt.Rows[i]["AddDatetime"].ToString();
                    insd.EditDatetime = dt.Rows[i]["EditDatetime"].ToString();
                    insd.Remarks = dt.Rows[i]["Remarks"].ToString();
                    list.Add(insd);
                }
                var pagelist = list.ToPagedList(pageNumber, pagesize);
                return View(pagelist);
            }
            else
            {
                return View("NoAuthorityView");
            }
        }
        [HttpPost]
        public ActionResult InsertDOOR(InsertDoor insd)
        {
            sc.ConnectionString = sql;
            string insertStr = "insert into NEWDOORDATA(Fab,DoorNumber,DoorName,DoorIstarControlName,DoorConnectType,DoorReaderType,DoorAcmNumber,DoorReaderPort,AddVendor,AddDatetime,Remarks)" +
                "VALUES('"+insd.Fab+"','" + insd.DoorNumber + "','" + insd.DoorName + "','" + insd.DoorIstarControlName + "','" + insd.DoorConnectType + "','" + insd.DoorReaderType + "','" + insd.DoorAcmNumber + "','"
                + insd.DoorReaderPort + "','" + insd.AddVendor + "','" + insd.AddDatetime + "','" + insd.Remarks + "')";
            //string.Format("insert into DeviceSwitchPort(ID,SwitchName,SwitchPort,DeviceType,AddDatetime,EditDatetime" +
            //")VALUES(" + ins.ID + "," + ins.SwitchName.Replace("'","''") + "," + ins.SwitchPort.Replace("'","''") + "," + ins.DeviceType.Replace("'","''")
            //+ "," + ins.AddDatetime.Replace("'","''") + "," + ins.EditDatetime.Replace("'","''") + ")"
            //);
            CommandQuerytosql(insertStr);
            return RedirectToAction("Door");
        }
        public ActionResult InsertDOOR(FabModel fab)
        {           
            string selectDoor = "Select *from Fab";
            DataTable dt = da(selectDoor);        
            return View(dt);           
        }
        public ActionResult DOORUpload()
        {
            return View();
        }
        [HttpPost]
        public ActionResult DOORUpload(InsertDoor dev, HttpPostedFileBase file)
        {
            string extension =
                    System.IO.Path.GetExtension(file.FileName);

            if (extension == ".xls" || extension == ".xlsx")
            {
                string fileLocation = Server.MapPath("~/Upload/") + file.FileName;
                file.SaveAs(fileLocation); // 存放檔案到伺服器上               
            }
            //String filePath = @"C:\Web\20201205廠區寫的Web\FactoryData - 複製\FactoryData\Upload\" + file.FileName;測試用
            String filePath = @"D:\NewWeb\Upload\" + file.FileName;
            IWorkbook workbook = null;
            FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read);
            if (filePath.IndexOf(".xlsx") > 0)
                workbook = new XSSFWorkbook(fs);
            else if (filePath.IndexOf(".xls") > 0)
                workbook = new HSSFWorkbook(fs);
            ISheet sheet = workbook.GetSheetAt(0);
            if (sheet != null)
            {
                int rowCount = sheet.LastRowNum;
                for (int i = 1; i <= rowCount; i++)
                {
                    IRow curROw = sheet.GetRow(i);
                    dev.Fab = curROw.GetCell(0).ToString();
                    dev.DoorNumber = curROw.GetCell(1).ToString();
                    dev.DoorName = curROw.GetCell(2).ToString();
                    dev.DoorIstarControlName = curROw.GetCell(3).ToString();
                    dev.DoorConnectType = curROw.GetCell(4).ToString();
                    dev.DoorReaderType = curROw.GetCell(5).ToString();
                    dev.DoorAcmNumber = curROw.GetCell(6).ToString();
                    dev.DoorReaderPort = curROw.GetCell(7).ToString();
                    dev.AddVendor= curROw.GetCell(8).ToString();
                    dev.AddDatetime = curROw.GetCell(9).ToString();
                    dev.Remarks = curROw.GetCell(10).ToString();
                    string insertStr = "insert into NEWDOORDATA(Fab,DoorNumber,DoorName,DoorIstarControlName,DoorConnectType,DoorReaderType,DoorAcmNumber,DoorReaderPort,AddVendor,AddDatetime,Remarks)" +
                        "VALUES('" + dev.Fab + "','" + dev.DoorNumber + "','" + dev.DoorName + "','" + dev.DoorIstarControlName + "','" + dev.DoorConnectType + "','" + dev.DoorReaderType + "','" + dev.DoorAcmNumber + "','"
                        + dev.DoorReaderPort + "','" + dev.AddVendor + "','" + dev.AddDatetime + "','" + dev.Remarks + "')";
                    CommandQuerytosql(insertStr);
                }
            }
            return RedirectToAction("Door");

        }
        public ActionResult DOORExcelExport(Insert ins)
        {
            //取得資料
            //var result = this.studentService.GetStudentDataList(request).ToList();
            String SqlselectDeviceText = "Select *from NEWDOORDATA";
            //建立Excel
            HSSFWorkbook hssfworkbook = new HSSFWorkbook(); //建立活頁簿
            ISheet sheet = hssfworkbook.CreateSheet("sheet"); //建立sheet

            //設定樣式
            ICellStyle headerStyle = hssfworkbook.CreateCellStyle();
            IFont headerfont = hssfworkbook.CreateFont();
            headerStyle.Alignment = HorizontalAlignment.Center; //水平置中
            headerStyle.VerticalAlignment = VerticalAlignment.Center; //垂直置中
            headerfont.FontName = "微軟正黑體";
            headerfont.FontHeightInPoints = 20;
            headerfont.Boldweight = (short)FontBoldWeight.Bold;
            headerStyle.SetFont(headerfont);
            //新增標題列
            sheet.CreateRow(0); //需先用CreateRow建立,才可通过GetRow取得該欄位
            //sheet.AddMergedRegion(new CellRangeAddress(0, 1, 0, 2)); //合併1~2列及A~C欄儲存格
            //sheet.GetRow(0).CreateCell(0).SetCellValue("昕力大學");
            //sheet.GetRow(0).GetCell(0).CellStyle = headerStyle; //套用樣式
            sheet.CreateRow(0).CreateCell(0).SetCellValue("Fab");
            sheet.GetRow(0).CreateCell(1).SetCellValue("DoorNumber");
            sheet.GetRow(0).CreateCell(2).SetCellValue("DoorName");
            sheet.GetRow(0).CreateCell(3).SetCellValue("DoorIstarControlName");
            sheet.GetRow(0).CreateCell(4).SetCellValue("DoorConnectType");
            sheet.GetRow(0).CreateCell(5).SetCellValue("DoorReaderType");
            sheet.GetRow(0).CreateCell(6).SetCellValue("DoorAcmNumber");
            sheet.GetRow(0).CreateCell(7).SetCellValue("DoorReaderPort");
            sheet.GetRow(0).CreateCell(8).SetCellValue("AddVendor");
            sheet.GetRow(0).CreateCell(9).SetCellValue("AddDatetime");
            sheet.GetRow(0).CreateCell(10).SetCellValue("EditDatetime");
            sheet.GetRow(0).CreateCell(11).SetCellValue("Remarks");
            DataTable dt = da(SqlselectDeviceText);
            //填入資料
            int rowIndex = 1;
            for (int row = 0; row < dt.Rows.Count; row++)
            {
                sheet.CreateRow(rowIndex).CreateCell(0).SetCellValue(dt.Rows[row]["Fab"].ToString());
                sheet.GetRow(rowIndex).CreateCell(1).SetCellValue(dt.Rows[row]["DoorNumber"].ToString());
                sheet.GetRow(rowIndex).CreateCell(2).SetCellValue(dt.Rows[row]["DoorName"].ToString());
                sheet.GetRow(rowIndex).CreateCell(3).SetCellValue(dt.Rows[row]["DoorIstarControlName"].ToString());
                sheet.GetRow(rowIndex).CreateCell(4).SetCellValue(dt.Rows[row]["DoorConnectType"].ToString());
                sheet.GetRow(rowIndex).CreateCell(5).SetCellValue(dt.Rows[row]["DoorReaderType"].ToString());
                sheet.GetRow(rowIndex).CreateCell(6).SetCellValue(dt.Rows[row]["DoorAcmNumber"].ToString());
                sheet.GetRow(rowIndex).CreateCell(7).SetCellValue(dt.Rows[row]["DoorReaderPort"].ToString());
                sheet.GetRow(rowIndex).CreateCell(8).SetCellValue(dt.Rows[row]["AddVendor"].ToString());
                sheet.GetRow(rowIndex).CreateCell(9).SetCellValue(dt.Rows[row]["AddDatetime"].ToString());
                sheet.GetRow(rowIndex).CreateCell(10).SetCellValue(dt.Rows[row]["EditDatetime"].ToString());
                sheet.GetRow(rowIndex).CreateCell(11).SetCellValue(dt.Rows[row]["Remarks"].ToString());
                rowIndex++;
            }
            var excelDatas = new MemoryStream();
            hssfworkbook.Write(excelDatas);
            return File(excelDatas.ToArray(), "application/vnd.ms-excel", string.Format($"門禁資料.xls"));
        }
        public ActionResult CCTV(string searchCCTVNumber, int? page, int? PageSize)
        {
            List<InsertCCTV> list = new List<InsertCCTV>();
            string strCCTV = "";
            ViewBag.PageSize = new List<SelectListItem>()
            {
                new SelectListItem() { Value="10", Text= "10" },
                new SelectListItem() { Value="30", Text= "30" },
                new SelectListItem() { Value="50", Text= "50" },
                new SelectListItem() { Value="100", Text= "100" },
                new SelectListItem() { Value="200", Text= "200" },
            };
            int pageNumber = (page ?? 1);
            int pagesize = (PageSize ?? 10);
            ViewBag.psize = pagesize;
            if (!string.IsNullOrEmpty(searchCCTVNumber))
            {
                strCCTV = "Select *from NEWCCTVDATA WHERE CCTVNumber like '%" + searchCCTVNumber + "%'";
            }
            else
            {
                strCCTV = "select *from NEWCCTVDATA order by ID DESC";
            }
            ReturnPermit ret = new ReturnPermit();
            String CCTV = ret.returnPermitCCTV(Session["Permit"].ToString(),sql);
            if (CCTV == "授權")
            {


                DataTable dt = da(strCCTV);
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    InsertCCTV ins = new InsertCCTV();
                    ins.ID = dt.Rows[i]["ID"].ToString();
                    ins.Fab = dt.Rows[i]["Fab"].ToString();
                    ins.CCTVNumber = dt.Rows[i]["CCTVNumber"].ToString();
                    ins.CCTVName = dt.Rows[i]["CCTVName"].ToString();
                    ins.CCTVIP = dt.Rows[i]["CCTVIP"].ToString();
                    ins.CCTVMAC = dt.Rows[i]["CCTVMAC"].ToString();
                    ins.CCTVBrand = dt.Rows[i]["CCTVBrand"].ToString();
                    ins.CCTVModel = dt.Rows[i]["CCTVModel"].ToString();
                    ins.CCTVType = dt.Rows[i]["CCTVType"].ToString();
                    ins.CCTVSwitchIp = dt.Rows[i]["CCTVSwitchIp"].ToString();
                    ins.CCTVSwitchPort = dt.Rows[i]["CCTVSwitchPort"].ToString();
                    ins.AddVendor = dt.Rows[i]["AddVendor"].ToString();
                    ins.AddDatetime = dt.Rows[i]["AddDatetime"].ToString();
                    ins.EditDatetime = dt.Rows[i]["EditDatetime"].ToString();
                    ins.Remarks = dt.Rows[i]["Remarks"].ToString();
                    list.Add(ins);
                }
                var pagelist = list.ToPagedList(pageNumber, pagesize);
                return View(pagelist);
            }
            else 
            {
                return View("NoAuthorityView");
            }
        }
        [HttpPost]
        public ActionResult InsertCCTV(InsertCCTV insc)
        {
            sc.ConnectionString = sql;
            string insertStr = "insert into NEWCCTVDATA(Fab,CCTVNumber,CCTVName,CCTVIP,CCTVMAC,CCTVBrand,CCTVModel,CCTVType,CCTVSwitchIp,CCTVSwitchPort,AddVendor,AddDatetime,Remarks)"+
                "VALUES('"+insc.Fab+"','"+insc.CCTVNumber+ "','" + insc.CCTVName + "','" + insc.CCTVIP + "','" + insc.CCTVMAC + "','" + insc.CCTVBrand + "','" + insc.CCTVModel + "','"
                + insc.CCTVType + "','" + insc.CCTVSwitchIp + "','" + insc.CCTVSwitchPort + "','" + insc.AddVendor + "','" + insc.AddDatetime + "','" + insc.Remarks + "')";
            //string.Format("insert into DeviceSwitchPort(ID,SwitchName,SwitchPort,DeviceType,AddDatetime,EditDatetime" +
            //")VALUES(" + ins.ID + "," + ins.SwitchName.Replace("'","''") + "," + ins.SwitchPort.Replace("'","''") + "," + ins.DeviceType.Replace("'","''")
            //+ "," + ins.AddDatetime.Replace("'","''") + "," + ins.EditDatetime.Replace("'","''") + ")"
            //);
            CommandQuerytosql(insertStr);
            return RedirectToAction("CCTV");

        }
        public ActionResult InsertCCTV(FabModel fab)
        {
            string selectFab = string.Format("Select* from Fab");
            DataTable dt = da(selectFab);
            return View(dt);
        }
        public ActionResult CCTVUpload()
        {
            return View();
        }
        [HttpPost]
        public ActionResult CCTVUpload(InsertCCTV dev, HttpPostedFileBase file)
        {
            string extension =
                    System.IO.Path.GetExtension(file.FileName);

            if (extension == ".xls" || extension == ".xlsx")
            {
                string fileLocation = Server.MapPath("~/Upload/") + file.FileName;
                file.SaveAs(fileLocation); // 存放檔案到伺服器上               
            }
            //String filePath = @"C:\Web\20201205廠區寫的Web\FactoryData - 複製\FactoryData\Upload\" + file.FileName;測試用
            String filePath = @"D:\NewWeb\Upload\" + file.FileName;

            IWorkbook workbook = null;
            FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read);
            if (filePath.IndexOf(".xlsx") > 0)
                workbook = new XSSFWorkbook(fs);
            else if (filePath.IndexOf(".xls") > 0)
                workbook = new HSSFWorkbook(fs);
            ISheet sheet = workbook.GetSheetAt(0);
            if (sheet != null)
            {
                int rowCount = sheet.LastRowNum;
                for (int i = 1; i <= rowCount; i++)
                {
                    IRow curROw = sheet.GetRow(i);
                    dev.Fab = curROw.GetCell(0).ToString();
                    dev.CCTVNumber = curROw.GetCell(1).ToString();
                    dev.CCTVName = curROw.GetCell(2).ToString();
                    dev.CCTVIP = curROw.GetCell(3).ToString();
                    dev.CCTVMAC = curROw.GetCell(4).ToString();
                    dev.CCTVBrand = curROw.GetCell(5).ToString();
                    dev.CCTVModel = curROw.GetCell(6).ToString();
                    dev.CCTVType = curROw.GetCell(7).ToString();
                    dev.CCTVSwitchIp = curROw.GetCell(8).ToString();
                    dev.CCTVSwitchPort = curROw.GetCell(9).ToString();               
                    dev.AddVendor = curROw.GetCell(10).ToString();
                    dev.AddDatetime = curROw.GetCell(11).ToString();
                    dev.Remarks = curROw.GetCell(12).ToString();
                    string insertStr = "insert into NEWCCTVDATA(Fab,CCTVNumber,CCTVName,CCTVIP,CCTVMAC,CCTVBrand,CCTVModel,CCTVType,CCTVSwitchIp,CCTVSwitchPort,AddVendor,AddDatetime,Remarks)" +
               "VALUES('" + dev.Fab + "','" + dev.CCTVNumber + "','" + dev.CCTVName + "','" + dev.CCTVIP + "','" + dev.CCTVMAC + "','" + dev.CCTVBrand + "','" + dev.CCTVModel + "','"
               + dev.CCTVType + "','" + dev.CCTVSwitchIp + "','" + dev.CCTVSwitchPort + "','" + dev.AddVendor + "','" + dev.AddDatetime + "','" + dev.Remarks + "')";
                    CommandQuerytosql(insertStr);
                }
            }
            return RedirectToAction("CCTV");

        }
        public ActionResult CCTVExcelExport(InsertCCTV ins)
        {
            //取得資料
            //var result = this.studentService.GetStudentDataList(request).ToList();
            String SqlselectDeviceText = "Select *from NEWCCTVDATA";
            //建立Excel
            HSSFWorkbook hssfworkbook = new HSSFWorkbook(); //建立活頁簿
            ISheet sheet = hssfworkbook.CreateSheet("sheet"); //建立sheet

            //設定樣式
            ICellStyle headerStyle = hssfworkbook.CreateCellStyle();
            IFont headerfont = hssfworkbook.CreateFont();
            headerStyle.Alignment = HorizontalAlignment.Center; //水平置中
            headerStyle.VerticalAlignment = VerticalAlignment.Center; //垂直置中
            headerfont.FontName = "微軟正黑體";
            headerfont.FontHeightInPoints = 20;
            headerfont.Boldweight = (short)FontBoldWeight.Bold;
            headerStyle.SetFont(headerfont);
            //新增標題列
            sheet.CreateRow(0); //需先用CreateRow建立,才可通过GetRow取得該欄位
            //sheet.AddMergedRegion(new CellRangeAddress(0, 1, 0, 2)); //合併1~2列及A~C欄儲存格
            //sheet.GetRow(0).CreateCell(0).SetCellValue("昕力大學");
            //sheet.GetRow(0).GetCell(0).CellStyle = headerStyle; //套用樣式
            sheet.CreateRow(0).CreateCell(0).SetCellValue("Fab");
            sheet.GetRow(0).CreateCell(1).SetCellValue("CCTVNumber");
            sheet.GetRow(0).CreateCell(2).SetCellValue("CCTVName");
            sheet.GetRow(0).CreateCell(3).SetCellValue("CCTVIP");
            sheet.GetRow(0).CreateCell(4).SetCellValue("CCTVMAC");
            sheet.GetRow(0).CreateCell(5).SetCellValue("CCTVBrand");
            sheet.GetRow(0).CreateCell(6).SetCellValue("CCTVModel");
            sheet.GetRow(0).CreateCell(7).SetCellValue("CCTVType");
            sheet.GetRow(0).CreateCell(8).SetCellValue("CCTVSwitchIp");
            sheet.GetRow(0).CreateCell(9).SetCellValue("CCTVSwitchPort");
            sheet.GetRow(0).CreateCell(10).SetCellValue("AddVendor");
            sheet.GetRow(0).CreateCell(11).SetCellValue("AddDatetime");
            sheet.GetRow(0).CreateCell(12).SetCellValue("EditDatetime");
            sheet.GetRow(0).CreateCell(13).SetCellValue("Remarks");
            DataTable dt = da(SqlselectDeviceText);
            //填入資料
            int rowIndex = 1;
            for (int row = 0; row < dt.Rows.Count; row++)
            {
                sheet.CreateRow(rowIndex).CreateCell(0).SetCellValue(dt.Rows[row]["Fab"].ToString());
                sheet.GetRow(rowIndex).CreateCell(1).SetCellValue(dt.Rows[row]["CCTVNumber"].ToString());
                sheet.GetRow(rowIndex).CreateCell(2).SetCellValue(dt.Rows[row]["CCTVName"].ToString());
                sheet.GetRow(rowIndex).CreateCell(3).SetCellValue(dt.Rows[row]["CCTVIP"].ToString());
                sheet.GetRow(rowIndex).CreateCell(4).SetCellValue(dt.Rows[row]["CCTVMAC"].ToString());
                sheet.GetRow(rowIndex).CreateCell(5).SetCellValue(dt.Rows[row]["CCTVBrand"].ToString());
                sheet.GetRow(rowIndex).CreateCell(6).SetCellValue(dt.Rows[row]["CCTVModel"].ToString());
                sheet.GetRow(rowIndex).CreateCell(7).SetCellValue(dt.Rows[row]["CCTVType"].ToString());
                sheet.GetRow(rowIndex).CreateCell(8).SetCellValue(dt.Rows[row]["CCTVSwitchIp"].ToString());
                sheet.GetRow(rowIndex).CreateCell(9).SetCellValue(dt.Rows[row]["CCTVSwitchPort"].ToString());
                sheet.GetRow(rowIndex).CreateCell(10).SetCellValue(dt.Rows[row]["AddVendor"].ToString());
                sheet.GetRow(rowIndex).CreateCell(11).SetCellValue(dt.Rows[row]["AddDatetime"].ToString());
                sheet.GetRow(rowIndex).CreateCell(12).SetCellValue(dt.Rows[row]["EditDatetime"].ToString());
                sheet.GetRow(rowIndex).CreateCell(13).SetCellValue(dt.Rows[row]["Remarks"].ToString());
                rowIndex++;
            }
            var excelDatas = new MemoryStream();
            hssfworkbook.Write(excelDatas);
            return File(excelDatas.ToArray(), "application/vnd.ms-excel", string.Format($"CCTV資料.xls"));
        }
        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
        public ActionResult DeviceSwitchPortDisplay(string searchDeviceSwitchPortName, int? page, int? PageSize)
        {
            ViewBag.Message = "Device";
            List<DeviceSwitchPortEdit> list = new List<DeviceSwitchPortEdit>();//利用泛型設定一個存放DeviceSwitchPortEdit model裡的元件
            string strDeviceSwitchPortName = "";
            ViewBag.PageSize = new List<SelectListItem>()//利用泛型設定取用於http.DropLise的預設值
            {
                new SelectListItem() { Value="10", Text= "10" },
                new SelectListItem() { Value="30", Text= "30" },
                new SelectListItem() { Value="50", Text= "50" },
                new SelectListItem() { Value="100", Text= "100" },
                new SelectListItem() { Value="200", Text= "200" },
            };
            int pageNumber = (page ?? 1);
            int pagesize = (PageSize ?? 10);
            ViewBag.psize = pagesize;
            if (!string.IsNullOrEmpty(searchDeviceSwitchPortName))
            {
                strDeviceSwitchPortName = "Select *from DeviceSwitchPort WHERE DeviceName like '%" + searchDeviceSwitchPortName + "%'";
            }
            else
            {
                strDeviceSwitchPortName = "select *from DeviceSwitchPort order by ID DESC";
            }
            ReturnPermit ret = new ReturnPermit();
            string device = ret.returnPermitDevice(Session["Permit"].ToString(),sql);
            if (device == "授權")
            {
                DataTable dt = da(strDeviceSwitchPortName);
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    DeviceSwitchPortEdit DSP = new DeviceSwitchPortEdit();
                    DSP.ID = dt.Rows[i]["ID"].ToString();
                    DSP.Fab = dt.Rows[i]["Fab"].ToString();
                    DSP.SwitchIP = dt.Rows[i]["SwitchIP"].ToString();
                    DSP.SwitchMAC = dt.Rows[i]["SwitchMAC"].ToString();
                    DSP.SwitchName = dt.Rows[i]["SwitchName"].ToString();
                    DSP.SwitchPort = dt.Rows[i]["SwitchPort"].ToString();
                    DSP.DeviceName = dt.Rows[i]["DeviceName"].ToString();
                    DSP.DeviceType = dt.Rows[i]["DeviceType"].ToString();
                    DSP.RackPeople = dt.Rows[i]["RackPeople"].ToString();
                    DSP.Remarks = dt.Rows[i]["Remarks"].ToString();
                    DSP.AddDatetime = dt.Rows[i]["AddDatetime"].ToString();
                    DSP.EditDatetime = dt.Rows[i]["EditDatetime"].ToString();
                    list.Add(DSP);
                }
                var pagelist = list.ToPagedList(pageNumber, pagesize);
                return View(pagelist);
            }
            else 
            {
                return View("NoAuthorityView");
            }
        }

        public void CommandQuerytosql(string sqlcom) //SQL指令回傳至資料庫
        {
            sc.ConnectionString = sql;
            sc.Open();
            SqlCommand com = new SqlCommand(sqlcom, sc);
            com.ExecuteNonQuery();//SqlCommand導入資料庫
            sc.Close();
        }
        public DataTable da(string sqlcom)//顯示資料庫資料
        {
            sc.ConnectionString = sql;
            SqlDataAdapter sd = new SqlDataAdapter(sqlcom, sc);
            DataSet ds = new DataSet();
            sd.Fill(ds);
            return ds.Tables[0];
        }
        [HttpPost]
        public ActionResult Insert(DeviceSwitchPortEdit ins)
        {
            sc.ConnectionString = sql;
            string insertStr = "insert into DeviceSwitchPort(Fab,SwitchName,SwitchIP,SwitchMAC,SwitchPort,DeviceName,DeviceType,RackPeople,AddDatetime,Remarks)" +
                "VALUES('" + ins.Fab+ "','"+ins.SwitchName + "','" + ins.SwitchIP + "','" + ins.SwitchMAC + "','" + ins.SwitchPort + "','" + ins.DeviceName + "','" + ins.DeviceType + "','" +ins.RackPeople+ "','" + ins.AddDatetime + "','" + ins.Remarks + "')";
            //string.Format("insert into DeviceSwitchPort(ID,SwitchName,SwitchPort,DeviceType,AddDatetime,EditDatetime" +
            //")VALUES(" + ins.ID + "," + ins.SwitchName.Replace("'","''") + "," + ins.SwitchPort.Replace("'","''") + "," + ins.DeviceType.Replace("'","''")
            //+ "," + ins.AddDatetime.Replace("'","''") + "," + ins.EditDatetime.Replace("'","''") + ")"
            //);
            CommandQuerytosql(insertStr);
            return RedirectToAction("DeviceSwitchPortDisplay");
        }
        public ActionResult DeviceSwitchUpload() 
        {
            return View();
        }
        [HttpPost]
        public ActionResult DeviceSwitchUpload(Insert dev,HttpPostedFileBase file)
        {
            string extension =
                    System.IO.Path.GetExtension(file.FileName);

            if (extension == ".xls" || extension == ".xlsx")
            {
                string fileLocation = Server.MapPath("~/Upload/") + file.FileName;
                file.SaveAs(fileLocation); // 存放檔案到伺服器上               
            }
        
            //String filePath = @"C:\Web\20201205廠區寫的Web\FactoryData - 複製\FactoryData\Upload\" + file.FileName;\\測試用
            String filePath = @"D:\NewWeb\Upload\" + file.FileName;

            IWorkbook workbook = null;
                FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read);
                if (filePath.IndexOf(".xlsx") > 0)
                    workbook = new XSSFWorkbook(fs);
                else if (filePath.IndexOf(".xls") > 0)
                    workbook = new HSSFWorkbook(fs);              
                ISheet sheet = workbook.GetSheetAt(0);
                if (sheet != null)
                {
                    int rowCount = sheet.LastRowNum;
                    for (int i = 1; i <= rowCount; i++)
                    {
                        IRow curROw = sheet.GetRow(i);
                         dev.Fab = curROw.GetCell(0).ToString();
                         dev.SwitchName = curROw.GetCell(1).ToString();
                         dev.SwitchIP = curROw.GetCell(2).ToString();
                         dev.SwitchMAC = curROw.GetCell(3).ToString();
                         dev.SwitchPort = curROw.GetCell(4).ToString();
                         dev.DeviceIPName = curROw.GetCell(5).ToString();
                         dev.DeviceName = curROw.GetCell(6).ToString();
                         dev.DeviceType = curROw.GetCell(7).ToString();
                         dev.RackPeople = curROw.GetCell(8).ToString();
                         dev.AddDatetime = curROw.GetCell(9).ToString();
                         dev.Remarks = curROw.GetCell(10).ToString();
                        string insertStr = "insert into DeviceSwitchPort(Fab,SwitchName,SwitchIP,SwitchMAC,SwitchPort,DeviceIPName,DeviceName,DeviceType,RackPeople,AddDatetime,Remarks)" +
                            "VALUES('" + dev.Fab + "','" + dev.SwitchName + "','" + dev.SwitchIP + "','" +dev.SwitchMAC + "','" + dev.SwitchPort + "','" + dev.DeviceIPName + "','" +dev.DeviceName + "','"
                            +dev.DeviceType + "','" +dev.RackPeople + "','"+dev.AddDatetime + "','" + dev.Remarks + "')";
                        CommandQuerytosql(insertStr);
                    }
                }                            
            return RedirectToAction("DeviceSwitchPortDisplay");

        }
        public ActionResult DeviceExcelExport(Insert ins) 
        {
            //取得資料
            //var result = this.studentService.GetStudentDataList(request).ToList();
            String SqlselectDeviceText = "Select *from DeviceSwitchPort";
            //建立Excel
            HSSFWorkbook hssfworkbook = new HSSFWorkbook(); //建立活頁簿
            ISheet sheet = hssfworkbook.CreateSheet("sheet"); //建立sheet

            //設定樣式
            ICellStyle headerStyle = hssfworkbook.CreateCellStyle();
            IFont headerfont = hssfworkbook.CreateFont();
            headerStyle.Alignment = HorizontalAlignment.Center; //水平置中
            headerStyle.VerticalAlignment = VerticalAlignment.Center; //垂直置中
            headerfont.FontName = "微軟正黑體";
            headerfont.FontHeightInPoints = 20;
            headerfont.Boldweight = (short)FontBoldWeight.Bold;
            headerStyle.SetFont(headerfont);
            //新增標題列
            sheet.CreateRow(0); //需先用CreateRow建立,才可通过GetRow取得該欄位
            //sheet.AddMergedRegion(new CellRangeAddress(0, 1, 0, 2)); //合併1~2列及A~C欄儲存格
            //sheet.GetRow(0).CreateCell(0).SetCellValue("昕力大學");
            //sheet.GetRow(0).GetCell(0).CellStyle = headerStyle; //套用樣式
            sheet.CreateRow(0).CreateCell(0).SetCellValue("Fab");
            sheet.GetRow(0).CreateCell(1).SetCellValue("SwitchName");
            sheet.GetRow(0).CreateCell(2).SetCellValue("SwitchIP");
            sheet.GetRow(0).CreateCell(3).SetCellValue("SwitchMAC");
            sheet.GetRow(0).CreateCell(4).SetCellValue("SwitchPort");
            sheet.GetRow(0).CreateCell(5).SetCellValue("DeviceIPName");
            sheet.GetRow(0).CreateCell(6).SetCellValue("DeviceName");
            sheet.GetRow(0).CreateCell(7).SetCellValue("DeviceType");
            sheet.GetRow(0).CreateCell(8).SetCellValue("RackPeople");
            sheet.GetRow(0).CreateCell(9).SetCellValue("AddDatetime");
            sheet.GetRow(0).CreateCell(10).SetCellValue("EditDatetime");
            sheet.GetRow(0).CreateCell(11).SetCellValue("Remarks");
            DataTable dt = da(SqlselectDeviceText);
            //填入資料
            int rowIndex = 1;
            for (int row = 0; row < dt.Rows.Count;row++)
            {
                sheet.CreateRow(rowIndex).CreateCell(0).SetCellValue(dt.Rows[row]["Fab"].ToString());
                sheet.GetRow(rowIndex).CreateCell(1).SetCellValue(dt.Rows[row]["SwitchName"].ToString());
                sheet.GetRow(rowIndex).CreateCell(2).SetCellValue(dt.Rows[row]["SwitchIP"].ToString());
                sheet.GetRow(rowIndex).CreateCell(3).SetCellValue(dt.Rows[row]["SwitchMAC"].ToString());
                sheet.GetRow(rowIndex).CreateCell(4).SetCellValue(dt.Rows[row]["SwitchPort"].ToString());
                sheet.GetRow(rowIndex).CreateCell(5).SetCellValue(dt.Rows[row]["DeviceIPName"].ToString());
                sheet.GetRow(rowIndex).CreateCell(6).SetCellValue(dt.Rows[row]["DeviceName"].ToString());
                sheet.GetRow(rowIndex).CreateCell(7).SetCellValue(dt.Rows[row]["DeviceType"].ToString());
                sheet.GetRow(rowIndex).CreateCell(8).SetCellValue(dt.Rows[row]["RackPeople"].ToString());
                sheet.GetRow(rowIndex).CreateCell(9).SetCellValue(dt.Rows[row]["AddDatetime"].ToString());
                sheet.GetRow(rowIndex).CreateCell(10).SetCellValue(dt.Rows[row]["EditDatetime"].ToString());
                sheet.GetRow(rowIndex).CreateCell(11).SetCellValue(dt.Rows[row]["Remarks"].ToString());
                rowIndex++;
            }
            var excelDatas = new MemoryStream();
            hssfworkbook.Write(excelDatas);
            return File(excelDatas.ToArray(), "application/vnd.ms-excel", string.Format($"設備資料.xls"));
        }
        public ActionResult ExampleDevic() 
        {
            string path = Server.MapPath("~/Upload/設備資料範本.xlsx");
            string filenamepath = System.IO.Path.GetFileName(path);
            FileStream stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
            return File(stream, "application/octet-stream", filenamepath);
        }
        public ActionResult ExampleDOOR()
        {
            string path = Server.MapPath("~/Upload/門禁設備範本.xlsx");
            string filenamepath = System.IO.Path.GetFileName(path);
            FileStream stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
            return File(stream, "application/octet-stream", filenamepath);
        }
        public ActionResult ExampleCCTV()
        {
            string path = Server.MapPath("~/Upload/CCTV資料範本.xlsx");
            string filenamepath = System.IO.Path.GetFileName(path);
            FileStream stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
            return File(stream, "application/octet-stream", filenamepath);
        }
        public ActionResult Insert(FabModel fab)
        {
            string selectDoor = "Select *from Fab";
            DataTable dt = da(selectDoor);
            return View(dt);
        }
        public ActionResult Delete(Insert ins)
        {
            string Deletestr = string.Format("Delete from DeviceSwitchPort WHERE ID =" + ins.ID.Replace("'", "''"));
            CommandQuerytosql(Deletestr);
            DataTable dt = da(Selectstr);
            return RedirectToAction("DeviceSwitchPortDisplay");
        }
        public ActionResult DoorDelete(InsertDoor insd)
        {
            string Deletestr = string.Format("Delete from NEWDOORDATA WHERE ID =" + insd.ID.Replace("'", "''"));
            CommandQuerytosql(Deletestr);
            DataTable dt = da(Selectstr);
            return RedirectToAction("Door");
        }
        public ActionResult CCTVDelete(InsertCCTV insc)
        {
            string Deletestr = string.Format("Delete from NEWCCTVDATA WHERE ID =" + insc.ID.Replace("'", "''"));
            CommandQuerytosql(Deletestr);
            DataTable dt = da(Selectstr);
            return RedirectToAction("CCTV");
        }
        [HttpPost]
        public ActionResult DeviceSwitchPortEdit(Insert edit,FabModel fab)
        {
            string SqlEdit = string.Format("UPDATE DeviceSwitchPort SET Fab = '"+edit.Fab+"',SwitchName = '" + edit.SwitchName + "',SwitchIP ='" + edit.SwitchIP + "',SwitchMAC ='" + edit.SwitchMAC + "',SwitchPort = '" + edit.SwitchPort + "',DeviceName = '" + edit.DeviceName + "',DeviceType = '" +
                edit.DeviceType + "',AddDatetime ='" + edit.AddDatetime + "',EditDatetime = GETDATE(),Remarks ='" + edit.Remarks + "'" + " WHERE ID = " + edit.ID);
            CommandQuerytosql(SqlEdit);
            return RedirectToAction("DeviceSwitchPortDisplay");
        }
        public ActionResult DeviceSwitchPortEdit(DeviceSwitchPortEdit DSEdit, string id)
        {
            id = DSEdit.ID;
            string sqlEditselect = string.Format("Select *from DeviceSwitchPort WHERE ID = " + DSEdit.ID);
            DataTable dt = da(sqlEditselect);
            return View(dt);
        }
        [HttpPost]
        public ActionResult DoorEdit(InsertDoor edit)
        {
            if (edit.AddVendor == Session["Remarks"].ToString())
            {
                string SqlEdit = string.Format("UPDATE NEWDOORDATA SET Fab='" + edit.Fab + "',DoorNumber = '" + edit.DoorNumber + "',DoorName ='" + edit.DoorName + "',DoorIstarControlName = '" + edit.DoorIstarControlName + "',DoorConnectType = '" + edit.DoorConnectType + "',DoorAcmNumber = '" +
                    edit.DoorAcmNumber + "',DoorReaderPort ='" + edit.DoorReaderPort + "',AddVendor='" + edit.AddVendor + "',AddDatetime ='" + edit.AddDatetime + "',EditDatetime = GETDATE()," + "Remarks='" + edit.Remarks + "'" + " WHERE ID = " + edit.ID);
                CommandQuerytosql(SqlEdit);
                return RedirectToAction("Door");
            }
            else 
            {        
                return RedirectToAction("DOORNoAuthority");
            }

        }
        public ActionResult DoorEdit(DOOREdit DOEdit, string id)
        {
           
                id = DOEdit.ID;
                string sqlEditselect = string.Format("Select *from NEWDOORDATA WHERE ID = " + DOEdit.ID);
                DataTable dt = da(sqlEditselect);
                return View(dt);           
        }
        public ActionResult DOORNoAuthority()
        {
            return View();
        }

        [HttpPost]
        public ActionResult CCTVEdit(InsertCCTV edit)
        {
            if (edit.AddVendor == Session["Remarks"].ToString())
            {
                string SqlEdit = string.Format("UPDATE NEWCCTVDATA SET Fab ='"+edit.Fab+"',CCTVNumber = '" + edit.CCTVNumber + "',CCTVName ='" + edit.CCTVName + "',CCTVIP = '" + edit.CCTVIP + "',CCTVMAC = '" + edit.CCTVMAC + "',CCTVBrand = '" +
                edit.CCTVBrand + "',CCTVModel ='" + edit.CCTVModel + "',CCTVType ='" + edit.CCTVType + "',CCTVSwitchIp = '" + edit.CCTVSwitchIp + "',CCTVSwitchPort = '" + edit.CCTVSwitchPort + "',AddVendor = '" + edit.AddVendor + "',AddDatetime ='" + edit.AddDatetime + "',EditDatetime = GETDATE(),Remarks ='" + edit.Remarks + "' WHERE ID = " + edit.ID);
            CommandQuerytosql(SqlEdit);
            return RedirectToAction("CCTV");
            }
            else
            {
                return View("CCTVNoAuthority");
            }
        }
        public ActionResult CCTVEdit(CCTVEdit DCTEdit, string id)
        {
           
                id = DCTEdit.ID;
                string sqlEditselect = string.Format("Select *from NEWCCTVDATA WHERE ID = " + DCTEdit.ID);
                DataTable dt = da(sqlEditselect);
                return View(dt);
           
        }
        public ActionResult CCTVNoAuthority() 
        {
            return View();
        }
        [HttpPost]
        public ActionResult Report(ReportModel rep, HttpPostedFileBase FilePhoto,HttpPostedFileBase File)
        {
            //string ImageName = "";
            //if (FilePhoto!= null)
            //{                                 
            //        ImageName = Path.GetFileName(FilePhoto.FileName);
            //        String path = Server.MapPath("~/Photos"+ ImageName) ;
            //        FilePhoto.SaveAs(path);
            //}
            //rep.ImageName = ImageName;
            string ImageName = "";          
                if (FilePhoto != null)//判斷檔案是否存在
                {
                    ImageName = System.IO.Path.GetFileName(FilePhoto.FileName);//取得File上傳資料的名稱              
                    string physicalPath = Server.MapPath("~/Photos/" + ImageName);//取得檔案路徑位置                   
                    FilePhoto.SaveAs(physicalPath);                   
                    rep.ImageName = ImageName;                  
                }
            string FileName = "";
            if (File != null)
            {
                FileName = System.IO.Path.GetFileName(File.FileName);
                string FilePathName = Server.MapPath("~/Upload/" + FileName);
                File.SaveAs(FilePathName);
                rep.FileName = FileName;
            }
            string insrepsql = "insert into Report(Title,Type,Content,RackPeople,ImageName,FileName,AddDatetime)" +
                "VALUES('" + rep.Title + "','" + rep.Type + "',N'" + rep.Content.Replace("'","''") + "','" + rep.RackPeople + "','" + rep.ImageName
                +"','" + rep.FileName + "','" + rep.AddDatetime + "')";
                CommandQuerytosql(insrepsql);
                return RedirectToAction("ReportDetial");
        }
        public ActionResult Report()
        {
            string StringStr = "select *from Company";
            DataTable dt = da(StringStr);
            return View(dt);
        }
        public ActionResult ReportDetial(string SearchReport, int? page, int? PageSize)
        {

            List<ReportModel> list = new List<ReportModel>();
            string StringSelect = "";
            ViewBag.PageSize = new List<SelectListItem>()
            {
                new SelectListItem() { Value="10", Text= "10" },
                new SelectListItem() { Value="30", Text= "30" },
                new SelectListItem() { Value="50", Text= "50" },
                new SelectListItem() { Value="100", Text= "100" },
                new SelectListItem() { Value="200", Text= "200" },
            };
            int pageNumber = (page ?? 1);
            int pagesize = (PageSize ?? 10);
            ViewBag.psize = pagesize;
            if (!string.IsNullOrEmpty(SearchReport))
            {
                StringSelect = string.Format("Select *from Report WHERE Title like '%" + SearchReport + "%'");
            }
            else
            {
                StringSelect = string.Format("Select *from Report order by ID DESC");
            }
            ReturnPermit ret = new ReturnPermit();
            string Report = ret.returnPermitReport(Session["Permit"].ToString(),sql);
            if (Report == "授權")
            {
                DataTable dt = da(StringSelect);
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    ReportModel rep = new ReportModel();
                    rep.ID = dt.Rows[i]["ID"].ToString();
                    rep.Title = dt.Rows[i]["Title"].ToString();
                    rep.Type = dt.Rows[i]["Type"].ToString();
                    rep.Content = dt.Rows[i]["Content"].ToString();
                    rep.RackPeople = dt.Rows[i]["RackPeople"].ToString();
                    rep.AddDatetime = dt.Rows[i]["AddDatetime"].ToString();
                    rep.EditDatetime = dt.Rows[i]["EditDatetime"].ToString();
                    list.Add(rep);
                }
                var pagedlist = list.ToPagedList(pageNumber, pagesize);
                return View(pagedlist);
            }
            else 
            {
                return View("NoAuthorityView");
            }
        }
        [HttpPost]
        public ActionResult ReportEdit(ReportModel edit)
        {
            if (Session["Remarks"].ToString() == edit.RackPeople)
            {
                string SqlEdit = string.Format("UPDATE Report SET Title = '" + edit.Title + "',Type ='" + edit.Type +
                    "',AddDatetime ='" + edit.AddDatetime + "',Content ='" + edit.Content + "',EditDatetime = GETDATE() WHERE ID = " + edit.ID);
                CommandQuerytosql(SqlEdit);
                return RedirectToAction("ReportDetial");
            }
            else 
            {
                return RedirectToAction("ReportNoEdirAuthority");
            }
        }
        public ActionResult ReportNoEdirAuthority()//Report無權限跳轉頁面 
        {
            return View();
        }
        public ActionResult ReportEdit(ReportEdit RepEdit, string id)//Report內容修改回傳
        {
            id = RepEdit.ID;
            string sqlEditselect = string.Format("Select *from Report WHERE ID = " + RepEdit.ID);
            DataTable dt = da(sqlEditselect);
            return View(dt);
        }
        public ActionResult ReportDelete(ReportModel insd)//Report資料刪除
        {
            string FileName = "";
            string PhotoName = "";
            string FileNamePath = "";
            string PhotoNamePath = "";
            string Deletestr = string.Format("Delete from Report WHERE ID =" + insd.ID);
            string SelectDelstr = string.Format("Select *from Report Where ID =" + insd.ID);
            DataTable dt = da(SelectDelstr);
            FileName = dt.Rows[0]["FileName"].ToString();
            PhotoName = dt.Rows[0]["ImageName"].ToString();
            //FileNamePath = System.IO.Path.GetFileName("~/Upload/" + FileName);
            //FileNamePath = @"D:\20201020FactoryData\FactoryData - 複製\FactoryData\Upload\" + FileName;//測試用
            FileNamePath = @"D:\NewWeb\Upload\"+FileName;//發佈後檔案位置
            PhotoNamePath = @"D:\NewWeb\Photos\" + PhotoName;//發佈後檔案位置
            //PhotoNamePath = System.IO.Path.GetFileName("~/Photos/" + PhotoName);
            //PhotoNamePath = @"D:\20201020FactoryData\FactoryData - 複製\FactoryData\Photos\" + PhotoName;//測試用
            FileInfo info = new FileInfo(FileNamePath);
            FileInfo info1 = new FileInfo(PhotoNamePath);
            if (info.Exists)
            {
                info.Delete();
            }
            if (info1.Exists)
            {
                info1.Delete();
            }
            CommandQuerytosql(Deletestr);
            return RedirectToAction("ReportDetial");
        }
        public ActionResult ReportContent(ReportModel rep, string id)//Display Report Content
        {
            id = rep.ID;
            string selectstr = "Select *from Report WHERE ID =" + rep.ID;
            DataTable dt = da(selectstr);
            return View(dt);
        }
        public ActionResult ReportFileDownload(ReportModel rep, string id)
        {
            id = rep.ID;
            string selectstr = "Select *from Report WHERE ID =" + rep.ID;
            DataTable dt = da(selectstr);
            string filename = dt.Rows[0]["FileName"].ToString();
            string path = Server.MapPath("~/Upload/" + filename);
            string filenamepath = System.IO.Path.GetFileName(path);
            FileStream stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
            return File(stream, "application/octet-stream", filenamepath);

        }
        public ActionResult OrnerReportFileDownload(ReportModel rep, string id)
        {
            id = rep.ID;
            string selectstr = "Select *from OrnerReport WHERE ID =" + rep.ID;
            DataTable dt = da(selectstr);
            string filename = dt.Rows[0]["FileName"].ToString();
            string path = Server.MapPath("~/Upload/" + filename);
            string filenamepath = System.IO.Path.GetFileName(path);
            FileStream stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
            return File(stream, "application/octet-stream", filenamepath);

        }
        public ActionResult OrnerReport(string SearchReport, int? page, int? PageSize)
        {

            List<OrnerReportModel> list = new List<OrnerReportModel>();
            string StringSelect = "";
            ViewBag.PageSize = new List<SelectListItem>()
            {
                new SelectListItem() { Value="10", Text= "10" },
                new SelectListItem() { Value="30", Text= "30" },
                new SelectListItem() { Value="50", Text= "50" },
                new SelectListItem() { Value="100", Text= "100" },
                new SelectListItem() { Value="200", Text= "200" },
            };
            int pageNumber = (page ?? 1);
            int pagesize = (PageSize ?? 10);
            ViewBag.psize = pagesize;
            if (!string.IsNullOrEmpty(SearchReport))
            {
                StringSelect = string.Format("Select *from OrnerReport WHERE Title like '%" + SearchReport + "%'");
            }
            else
            {
                StringSelect = string.Format("Select *from OrnerReport order by ID DESC");
            }
            ReturnPermit ret = new ReturnPermit();
            string OrnerReport = ret.returnPermitOrnerReport(Session["Permit"].ToString(),sql);
            if (OrnerReport == "授權")
            {
                DataTable dt = da(StringSelect);
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    OrnerReportModel orm = new OrnerReportModel();
                    orm.ID = dt.Rows[i]["ID"].ToString();
                    orm.Title = dt.Rows[i]["Title"].ToString();
                    orm.Type = dt.Rows[i]["Type"].ToString();
                    orm.Content = dt.Rows[i]["Content"].ToString();
                    orm.RackPeople = dt.Rows[i]["RackPeople"].ToString();
                    orm.AddDatetime = dt.Rows[i]["AddDatetime"].ToString();
                    orm.EditDatetime = dt.Rows[i]["EditDatetime"].ToString();
                    list.Add(orm);
                }
                var pagedlist = list.ToPagedList(pageNumber, pagesize);
                return View(pagedlist);
            }
            else
            {
                return View("NoAuthorityView");
            }
        }
        
        public ActionResult NoAuthorityView()
        {
            return View();
        }
        public ActionResult OrnerInsertReport()
        {
            return View();
        }
        [HttpPost]
        public ActionResult OrnerInsertReport(OrnerReportModel orm, HttpPostedFileBase FilePhoto, HttpPostedFileBase File)
        {
            //string ImageName = "";
            //if (FilePhoto!= null)
            //{                                 
            //        ImageName = Path.GetFileName(FilePhoto.FileName);
            //        String path = Server.MapPath("~/Photos"+ ImageName) ;
            //        FilePhoto.SaveAs(path);
            //}
            //rep.ImageName = ImageName;
            string ImageName = "";
            if (FilePhoto != null)
            {
                ImageName = System.IO.Path.GetFileName(FilePhoto.FileName);
                string physicalPath = Server.MapPath("~/Photos/" + ImageName);
                FilePhoto.SaveAs(physicalPath);
                orm.ImageName = ImageName;
            }
            string FileName = "";
            if (File != null)
            {
                FileName = System.IO.Path.GetFileName(File.FileName);
                string FilePathName = Server.MapPath("~/Upload/" + FileName);
                File.SaveAs(FilePathName);
                orm.FileName = FileName;
            }
            string insrepsql = "insert into OrnerReport(Title,Type,Content,RackPeople,ImageName,FileName,AddDatetime)" +
            "VALUES('" + orm.Title + "','" + orm.Type + "','" + orm.Content + "','" + orm.RackPeople + "','" + orm.ImageName
            + "','" + orm.FileName + "','" + orm.AddDatetime + "')";
            CommandQuerytosql(insrepsql);
            return RedirectToAction("OrnerReport");
        }
        [HttpPost]
        public ActionResult OrnerReportEdit(OrnerReportModel orm)
        {
            if (Session["Remarks"].ToString() == orm.RackPeople)
            {
                string SqlEdit = string.Format("UPDATE OrnerReport SET Title = '" + orm.Title + "',Type ='" + orm.Type +
                    "',AddDatetime ='" + orm.AddDatetime + "',Content ='" + orm.Content + "',EditDatetime = GETDATE() WHERE ID = " + orm.ID);
                CommandQuerytosql(SqlEdit);
                return RedirectToAction("OrnerReport");
            }
            else 
            {
                return RedirectToAction("NoOrnerReportEditAuthority");
            }
        }
        public ActionResult NoOrnerReportEditAuthority() 
        {
            return View();
        }
        public ActionResult OrnerReportEdit(OrnerReportModel orm, string id)
        {
            id = orm.ID;
            string sqlEditselect = string.Format("Select *from OrnerReport WHERE ID = " + orm.ID);
            DataTable dt = da(sqlEditselect);
            return View(dt);
        }
        public ActionResult OrnerReportDelete(OrnerReportModel insd)
        {
            string FileName = "";
            string PhotoName = "";
            string FileNamePath = "";
            string PhotoNamePath = "";
            string Deletestr = string.Format("Delete from OrnerReport WHERE ID =" + insd.ID);
            string SelectDelstr = string.Format("Select *from OrnerReport Where ID =" + insd.ID);
            DataTable dt = da(SelectDelstr);
            FileName = dt.Rows[0]["FileName"].ToString();
            PhotoName = dt.Rows[0]["ImageName"].ToString();
            //FileNamePath = System.IO.Path.GetFileName("~/Upload/" + FileName);
            //FileNamePath = @"D:\20201020FactoryData\FactoryData - 複製\FactoryData\Upload\" + FileName;//測試用
            FileNamePath = @"D:\NewWeb\Upload\"+FileName;//發佈後檔案位置
            PhotoNamePath = @"D:\NewWeb\Photos\" + PhotoName;//發佈後檔案位置
            //PhotoNamePath = System.IO.Path.GetFileName("~/Photos/" + PhotoName);
            //PhotoNamePath = @"D:\20201020FactoryData\FactoryData - 複製\FactoryData\Photos\" + PhotoName;//測試用
            FileInfo info = new FileInfo(FileNamePath);
            FileInfo info1 = new FileInfo(PhotoNamePath);
            if (info.Exists)
            {
                info.Delete();
            }
            if (info1.Exists)
            {
                info1.Delete();
            }
            CommandQuerytosql(Deletestr);
            return RedirectToAction("OrnerReport");
        }
        public ActionResult OrnerReportContent(OrnerReportModel orm, string id)
        {
            id = orm.ID;
            string selectstr = "Select *from OrnerReport WHERE ID =" + orm.ID;
            DataTable dt = da(selectstr);
            return View(dt);
        }
        public ActionResult permition(int? page, int? PageSize) 
        {
            string persql = "Select *from Permition";
            List<PermitionModel> list = new List<PermitionModel>();
            ViewBag.PageSize = new List<SelectListItem>()
            {
                new SelectListItem() { Value="10", Text= "10" },
                new SelectListItem() { Value="30", Text= "30" },
                new SelectListItem() { Value="50", Text= "50" },
                new SelectListItem() { Value="100", Text= "100" },
                new SelectListItem() { Value="200", Text= "200" },
            };
            int pageNumber = (page ?? 1);
            int pagesize = (PageSize ?? 10);
            ViewBag.psize = pagesize;
            ReturnPermit ret = new ReturnPermit();
            string Permit = ret.returnPermitPermit(Session["Permit"].ToString(),sql);
            if (Permit == "授權")
            {
                DataTable dt = da(persql);
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    PermitionModel permit = new PermitionModel();
                    permit.ID = dt.Rows[i]["ID"].ToString();
                    permit.GroupName = dt.Rows[i]["GroupName"].ToString();
                    permit.Device = dt.Rows[i]["Device"].ToString();
                    permit.Fab = dt.Rows[i]["Fab"].ToString();
                    permit.CCTV = dt.Rows[i]["CCTV"].ToString();
                    permit.DOOR = dt.Rows[i]["DOOR"].ToString();
                    permit.Report = dt.Rows[i]["Report"].ToString();
                    permit.OrnerReport = dt.Rows[i]["OrnerReport"].ToString();
                    permit.Account = dt.Rows[i]["Account"].ToString();
                    permit.Permit = dt.Rows[i]["Permit"].ToString();
                    permit.SMSCCTV = dt.Rows[i]["SMSCCTV"].ToString();
                    permit.SMSReader = dt.Rows[i]["SMSReader"].ToString();
                    permit.SMSAlarmSystem = dt.Rows[i]["SMSAlarmSystem"].ToString();
                    permit.SMSBarriergate = dt.Rows[i]["SMSBarriergate"].ToString();
                    permit.SMSGate = dt.Rows[i]["SMSGate"].ToString();
                    permit.SMSMakeCard = dt.Rows[i]["SMSMakeCard"].ToString();
                    permit.RackPeople = dt.Rows[i]["RackPeople"].ToString();
                    permit.AddDatetime = dt.Rows[i]["AddDatetime"].ToString();
                    permit.EditDatetime = dt.Rows[i]["EditDatetime"].ToString();
                    list.Add(permit);
                }
                var pagedlist = list.ToPagedList(pageNumber, pagesize);
                return View(pagedlist);
            }
            else 
            {
                return View("NoAuthorityView");
            }
        }
        public ActionResult InsertPermition() 
        {
            return View();    
        }
        [HttpPost]
        public ActionResult InsertPermition(PermitionModel per)
        {
            string perstrins = "Insert into Permition(GroupName,Device,Fab,CCTV,DOOR,Report,OrnerReport,Account,Permit,SMSCCTV,SMSReader,SMSAlarmSystem,SMSBarriergate,SMSGate,SMSMakeCard,RackPeople,AddDatetime)" +
                "VALUES('"+per.GroupName+"','"+per.Device+"','"+per.Fab+"','"+per.CCTV+"','"+per.DOOR+"','"+
                per.Report+"','"+per.OrnerReport+"','"+per.Account+"','"+per.Permit+ "','"+per.SMSCCTV + "','" + per.SMSReader + "','" + per.SMSAlarmSystem + "','" + per.SMSBarriergate + "','" + per.SMSGate + "','" + per.SMSMakeCard + "','" +per.RackPeople+"','"+per.AddDatetime+"')";
            CommandQuerytosql(perstrins);
            return RedirectToAction("Permition"); ;
        }
        [HttpPost]
        public ActionResult PermitionEdit(PermitionModel per)
        {
            if (Session["Remarks"].ToString() == per.RackPeople)
            {
                string SqlEdit = string.Format("UPDATE Permition SET GroupName = '" + per.GroupName + "',Device ='" + per.Device + "',Fab ='" + per.Device + "',CCTV ='" + per.CCTV + "',DOOR='" + per.DOOR + "',Report ='" + per.Report
                    + "',OrnerReport ='" + per.OrnerReport + "',Account ='" + per.Account + "',permit ='" + per.Permit + "',SMSCCTV ='"+per.SMSCCTV+ "',SMSReader ='" + per.SMSReader + "',SMSAlarmSystem ='" + per.SMSAlarmSystem + "',SMSBarriergate ='" + per.SMSBarriergate + "',SMSGate ='" + per.SMSGate + "',SMSMakeCard ='" + per.SMSMakeCard + "',RackPeople ='" + per.RackPeople +
                    "',AddDatetime ='" + per.AddDatetime + "',EditDatetime = GETDATE() WHERE ID = " + per.ID);
                CommandQuerytosql(SqlEdit);
                return RedirectToAction("permition");
            }
            else 
            {
                return View("NoPermitionEditAuthority");
            }               
        }
        public ActionResult NoPermitionEditAuthority() 
        {
            return View();
        }
        public ActionResult PermitionEdit(PermitionModel per, string id)
        {
            id = per.ID;
            string sqlEditselect = string.Format("Select *from Permition WHERE ID = " + per.ID);
            DataTable dt = da(sqlEditselect);
            return View(dt);
        }
        public ActionResult PermitionDelete(PermitionModel per)
        {
            string Deletestr = string.Format("Delete from permition WHERE ID =" + per.ID.Replace("'", "''"));
            CommandQuerytosql(Deletestr);
            DataTable dt = da(Selectstr);
            return RedirectToAction("permition");
        }
        //public string str(Report rep) 
        //{
        //    string sqlse = "Select *from Report";
        //    string stringforen = "";
        //    DataTable ds = da(sqlse);
        //    for (int i = 0; i < ds.Rows.Count; i++) 
        //    {
        //        stringforen += ds.Rows[i]["Title"].ToString()+"/<br>";
        //    }
        //    return stringforen;
        //}

        //public string sqltext() 
        //{
        //    SqlConnection scon = new SqlConnection();
        //    string sqle = @"server=DESKTOP-DGKQ6QE\SQLEXPRESS;initial catalog=Tabletest;user id=dino;password=iamthegad123";
        //    scon.ConnectionString = sqle;
        //    scon.Open();
        //    SqlDataAdapter sd = new SqlDataAdapter("select *from cctv$",scon);
        //    string str = "";
        //    DataTable da = new DataTable();
        //    DataSet ds = new DataSet();
        //    sd.Fill(ds);
        //    da = ds.Tables[0];
        //    for (int i = 0; i <1; i++) 
        //    {
        //        for (int j = 0; j <da.Rows.Count; j++) 
        //        {
        //            if (da.Rows[i]["CCTV"] == da.Rows[j]["CCTV"]) 
        //            {
        //                int Count = 0;
        //                Count++;
        //                if (da.Rows[i]["CCTV"] == da.Rows[j + 1]["CCTV"]) 
        //                {
        //                    Count++;
        //                }
        //                str += da.Rows[i]["CCTV"] + " " + da.Rows[j]["CCTV"] +Count.ToString()+"</br>";
        //            }


        //        }

        //    }

        //    return str;
        //}
    }
}
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using FactoryData.Models;
using PagedList;
using PagedList.Mvc;
using System.Data.SqlClient;
using System.Data;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;

namespace FactoryData.Controllers
{
    public class FAMController : Controller
    {
        //string sql = @"server=COSDF15P6VSS21\SQLEXPRESS;initial catalog=FactoryData;user id=sa;password=iamthegad123";//實體server
        //string sql = @"server=COSDF15P6VSS21\SQLEXPRESS;initial catalog=FactoryData;user id=sa;password=iamthegad123";//寫程式的Server
        //string sql = @"server=192.168.1.36;initial catalog=FactoryData;user id=sa;password=iamthegad123";//公司筆電
        string sql = @"server=192.168.1.36;initial catalog=FactoryData;user id=sa;password=iamthegad123";//公司筆電
        // GET: FAM
        SqlConnection sc = new SqlConnection();
        public ActionResult Index()
        {
            return View();
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
        public ActionResult SMSCCTV(int ? page,int ? PageSize,string searchCCTVNumber)
        {
            List<SMSCCTVModel> list = new List<SMSCCTVModel>();//SMSCCTVModel的泛型空間
            string strCCTV = "";
            ViewBag.PageSize = new List<SelectListItem>()//數值放入ViewBag暫存
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
                strCCTV = "Select *from SMSCCTV WHERE ProductNumber like '%" + searchCCTVNumber + "%'";
            }
            else
            {
                strCCTV = "select *from SMSCCTV order by ID DESC";
            }
            ReturnPermit ret = new ReturnPermit();
            String SMSCCTV1 = ret.returnSMSCCTV(Session["Permit"].ToString(), sql);//取得驗證權限的資訊
            if (SMSCCTV1 == "授權")
            {


                DataTable dt = da(strCCTV);
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    SMSCCTVModel ins = new SMSCCTVModel();
                    ins.ID = dt.Rows[i]["ID"].ToString();
                    ins.RemoveClass = dt.Rows[i]["RemoveClass"].ToString();
                    ins.System = dt.Rows[i]["System"].ToString();
                    ins.SonSystem = dt.Rows[i]["SonSystem"].ToString();
                    ins.ProductNumber = dt.Rows[i]["ProductNumber"].ToString();
                    ins.ProductDescription = dt.Rows[i]["ProductDescription"].ToString();
                    ins.Factory = dt.Rows[i]["Factory"].ToString();
                    ins.Phase = dt.Rows[i]["Phase"].ToString();
                    ins.Buliding = dt.Rows[i]["Buliding"].ToString();
                    ins.Floor = dt.Rows[i]["Floor"].ToString();
                    ins.ClassNO = dt.Rows[i]["ClassNO"].ToString();
                    ins.DeviceClass = dt.Rows[i]["DeviceClass"].ToString();
                    ins.Brand = dt.Rows[i]["Brand"].ToString();
                    ins.Type = dt.Rows[i]["Type"].ToString();
                    ins.ProductClassful = dt.Rows[i]["ProductClassful"].ToString();
                    ins.RackPeople = dt.Rows[i]["RackPeople"].ToString();
                    ins.Remarks = dt.Rows[i]["Remarks"].ToString();
                    ins.AddDatetime = dt.Rows[i]["AddDatetime"].ToString();
                    ins.Machinestate = dt.Rows[i]["Machinestate"].ToString();
                    ins.DownDate = dt.Rows[i]["DownDate"].ToString();
                    ins.EditDatetime = dt.Rows[i]["EditDatetime"].ToString();                  
                    list.Add(ins);
                }
                var pagelist = list.ToPagedList(pageNumber, pagesize);
                return View(pagelist);
            }
            else
            {
                return RedirectToAction("NoAuthorityView","Home");
            }

        }
        [HttpPost]
        public ActionResult InsertSMSCCTV(SMSCCTVModel insc)//新增SMSCCTV的資訊
        {
            sc.ConnectionString = sql;
            string insertStr = "insert into SMSCCTV(RemoveClass,System,SonSystem,ProductNumber,ProductDescription,Factory,Phase,Buliding,Floor,ClassNO,DeviceClass,Brand,Type,ProductClassful,RackPeople,Remarks,AddDatetime,Machinestate,DownDate)" +
                "VALUES('" + insc.RemoveClass + "','" + insc.System + "','" + insc.SonSystem + "','" + insc.ProductNumber + "','" + insc.ProductDescription + "','" + insc.Factory + "','" + insc.Phase + "','"
                + insc.Buliding + "','" + insc.Floor + "','" + insc.ClassNO + "','" + insc.DeviceClass + "','" +insc.Brand + "','" +insc.Type + "','"+insc.ProductClassful + "','"+insc.RackPeople + "','"+insc.Remarks + "','" +insc.AddDatetime + "','" +insc.Machinestate + "','" + insc.DownDate + "')";
            //string.Format("insert into DeviceSwitchPort(ID,SwitchName,SwitchPort,DeviceType,AddDatetime,EditDatetime" +
            //")VALUES(" + ins.ID + "," + ins.SwitchName.Replace("'","''") + "," + ins.SwitchPort.Replace("'","''") + "," + ins.DeviceType.Replace("'","''")
            //+ "," + ins.AddDatetime.Replace("'","''") + "," + ins.EditDatetime.Replace("'","''") + ")"
            //);
            CommandQuerytosql(insertStr);
            return RedirectToAction("SMSCCTV");

        }
        public ActionResult InsertSMSCCTV(FabModel fab)
        {
            string selectFab = string.Format("Select* from Fab");
            DataTable dt = da(selectFab);
            return View(dt);
        }
        public ActionResult SMSCCTVUpload()
        {
            return View();
        }
        [HttpPost]
        public ActionResult SMSCCTVUpload(SMSCCTVModel dev, HttpPostedFileBase file)
        {
            string extension =
                    System.IO.Path.GetExtension(file.FileName);

            if (extension == ".xls" || extension == ".xlsx")
            {
                string fileLocation = Server.MapPath("~/Upload/") + file.FileName;
                file.SaveAs(fileLocation); // 存放檔案到伺服器上               
            }
            //String filePath = @"D:\20201205廠區寫的Web\FactoryData - 複製\FactoryData\Upload\" + file.FileName;//測試用
            String filePath = @"C:\20201218廠區資料\20201218場內寫的WEB\20201205廠區寫的Web\FactoryData - 複製\FactoryData\Upload\" + file.FileName;//測試用

            //String filePath = @"D:\NewWeb\Upload\" + file.FileName;

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
                    dev.RemoveClass = curROw.GetCell(0).ToString();
                    dev.System = curROw.GetCell(1).ToString();
                    dev.SonSystem = curROw.GetCell(2).ToString();
                    dev.ProductNumber = curROw.GetCell(3).ToString();
                    dev.ProductDescription = curROw.GetCell(4).ToString();
                    dev.Factory = curROw.GetCell(5).ToString();
                    dev.Phase = curROw.GetCell(6).ToString();
                    dev.Buliding = curROw.GetCell(7).ToString();
                    dev.Floor = curROw.GetCell(8).ToString();
                    dev.ClassNO = curROw.GetCell(9).ToString();
                    dev.DeviceClass = curROw.GetCell(10).ToString();
                    dev.Brand = curROw.GetCell(11).ToString();
                    dev.Type = curROw.GetCell(12).ToString();
                    dev.ProductClassful = curROw.GetCell(13).ToString();
                    dev.RackPeople = curROw.GetCell(14).ToString();
                    dev.Remarks = curROw.GetCell(15).ToString();
                    dev.AddDatetime = curROw.GetCell(16).ToString();
                    dev.Machinestate = curROw.GetCell(17).ToString();
                    dev.DownDate= curROw.GetCell(18).ToString();
                    string insertStr = "insert into SMSCCTV(RemoveClass,System,SonSystem,ProductNumber,ProductDescription,Factory,Phase,Buliding,Floor,ClassNO,DeviceClass,Brand,Type,ProductClassful,RackPeople,Remarks,AddDatetime,Machinestate,DownDate)" +
                "VALUES('" + dev.RemoveClass + "','" + dev.System + "','" + dev.SonSystem + "','" + dev.ProductNumber + "','" + dev.ProductDescription + "','" + dev.Factory + "','" + dev.Phase + "','"
                + dev.Buliding + "','" + dev.Floor + "','" + dev.ClassNO + "','" + dev.DeviceClass + "','" + dev.Brand + "','" + dev.Type + "','" + dev.ProductClassful + "','" + dev.RackPeople + "','" + dev.Remarks + "','" + dev.AddDatetime + "','"+dev.Machinestate+"','" +dev.DownDate + "')";
                    CommandQuerytosql(insertStr);
                }
            }
            return RedirectToAction("SMSCCTV");

        }
        public ActionResult SMSCCTVExcelExport(SMSCCTVModel ins)
        {
            //取得資料
            //var result = this.studentService.GetStudentDataList(request).ToList();
            String SqlselectDeviceText = "Select *from SMSCCTV";
            //建立Excel
            HSSFWorkbook hssfworkbook = new HSSFWorkbook(); //建立活頁簿
            ISheet sheet = hssfworkbook.CreateSheet("資產"); //建立sheet叫資產

            //設定樣式
            ICellStyle headerStyle = hssfworkbook.CreateCellStyle();//標題樣式
            HSSFCellStyle CS = (HSSFCellStyle)hssfworkbook.CreateCellStyle();//內容樣式
            
            IFont headerfont = hssfworkbook.CreateFont();//標題文字樣式
            HSSFFont font = (HSSFFont)hssfworkbook.CreateFont();//內容文字樣式        
            font.FontName = "細明體";
            font.FontHeightInPoints = 12;
            headerStyle.Alignment = HorizontalAlignment.Left; //水平至左
            headerStyle.VerticalAlignment = VerticalAlignment.Center; //垂直置中
            //headerStyle.FillPattern = NPOI.SS.UserModel.FillPattern.SolidForeground;
            //headerStyle.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.Blue.Index;
            headerfont.FontName = "細明體";
            headerfont.FontHeightInPoints = 12;
            headerfont.Boldweight = (short)FontBoldWeight.Bold;
            headerStyle.SetFont(headerfont);//將表格內樣式套給字行
            CS.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.Blue.Index;          
            CS.SetFont(font);
            //新增標題列
            sheet.CreateRow(0); //需先用CreateRow建立,才可通过GetRow取得該欄位
            

            //sheet.AddMergedRegion(new CellRangeAddress(0, 1, 0, 2)); //合併1~2列及A~C欄儲存格
            //sheet.GetRow(0).CreateCell(0).SetCellValue("昕力大學");
            
            sheet.CreateRow(0).CreateCell(0).SetCellValue("異動類別*");
            sheet.GetRow(0).CreateCell(1).SetCellValue("系統*");
            sheet.GetRow(0).CreateCell(2).SetCellValue("子系統*");
            sheet.GetRow(0).CreateCell(3).SetCellValue("資產編號*");
            sheet.GetRow(0).CreateCell(4).SetCellValue("資產描述*");
            sheet.GetRow(0).CreateCell(5).SetCellValue("廠區*");
            sheet.GetRow(0).CreateCell(6).SetCellValue("Phase*");
            sheet.GetRow(0).CreateCell(7).SetCellValue("棟別*");
            sheet.GetRow(0).CreateCell(8).SetCellValue("樓層*");
            sheet.GetRow(0).CreateCell(9).SetCellValue("課別*");
            sheet.GetRow(0).CreateCell(10).SetCellValue("設備分類*");
            sheet.GetRow(0).CreateCell(11).SetCellValue("廠牌*");
            sheet.GetRow(0).CreateCell(12).SetCellValue("型號*");
            sheet.GetRow(0).CreateCell(13).SetCellValue("資產分類*");
            sheet.GetRow(0).GetCell(0).CellStyle = headerStyle; //套用樣式
            sheet.GetRow(0).GetCell(1).CellStyle = headerStyle; //套用樣式
            sheet.GetRow(0).GetCell(2).CellStyle = headerStyle; //套用樣式
            sheet.GetRow(0).GetCell(3).CellStyle = headerStyle; //套用樣式
            sheet.GetRow(0).GetCell(4).CellStyle = headerStyle; //套用樣式
            sheet.GetRow(0).GetCell(5).CellStyle = headerStyle; //套用樣式
            sheet.GetRow(0).GetCell(6).CellStyle = headerStyle; //套用樣式
            sheet.GetRow(0).GetCell(7).CellStyle = headerStyle; //套用樣式
            sheet.GetRow(0).GetCell(8).CellStyle = headerStyle; //套用樣式
            sheet.GetRow(0).GetCell(9).CellStyle = headerStyle; //套用樣式
            sheet.GetRow(0).GetCell(10).CellStyle = headerStyle; //套用樣式
            sheet.GetRow(0).GetCell(11).CellStyle = headerStyle; //套用樣式
            sheet.GetRow(0).GetCell(12).CellStyle = headerStyle; //套用樣式
            sheet.GetRow(0).GetCell(13).CellStyle = headerStyle; //套用樣式

            DataTable dt = da(SqlselectDeviceText);
            //填入資料
            int rowIndex = 1;
            for (int row = 0; row < dt.Rows.Count; row++)
            {
                sheet.CreateRow(rowIndex).CreateCell(0).SetCellValue(dt.Rows[row]["RemoveClass"].ToString());
                sheet.GetRow(rowIndex).CreateCell(1).SetCellValue(dt.Rows[row]["System"].ToString());
                sheet.GetRow(rowIndex).CreateCell(2).SetCellValue(dt.Rows[row]["SonSystem"].ToString());
                sheet.GetRow(rowIndex).CreateCell(3).SetCellValue(dt.Rows[row]["ProductNumber"].ToString());
                sheet.GetRow(rowIndex).CreateCell(4).SetCellValue(dt.Rows[row]["ProductDescription"].ToString());
                sheet.GetRow(rowIndex).CreateCell(5).SetCellValue(dt.Rows[row]["Factory"].ToString());
                sheet.GetRow(rowIndex).CreateCell(6).SetCellValue(dt.Rows[row]["Phase"].ToString());
                sheet.GetRow(rowIndex).CreateCell(7).SetCellValue(dt.Rows[row]["Buliding"].ToString());
                sheet.GetRow(rowIndex).CreateCell(8).SetCellValue(dt.Rows[row]["Floor"].ToString());
                sheet.GetRow(rowIndex).CreateCell(9).SetCellValue("'"+dt.Rows[row]["ClassNO"].ToString());
                sheet.GetRow(rowIndex).CreateCell(10).SetCellValue(dt.Rows[row]["DeviceClass"].ToString());
                sheet.GetRow(rowIndex).CreateCell(11).SetCellValue(dt.Rows[row]["Brand"].ToString());
                sheet.GetRow(rowIndex).CreateCell(12).SetCellValue(dt.Rows[row]["Type"].ToString());
                if (dt.Rows[row]["ProductClassful"].ToString() == "N/A")
                {
                    sheet.GetRow(rowIndex).CreateCell(13).SetCellValue("");
                }
                else {
                    sheet.GetRow(rowIndex).CreateCell(13).SetCellValue(dt.Rows[row]["ProductClassful"].ToString());
                }
                sheet.GetRow(rowIndex).GetCell(0).CellStyle = CS; //套用樣式
                sheet.GetRow(rowIndex).GetCell(1).CellStyle = CS; //套用樣式
                sheet.GetRow(rowIndex).GetCell(2).CellStyle = CS; //套用樣式
                sheet.GetRow(rowIndex).GetCell(3).CellStyle = CS; //套用樣式
                sheet.GetRow(rowIndex).GetCell(4).CellStyle = CS; //套用樣式
                sheet.GetRow(rowIndex).GetCell(5).CellStyle = CS; //套用樣式
                sheet.GetRow(rowIndex).GetCell(6).CellStyle = CS; //套用樣式
                sheet.GetRow(rowIndex).GetCell(7).CellStyle = CS; //套用樣式
                sheet.GetRow(rowIndex).GetCell(8).CellStyle = CS; //套用樣式
                sheet.GetRow(rowIndex).GetCell(9).CellStyle = CS; //套用樣式
                sheet.GetRow(rowIndex).GetCell(10).CellStyle = CS; //套用樣式
                sheet.GetRow(rowIndex).GetCell(11).CellStyle = CS; //套用樣式
                sheet.GetRow(rowIndex).GetCell(12).CellStyle = CS; //套用樣式
                sheet.GetRow(rowIndex).GetCell(13).CellStyle = CS; //套用樣式
                sheet.SetColumnWidth(rowIndex-1, 5000);
                rowIndex++;
            }
            var excelDatas = new MemoryStream();
            hssfworkbook.Write(excelDatas);
            return File(excelDatas.ToArray(), "application/vnd.ms-excel", string.Format($"FMCS_SMSCCTV_資產.xls"));
        }
        public ActionResult ExampleSMSCCTV()
        {
            string path = Server.MapPath("~/Upload/FAM_SMSCCTV_範本.xls");
            string filenamepath = System.IO.Path.GetFileName(path);
            FileStream stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
            return File(stream, "application/octet-stream", filenamepath);
        }
        [HttpPost]
        public ActionResult SMSCCTVEdit(SMSCCTVModel edit)
        {
            if (edit.RackPeople == Session["Remarks"].ToString())
            {
                string SqlEdit = string.Format("UPDATE SMSCCTV SET RemoveClass ='" + edit.RemoveClass + "',System = '" + edit.System + "',SonSystem ='" + edit.SonSystem + "',ProductNumber = '" + edit.ProductNumber + "',ProductDescription = '" + edit.ProductDescription + "',Factory = '" +
                edit.Factory + "',Phase ='" + edit.Phase + "',Buliding ='" + edit.Buliding + "',Floor = '" + edit.Floor + "',ClassNO = '" + edit.ClassNO + "',DeviceClass = '" + edit.DeviceClass + "',Brand = '" +edit.Brand+ "',Type = '"+edit.Type + "',ProductClassful = '"+edit.ProductClassful+ "',Remarks = '" + edit.Remarks  + "',AddDatetime ='" + edit.AddDatetime + "',Machinestate ='" + edit.Machinestate + "',DownDate ='" + edit.DownDate + "',EditDatetime = GETDATE()" + " WHERE ID = " + edit.ID);
                CommandQuerytosql(SqlEdit);
                return RedirectToAction("SMSCCTV");
            }
            else
            {
                return RedirectToAction("NoAuthoritySMSCCTV", "Home");
            }
        }
        public ActionResult SMSCCTVEdit(SMSCCTVModel DCTEdit, string id)
        {

            id = DCTEdit.ID;
            string sqlEditselect = string.Format("Select *from SMSCCTV WHERE ID = " + DCTEdit.ID);
            DataTable dt = da(sqlEditselect);
            return View(dt);

        }
        public ActionResult SMSCCTVDelete(SMSCCTVModel insc)
        {
            string Deletestr = string.Format("Delete from SMSCCTV WHERE ID =" + insc.ID.Replace("'", "''"));
            CommandQuerytosql(Deletestr);          
            return RedirectToAction("SMSCCTV","FAM");
        }
        public ActionResult SMSReader(int? page, int? PageSize, string searchReaderNumber)
        {
            List<SMSCCTVModel> list = new List<SMSCCTVModel>();//SMSCCTVModel的泛型空間
            string strCCTV = "";
            ViewBag.PageSize = new List<SelectListItem>()//數值放入ViewBag暫存
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
            if (!string.IsNullOrEmpty(searchReaderNumber))
            {
                strCCTV = "Select *from SMSReader WHERE ProductNumber like '%" + searchReaderNumber + "%'";
            }
            else
            {
                strCCTV = "select *from SMSReader order by ID DESC";
            }
            ReturnPermit ret = new ReturnPermit();
            String SMSCCTV1 = ret.returnSMSReader(Session["Permit"].ToString(), sql);//取得驗證權限的資訊
            if (SMSCCTV1 == "授權")
            {


                DataTable dt = da(strCCTV);
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    SMSCCTVModel ins = new SMSCCTVModel();
                    ins.ID = dt.Rows[i]["ID"].ToString();
                    ins.RemoveClass = dt.Rows[i]["RemoveClass"].ToString();
                    ins.System = dt.Rows[i]["System"].ToString();
                    ins.SonSystem = dt.Rows[i]["SonSystem"].ToString();
                    ins.ProductNumber = dt.Rows[i]["ProductNumber"].ToString();
                    ins.ProductDescription = dt.Rows[i]["ProductDescription"].ToString();
                    ins.Factory = dt.Rows[i]["Factory"].ToString();
                    ins.Phase = dt.Rows[i]["Phase"].ToString();
                    ins.Buliding = dt.Rows[i]["Buliding"].ToString();
                    ins.Floor = dt.Rows[i]["Floor"].ToString();
                    ins.ClassNO = dt.Rows[i]["ClassNO"].ToString();
                    ins.DeviceClass = dt.Rows[i]["DeviceClass"].ToString();
                    ins.Brand = dt.Rows[i]["Brand"].ToString();
                    ins.Type = dt.Rows[i]["Type"].ToString();
                    ins.ProductClassful = dt.Rows[i]["ProductClassful"].ToString();
                    ins.RackPeople = dt.Rows[i]["RackPeople"].ToString();
                    ins.Remarks = dt.Rows[i]["Remarks"].ToString();
                    ins.AddDatetime = dt.Rows[i]["AddDatetime"].ToString();
                    ins.Machinestate = dt.Rows[i]["Machinestate"].ToString();
                    ins.DownDate = dt.Rows[i]["DownDate"].ToString();
                    ins.EditDatetime = dt.Rows[i]["EditDatetime"].ToString();
                    list.Add(ins);
                }
                var pagelist = list.ToPagedList(pageNumber, pagesize);
                return View(pagelist);
            }
            else
            {
                return RedirectToAction("NoAuthorityView", "Home");
            }

        }
        [HttpPost]
        public ActionResult InsertSMSReader(SMSCCTVModel insc)//新增SMSCCTV的資訊
        {
            sc.ConnectionString = sql;
            string insertStr = "insert into SMSReader(RemoveClass,System,SonSystem,ProductNumber,ProductDescription,Factory,Phase,Buliding,Floor,ClassNO,DeviceClass,Brand,Type,ProductClassful,RackPeople,Remarks,AddDatetime,Machinestate,DownDate)" +
                "VALUES('" + insc.RemoveClass + "','" + insc.System + "','" + insc.SonSystem + "','" + insc.ProductNumber + "','" + insc.ProductDescription + "','" + insc.Factory + "','" + insc.Phase + "','"
                + insc.Buliding + "','" + insc.Floor + "','" + insc.ClassNO + "','" + insc.DeviceClass + "','" + insc.Brand + "','" + insc.Type + "','" + insc.ProductClassful + "','" + insc.RackPeople + "','" + insc.Remarks + "','" + insc.AddDatetime + "','" + insc.Machinestate + "','" + insc.DownDate + "')";
            //string.Format("insert into DeviceSwitchPort(ID,SwitchName,SwitchPort,DeviceType,AddDatetime,EditDatetime" +
            //")VALUES(" + ins.ID + "," + ins.SwitchName.Replace("'","''") + "," + ins.SwitchPort.Replace("'","''") + "," + ins.DeviceType.Replace("'","''")
            //+ "," + ins.AddDatetime.Replace("'","''") + "," + ins.EditDatetime.Replace("'","''") + ")"
            //);
            CommandQuerytosql(insertStr);
            return RedirectToAction("SMSReader");

        }
        public ActionResult InsertSMSReader(FabModel fab)
        {
            string selectFab = string.Format("Select* from Fab");
            DataTable dt = da(selectFab);
            return View(dt);
        }
        public ActionResult SMSReaderUpload()
        {
            return View();
        }
        [HttpPost]
        public ActionResult SMSReaderUpload(SMSCCTVModel dev, HttpPostedFileBase file)
        {
            string extension =
                    System.IO.Path.GetExtension(file.FileName);

            if (extension == ".xls" || extension == ".xlsx")
            {
                string fileLocation = Server.MapPath("~/Upload/") + file.FileName;
                file.SaveAs(fileLocation); // 存放檔案到伺服器上               
            }
            //String filePath = @"D:\20201205廠區寫的Web\FactoryData - 複製\FactoryData\Upload\" + file.FileName;//測試用
            String filePath = @"C:\20201218廠區資料\20201218場內寫的WEB\20201205廠區寫的Web\FactoryData - 複製\FactoryData\Upload\" + file.FileName;//測試用

            //String filePath = @"D:\NewWeb\Upload\" + file.FileName;

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
                    dev.RemoveClass = curROw.GetCell(0).ToString();
                    dev.System = curROw.GetCell(1).ToString();
                    dev.SonSystem = curROw.GetCell(2).ToString();
                    dev.ProductNumber = curROw.GetCell(3).ToString();
                    dev.ProductDescription = curROw.GetCell(4).ToString();
                    dev.Factory = curROw.GetCell(5).ToString();
                    dev.Phase = curROw.GetCell(6).ToString();
                    dev.Buliding = curROw.GetCell(7).ToString();
                    dev.Floor = curROw.GetCell(8).ToString();
                    dev.ClassNO = curROw.GetCell(9).ToString();
                    dev.DeviceClass = curROw.GetCell(10).ToString();
                    dev.Brand = curROw.GetCell(11).ToString();
                    dev.Type = curROw.GetCell(12).ToString();
                    dev.ProductClassful = curROw.GetCell(13).ToString();
                    dev.RackPeople = curROw.GetCell(14).ToString();
                    dev.Remarks = curROw.GetCell(15).ToString();
                    dev.AddDatetime = curROw.GetCell(16).ToString();
                    dev.Machinestate = curROw.GetCell(17).ToString();
                    dev.DownDate = curROw.GetCell(18).ToString();
                    string insertStr = "insert into SMSReader(RemoveClass,System,SonSystem,ProductNumber,ProductDescription,Factory,Phase,Buliding,Floor,ClassNO,DeviceClass,Brand,Type,ProductClassful,RackPeople,Remarks,AddDatetime,Machinestate,DownDate)" +
                "VALUES('" + dev.RemoveClass + "','" + dev.System + "','" + dev.SonSystem + "','" + dev.ProductNumber + "','" + dev.ProductDescription + "','" + dev.Factory + "','" + dev.Phase + "','"
                + dev.Buliding + "','" + dev.Floor + "','" + dev.ClassNO + "','" + dev.DeviceClass + "','" + dev.Brand + "','" + dev.Type + "','" + dev.ProductClassful + "','" + dev.RackPeople + "','" + dev.Remarks + "','" + dev.AddDatetime + "','" + dev.Machinestate + "','" + dev.DownDate + "')";
                    CommandQuerytosql(insertStr);
                }
            }
            return RedirectToAction("SMSReader");

        }
        public ActionResult SMSReaderExcelExport(SMSCCTVModel ins)
        {
            //取得資料
            //var result = this.studentService.GetStudentDataList(request).ToList();
            String SqlselectDeviceText = "Select *from SMSReader";
            //建立Excel
            HSSFWorkbook hssfworkbook = new HSSFWorkbook(); //建立活頁簿
            ISheet sheet = hssfworkbook.CreateSheet("資產"); //建立sheet叫資產

            //設定樣式
            ICellStyle headerStyle = hssfworkbook.CreateCellStyle();//標題樣式
            HSSFCellStyle CS = (HSSFCellStyle)hssfworkbook.CreateCellStyle();//內容樣式

            IFont headerfont = hssfworkbook.CreateFont();//標題文字樣式
            HSSFFont font = (HSSFFont)hssfworkbook.CreateFont();//內容文字樣式        
            font.FontName = "細明體";
            font.FontHeightInPoints = 12;
            headerStyle.Alignment = HorizontalAlignment.Left; //水平至左
            headerStyle.VerticalAlignment = VerticalAlignment.Center; //垂直置中
            //headerStyle.FillPattern = NPOI.SS.UserModel.FillPattern.SolidForeground;
            //headerStyle.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.Blue.Index;
            headerfont.FontName = "細明體";
            headerfont.FontHeightInPoints = 12;
            headerfont.Boldweight = (short)FontBoldWeight.Bold;
            headerStyle.SetFont(headerfont);//將表格內樣式套給字行
            CS.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.Blue.Index;
            CS.SetFont(font);
            //新增標題列
            sheet.CreateRow(0); //需先用CreateRow建立,才可通过GetRow取得該欄位


            //sheet.AddMergedRegion(new CellRangeAddress(0, 1, 0, 2)); //合併1~2列及A~C欄儲存格
            //sheet.GetRow(0).CreateCell(0).SetCellValue("昕力大學");

            sheet.CreateRow(0).CreateCell(0).SetCellValue("異動類別*");
            sheet.GetRow(0).CreateCell(1).SetCellValue("系統*");
            sheet.GetRow(0).CreateCell(2).SetCellValue("子系統*");
            sheet.GetRow(0).CreateCell(3).SetCellValue("資產編號*");
            sheet.GetRow(0).CreateCell(4).SetCellValue("資產描述*");
            sheet.GetRow(0).CreateCell(5).SetCellValue("廠區*");
            sheet.GetRow(0).CreateCell(6).SetCellValue("Phase*");
            sheet.GetRow(0).CreateCell(7).SetCellValue("棟別*");
            sheet.GetRow(0).CreateCell(8).SetCellValue("樓層*");
            sheet.GetRow(0).CreateCell(9).SetCellValue("課別*");
            sheet.GetRow(0).CreateCell(10).SetCellValue("設備分類*");
            sheet.GetRow(0).CreateCell(11).SetCellValue("廠牌*");
            sheet.GetRow(0).CreateCell(12).SetCellValue("型號*");
            sheet.GetRow(0).CreateCell(13).SetCellValue("資產分類*");
            sheet.GetRow(0).GetCell(0).CellStyle = headerStyle; //套用樣式
            sheet.GetRow(0).GetCell(1).CellStyle = headerStyle; //套用樣式
            sheet.GetRow(0).GetCell(2).CellStyle = headerStyle; //套用樣式
            sheet.GetRow(0).GetCell(3).CellStyle = headerStyle; //套用樣式
            sheet.GetRow(0).GetCell(4).CellStyle = headerStyle; //套用樣式
            sheet.GetRow(0).GetCell(5).CellStyle = headerStyle; //套用樣式
            sheet.GetRow(0).GetCell(6).CellStyle = headerStyle; //套用樣式
            sheet.GetRow(0).GetCell(7).CellStyle = headerStyle; //套用樣式
            sheet.GetRow(0).GetCell(8).CellStyle = headerStyle; //套用樣式
            sheet.GetRow(0).GetCell(9).CellStyle = headerStyle; //套用樣式
            sheet.GetRow(0).GetCell(10).CellStyle = headerStyle; //套用樣式
            sheet.GetRow(0).GetCell(11).CellStyle = headerStyle; //套用樣式
            sheet.GetRow(0).GetCell(12).CellStyle = headerStyle; //套用樣式
            sheet.GetRow(0).GetCell(13).CellStyle = headerStyle; //套用樣式

            DataTable dt = da(SqlselectDeviceText);
            //填入資料
            int rowIndex = 1;
            for (int row = 0; row < dt.Rows.Count; row++)
            {
                sheet.CreateRow(rowIndex).CreateCell(0).SetCellValue(dt.Rows[row]["RemoveClass"].ToString());
                sheet.GetRow(rowIndex).CreateCell(1).SetCellValue(dt.Rows[row]["System"].ToString());
                sheet.GetRow(rowIndex).CreateCell(2).SetCellValue(dt.Rows[row]["SonSystem"].ToString());
                sheet.GetRow(rowIndex).CreateCell(3).SetCellValue(dt.Rows[row]["ProductNumber"].ToString());
                sheet.GetRow(rowIndex).CreateCell(4).SetCellValue(dt.Rows[row]["ProductDescription"].ToString());
                sheet.GetRow(rowIndex).CreateCell(5).SetCellValue(dt.Rows[row]["Factory"].ToString());
                sheet.GetRow(rowIndex).CreateCell(6).SetCellValue(dt.Rows[row]["Phase"].ToString());
                sheet.GetRow(rowIndex).CreateCell(7).SetCellValue(dt.Rows[row]["Buliding"].ToString());
                sheet.GetRow(rowIndex).CreateCell(8).SetCellValue(dt.Rows[row]["Floor"].ToString());
                sheet.GetRow(rowIndex).CreateCell(9).SetCellValue("'" + dt.Rows[row]["ClassNO"].ToString());
                sheet.GetRow(rowIndex).CreateCell(10).SetCellValue(dt.Rows[row]["DeviceClass"].ToString());
                sheet.GetRow(rowIndex).CreateCell(11).SetCellValue(dt.Rows[row]["Brand"].ToString());
                sheet.GetRow(rowIndex).CreateCell(12).SetCellValue(dt.Rows[row]["Type"].ToString());
                if (dt.Rows[row]["ProductClassful"].ToString() == "N/A")
                {
                    sheet.GetRow(rowIndex).CreateCell(13).SetCellValue("");
                }
                else
                {
                    sheet.GetRow(rowIndex).CreateCell(13).SetCellValue(dt.Rows[row]["ProductClassful"].ToString());
                }
                sheet.GetRow(rowIndex).GetCell(0).CellStyle = CS; //套用樣式
                sheet.GetRow(rowIndex).GetCell(1).CellStyle = CS; //套用樣式
                sheet.GetRow(rowIndex).GetCell(2).CellStyle = CS; //套用樣式
                sheet.GetRow(rowIndex).GetCell(3).CellStyle = CS; //套用樣式
                sheet.GetRow(rowIndex).GetCell(4).CellStyle = CS; //套用樣式
                sheet.GetRow(rowIndex).GetCell(5).CellStyle = CS; //套用樣式
                sheet.GetRow(rowIndex).GetCell(6).CellStyle = CS; //套用樣式
                sheet.GetRow(rowIndex).GetCell(7).CellStyle = CS; //套用樣式
                sheet.GetRow(rowIndex).GetCell(8).CellStyle = CS; //套用樣式
                sheet.GetRow(rowIndex).GetCell(9).CellStyle = CS; //套用樣式
                sheet.GetRow(rowIndex).GetCell(10).CellStyle = CS; //套用樣式
                sheet.GetRow(rowIndex).GetCell(11).CellStyle = CS; //套用樣式
                sheet.GetRow(rowIndex).GetCell(12).CellStyle = CS; //套用樣式
                sheet.GetRow(rowIndex).GetCell(13).CellStyle = CS; //套用樣式
                sheet.SetColumnWidth(rowIndex - 1, 5000);
                rowIndex++;
            }
            var excelDatas = new MemoryStream();
            hssfworkbook.Write(excelDatas);
            return File(excelDatas.ToArray(), "application/vnd.ms-excel", string.Format($"FMCS_SMS-讀卡機_資產.xls"));
        }
        public ActionResult ExampleSMSReader()
        {
            string path = Server.MapPath("~/Upload/SMS-讀卡機_範本.xls");
            string filenamepath = System.IO.Path.GetFileName(path);
            FileStream stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
            return File(stream, "application/octet-stream", filenamepath);
        }
        [HttpPost]
        public ActionResult SMSReaderEdit(SMSCCTVModel edit)
        {
            if (edit.RackPeople == Session["Remarks"].ToString())
            {
                string SqlEdit = string.Format("UPDATE SMSCCTV SET RemoveClass ='" + edit.RemoveClass + "',System = '" + edit.System + "',SonSystem ='" + edit.SonSystem + "',ProductNumber = '" + edit.ProductNumber + "',ProductDescription = '" + edit.ProductDescription + "',Factory = '" +
                edit.Factory + "',Phase ='" + edit.Phase + "',Buliding ='" + edit.Buliding + "',Floor = '" + edit.Floor + "',ClassNO = '" + edit.ClassNO + "',DeviceClass = '" + edit.DeviceClass + "',Brand = '" + edit.Brand + "',Type = '" + edit.Type + "',ProductClassful = '" + edit.ProductClassful + "',Remarks = '" + edit.Remarks + "',AddDatetime ='" + edit.AddDatetime + "',Machinestate ='" + edit.Machinestate + "',DownDate ='" + edit.DownDate + "',EditDatetime = GETDATE()" + " WHERE ID = " + edit.ID);
                CommandQuerytosql(SqlEdit);
                return RedirectToAction("SMSReader");
            }
            else
            {
                return RedirectToAction("NoAuthoritySMSCCTV", "Home");
            }
        }
        public ActionResult SMSReaderEdit(SMSCCTVModel DCTEdit, string id)
        {

            id = DCTEdit.ID;
            string sqlEditselect = string.Format("Select *from SMSReader WHERE ID = " + DCTEdit.ID);
            DataTable dt = da(sqlEditselect);
            return View(dt);

        }
        public ActionResult SMSReaderDelete(SMSCCTVModel insc)
        {
            string Deletestr = string.Format("Delete from SMSReader WHERE ID =" + insc.ID.Replace("'", "''"));
            CommandQuerytosql(Deletestr);
            return RedirectToAction("SMSReader", "FAM");
        }
        public ActionResult SMSAlarmSystem(int? page, int? PageSize, string searchAlarmNumber)
        {
            List<SMSCCTVModel> list = new List<SMSCCTVModel>();//SMSCCTVModel的泛型空間
            string strCCTV = "";
            ViewBag.PageSize = new List<SelectListItem>()//數值放入ViewBag暫存
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
            if (!string.IsNullOrEmpty(searchAlarmNumber))
            {
                strCCTV = "Select *from SMSAlarmSystem WHERE ProductNumber like '%" + searchAlarmNumber + "%'";
            }
            else
            {
                strCCTV = "select *from SMSAlarmSystem order by ID DESC";
            }
            ReturnPermit ret = new ReturnPermit();
            String SMSCCTV1 = ret.returnSMSAlarmSystem(Session["Permit"].ToString(), sql);//取得驗證權限的資訊
            if (SMSCCTV1 == "授權")
            {


                DataTable dt = da(strCCTV);
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    SMSCCTVModel ins = new SMSCCTVModel();
                    ins.ID = dt.Rows[i]["ID"].ToString();
                    ins.RemoveClass = dt.Rows[i]["RemoveClass"].ToString();
                    ins.System = dt.Rows[i]["System"].ToString();
                    ins.SonSystem = dt.Rows[i]["SonSystem"].ToString();
                    ins.ProductNumber = dt.Rows[i]["ProductNumber"].ToString();
                    ins.ProductDescription = dt.Rows[i]["ProductDescription"].ToString();
                    ins.Factory = dt.Rows[i]["Factory"].ToString();
                    ins.Phase = dt.Rows[i]["Phase"].ToString();
                    ins.Buliding = dt.Rows[i]["Buliding"].ToString();
                    ins.Floor = dt.Rows[i]["Floor"].ToString();
                    ins.ClassNO = dt.Rows[i]["ClassNO"].ToString();
                    ins.DeviceClass = dt.Rows[i]["DeviceClass"].ToString();
                    ins.Brand = dt.Rows[i]["Brand"].ToString();
                    ins.Type = dt.Rows[i]["Type"].ToString();
                    ins.ProductClassful = dt.Rows[i]["ProductClassful"].ToString();
                    ins.RackPeople = dt.Rows[i]["RackPeople"].ToString();
                    ins.Remarks = dt.Rows[i]["Remarks"].ToString();
                    ins.AddDatetime = dt.Rows[i]["AddDatetime"].ToString();
                    ins.Machinestate = dt.Rows[i]["Machinestate"].ToString();
                    ins.DownDate = dt.Rows[i]["DownDate"].ToString();
                    ins.EditDatetime = dt.Rows[i]["EditDatetime"].ToString();
                    list.Add(ins);
                }
                var pagelist = list.ToPagedList(pageNumber, pagesize);
                return View(pagelist);
            }
            else
            {
                return RedirectToAction("NoAuthorityView", "Home");
            }

        }
        [HttpPost]
        public ActionResult InsertSMSAlarmSystem(SMSCCTVModel insc)//新增SMSCCTV的資訊
        {
            sc.ConnectionString = sql;
            string insertStr = "insert into SMSAlarmSystem(RemoveClass,System,SonSystem,ProductNumber,ProductDescription,Factory,Phase,Buliding,Floor,ClassNO,DeviceClass,Brand,Type,ProductClassful,RackPeople,Remarks,AddDatetime,Machinestate,DownDate)" +
                "VALUES('" + insc.RemoveClass + "','" + insc.System + "','" + insc.SonSystem + "','" + insc.ProductNumber + "','" + insc.ProductDescription + "','" + insc.Factory + "','" + insc.Phase + "','"
                + insc.Buliding + "','" + insc.Floor + "','" + insc.ClassNO + "','" + insc.DeviceClass + "','" + insc.Brand + "','" + insc.Type + "','" + insc.ProductClassful + "','" + insc.RackPeople + "','" + insc.Remarks + "','" + insc.AddDatetime + "','" + insc.Machinestate + "','" + insc.DownDate + "')";
            //string.Format("insert into DeviceSwitchPort(ID,SwitchName,SwitchPort,DeviceType,AddDatetime,EditDatetime" +
            //")VALUES(" + ins.ID + "," + ins.SwitchName.Replace("'","''") + "," + ins.SwitchPort.Replace("'","''") + "," + ins.DeviceType.Replace("'","''")
            //+ "," + ins.AddDatetime.Replace("'","''") + "," + ins.EditDatetime.Replace("'","''") + ")"
            //);
            CommandQuerytosql(insertStr);
            return RedirectToAction("SMSAlarmSystem");

        }
        public ActionResult InsertSMSAlarmSystem(FabModel fab)
        {
            string selectFab = string.Format("Select* from Fab");
            DataTable dt = da(selectFab);
            return View(dt);
        }
        public ActionResult SMSAlarmSystemUpload()
        {
            return View();
        }
        [HttpPost]
        public ActionResult SMSAlarmSystemUpload(SMSCCTVModel dev, HttpPostedFileBase file)
        {
            string extension =
                    System.IO.Path.GetExtension(file.FileName);

            if (extension == ".xls" || extension == ".xlsx")
            {
                string fileLocation = Server.MapPath("~/Upload/") + file.FileName;
                file.SaveAs(fileLocation); // 存放檔案到伺服器上               
            }
            //String filePath = @"D:\20201205廠區寫的Web\FactoryData - 複製\FactoryData\Upload\" + file.FileName;//測試用
            String filePath = @"C:\20201218廠區資料\20201218場內寫的WEB\20201205廠區寫的Web\FactoryData - 複製\FactoryData\Upload\" + file.FileName;//測試用

            //String filePath = @"D:\NewWeb\Upload\" + file.FileName;

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
                    dev.RemoveClass = curROw.GetCell(0).ToString();
                    dev.System = curROw.GetCell(1).ToString();
                    dev.SonSystem = curROw.GetCell(2).ToString();
                    dev.ProductNumber = curROw.GetCell(3).ToString();
                    dev.ProductDescription = curROw.GetCell(4).ToString();
                    dev.Factory = curROw.GetCell(5).ToString();
                    dev.Phase = curROw.GetCell(6).ToString();
                    dev.Buliding = curROw.GetCell(7).ToString();
                    dev.Floor = curROw.GetCell(8).ToString();
                    dev.ClassNO = curROw.GetCell(9).ToString();
                    dev.DeviceClass = curROw.GetCell(10).ToString();
                    dev.Brand = curROw.GetCell(11).ToString();
                    dev.Type = curROw.GetCell(12).ToString();
                    dev.ProductClassful = curROw.GetCell(13).ToString();
                    dev.RackPeople = curROw.GetCell(14).ToString();
                    dev.Remarks = curROw.GetCell(15).ToString();
                    dev.AddDatetime = curROw.GetCell(16).ToString();
                    dev.Machinestate = curROw.GetCell(17).ToString();
                    dev.DownDate = curROw.GetCell(18).ToString();
                    string insertStr = "insert into SMSAlarmSystem(RemoveClass,System,SonSystem,ProductNumber,ProductDescription,Factory,Phase,Buliding,Floor,ClassNO,DeviceClass,Brand,Type,ProductClassful,RackPeople,Remarks,AddDatetime,Machinestate,DownDate)" +
                "VALUES('" + dev.RemoveClass + "','" + dev.System + "','" + dev.SonSystem + "','" + dev.ProductNumber + "','" + dev.ProductDescription + "','" + dev.Factory + "','" + dev.Phase + "','"
                + dev.Buliding + "','" + dev.Floor + "','" + dev.ClassNO + "','" + dev.DeviceClass + "','" + dev.Brand + "','" + dev.Type + "','" + dev.ProductClassful + "','" + dev.RackPeople + "','" + dev.Remarks + "','" + dev.AddDatetime + "','" + dev.Machinestate + "','" + dev.DownDate + "')";
                    CommandQuerytosql(insertStr);
                }
            }
            return RedirectToAction("SMSAlarmSystem");

        }
        public ActionResult SMSAlarmSystemExcelExport(SMSCCTVModel ins)
        {
            //取得資料
            //var result = this.studentService.GetStudentDataList(request).ToList();
            String SqlselectDeviceText = "Select *from SMSAlarmSystem";
            //建立Excel
            HSSFWorkbook hssfworkbook = new HSSFWorkbook(); //建立活頁簿
            ISheet sheet = hssfworkbook.CreateSheet("資產"); //建立sheet叫資產

            //設定樣式
            ICellStyle headerStyle = hssfworkbook.CreateCellStyle();//標題樣式
            HSSFCellStyle CS = (HSSFCellStyle)hssfworkbook.CreateCellStyle();//內容樣式

            IFont headerfont = hssfworkbook.CreateFont();//標題文字樣式
            HSSFFont font = (HSSFFont)hssfworkbook.CreateFont();//內容文字樣式        
            font.FontName = "細明體";
            font.FontHeightInPoints = 12;
            headerStyle.Alignment = HorizontalAlignment.Left; //水平至左
            headerStyle.VerticalAlignment = VerticalAlignment.Center; //垂直置中
            //headerStyle.FillPattern = NPOI.SS.UserModel.FillPattern.SolidForeground;
            //headerStyle.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.Blue.Index;
            headerfont.FontName = "細明體";
            headerfont.FontHeightInPoints = 12;
            headerfont.Boldweight = (short)FontBoldWeight.Bold;
            headerStyle.SetFont(headerfont);//將表格內樣式套給字行
            CS.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.Blue.Index;
            CS.SetFont(font);
            //新增標題列
            sheet.CreateRow(0); //需先用CreateRow建立,才可通过GetRow取得該欄位


            //sheet.AddMergedRegion(new CellRangeAddress(0, 1, 0, 2)); //合併1~2列及A~C欄儲存格
            //sheet.GetRow(0).CreateCell(0).SetCellValue("昕力大學");

            sheet.CreateRow(0).CreateCell(0).SetCellValue("異動類別*");
            sheet.GetRow(0).CreateCell(1).SetCellValue("系統*");
            sheet.GetRow(0).CreateCell(2).SetCellValue("子系統*");
            sheet.GetRow(0).CreateCell(3).SetCellValue("資產編號*");
            sheet.GetRow(0).CreateCell(4).SetCellValue("資產描述*");
            sheet.GetRow(0).CreateCell(5).SetCellValue("廠區*");
            sheet.GetRow(0).CreateCell(6).SetCellValue("Phase*");
            sheet.GetRow(0).CreateCell(7).SetCellValue("棟別*");
            sheet.GetRow(0).CreateCell(8).SetCellValue("樓層*");
            sheet.GetRow(0).CreateCell(9).SetCellValue("課別*");
            sheet.GetRow(0).CreateCell(10).SetCellValue("設備分類*");
            sheet.GetRow(0).CreateCell(11).SetCellValue("廠牌*");
            sheet.GetRow(0).CreateCell(12).SetCellValue("型號*");
            sheet.GetRow(0).CreateCell(13).SetCellValue("資產分類*");
            sheet.GetRow(0).GetCell(0).CellStyle = headerStyle; //套用樣式
            sheet.GetRow(0).GetCell(1).CellStyle = headerStyle; //套用樣式
            sheet.GetRow(0).GetCell(2).CellStyle = headerStyle; //套用樣式
            sheet.GetRow(0).GetCell(3).CellStyle = headerStyle; //套用樣式
            sheet.GetRow(0).GetCell(4).CellStyle = headerStyle; //套用樣式
            sheet.GetRow(0).GetCell(5).CellStyle = headerStyle; //套用樣式
            sheet.GetRow(0).GetCell(6).CellStyle = headerStyle; //套用樣式
            sheet.GetRow(0).GetCell(7).CellStyle = headerStyle; //套用樣式
            sheet.GetRow(0).GetCell(8).CellStyle = headerStyle; //套用樣式
            sheet.GetRow(0).GetCell(9).CellStyle = headerStyle; //套用樣式
            sheet.GetRow(0).GetCell(10).CellStyle = headerStyle; //套用樣式
            sheet.GetRow(0).GetCell(11).CellStyle = headerStyle; //套用樣式
            sheet.GetRow(0).GetCell(12).CellStyle = headerStyle; //套用樣式
            sheet.GetRow(0).GetCell(13).CellStyle = headerStyle; //套用樣式

            DataTable dt = da(SqlselectDeviceText);
            //填入資料
            int rowIndex = 1;
            for (int row = 0; row < dt.Rows.Count; row++)
            {
                sheet.CreateRow(rowIndex).CreateCell(0).SetCellValue(dt.Rows[row]["RemoveClass"].ToString());
                sheet.GetRow(rowIndex).CreateCell(1).SetCellValue(dt.Rows[row]["System"].ToString());
                sheet.GetRow(rowIndex).CreateCell(2).SetCellValue(dt.Rows[row]["SonSystem"].ToString());
                sheet.GetRow(rowIndex).CreateCell(3).SetCellValue(dt.Rows[row]["ProductNumber"].ToString());
                sheet.GetRow(rowIndex).CreateCell(4).SetCellValue(dt.Rows[row]["ProductDescription"].ToString());
                sheet.GetRow(rowIndex).CreateCell(5).SetCellValue(dt.Rows[row]["Factory"].ToString());
                sheet.GetRow(rowIndex).CreateCell(6).SetCellValue(dt.Rows[row]["Phase"].ToString());
                sheet.GetRow(rowIndex).CreateCell(7).SetCellValue(dt.Rows[row]["Buliding"].ToString());
                sheet.GetRow(rowIndex).CreateCell(8).SetCellValue(dt.Rows[row]["Floor"].ToString());
                sheet.GetRow(rowIndex).CreateCell(9).SetCellValue("'" + dt.Rows[row]["ClassNO"].ToString());
                sheet.GetRow(rowIndex).CreateCell(10).SetCellValue(dt.Rows[row]["DeviceClass"].ToString());
                sheet.GetRow(rowIndex).CreateCell(11).SetCellValue(dt.Rows[row]["Brand"].ToString());
                sheet.GetRow(rowIndex).CreateCell(12).SetCellValue(dt.Rows[row]["Type"].ToString());
                if (dt.Rows[row]["ProductClassful"].ToString() == "N/A")
                {
                    sheet.GetRow(rowIndex).CreateCell(13).SetCellValue("");
                }
                else
                {
                    sheet.GetRow(rowIndex).CreateCell(13).SetCellValue(dt.Rows[row]["ProductClassful"].ToString());
                }
                sheet.GetRow(rowIndex).GetCell(0).CellStyle = CS; //套用樣式
                sheet.GetRow(rowIndex).GetCell(1).CellStyle = CS; //套用樣式
                sheet.GetRow(rowIndex).GetCell(2).CellStyle = CS; //套用樣式
                sheet.GetRow(rowIndex).GetCell(3).CellStyle = CS; //套用樣式
                sheet.GetRow(rowIndex).GetCell(4).CellStyle = CS; //套用樣式
                sheet.GetRow(rowIndex).GetCell(5).CellStyle = CS; //套用樣式
                sheet.GetRow(rowIndex).GetCell(6).CellStyle = CS; //套用樣式
                sheet.GetRow(rowIndex).GetCell(7).CellStyle = CS; //套用樣式
                sheet.GetRow(rowIndex).GetCell(8).CellStyle = CS; //套用樣式
                sheet.GetRow(rowIndex).GetCell(9).CellStyle = CS; //套用樣式
                sheet.GetRow(rowIndex).GetCell(10).CellStyle = CS; //套用樣式
                sheet.GetRow(rowIndex).GetCell(11).CellStyle = CS; //套用樣式
                sheet.GetRow(rowIndex).GetCell(12).CellStyle = CS; //套用樣式
                sheet.GetRow(rowIndex).GetCell(13).CellStyle = CS; //套用樣式
                sheet.SetColumnWidth(rowIndex - 1, 5000);
                rowIndex++;
            }
            var excelDatas = new MemoryStream();
            hssfworkbook.Write(excelDatas);
            return File(excelDatas.ToArray(), "application/vnd.ms-excel", string.Format($"FMCS_SMS-AlarmSystem_資產.xls"));
        }
        public ActionResult ExampleSMSAlarmSystem()
        {
            string path = Server.MapPath("~/Upload/SMS-AlarmSystem_範本.xls");
            string filenamepath = System.IO.Path.GetFileName(path);
            FileStream stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
            return File(stream, "application/octet-stream", filenamepath);
        }
        [HttpPost]
        public ActionResult SMSAlarmSystemEdit(SMSCCTVModel edit)
        {
            if (edit.RackPeople == Session["Remarks"].ToString())
            {
                string SqlEdit = string.Format("UPDATE SMSAlarmSystem SET RemoveClass ='" + edit.RemoveClass + "',System = '" + edit.System + "',SonSystem ='" + edit.SonSystem + "',ProductNumber = '" + edit.ProductNumber + "',ProductDescription = '" + edit.ProductDescription + "',Factory = '" +
                edit.Factory + "',Phase ='" + edit.Phase + "',Buliding ='" + edit.Buliding + "',Floor = '" + edit.Floor + "',ClassNO = '" + edit.ClassNO + "',DeviceClass = '" + edit.DeviceClass + "',Brand = '" + edit.Brand + "',Type = '" + edit.Type + "',ProductClassful = '" + edit.ProductClassful + "',Remarks = '" + edit.Remarks + "',AddDatetime ='" + edit.AddDatetime + "',Machinestate ='" + edit.Machinestate + "',DownDate ='" + edit.DownDate + "',EditDatetime = GETDATE()" + " WHERE ID = " + edit.ID);
                CommandQuerytosql(SqlEdit);
                return RedirectToAction("SMSAlarmSystem");
            }
            else
            {
                return RedirectToAction("NoAuthoritySMSCCTV", "Home");
            }
        }
        public ActionResult SMSAlarmSystemEdit(SMSCCTVModel DCTEdit, string id)
        {

            id = DCTEdit.ID;
            string sqlEditselect = string.Format("Select *from SMSAlarmSystem WHERE ID = " + DCTEdit.ID);
            DataTable dt = da(sqlEditselect);
            return View(dt);

        }
        public ActionResult SMSAlarmSystemDelete(SMSCCTVModel insc)
        {
            string Deletestr = string.Format("Delete from SMSAlarmSystem WHERE ID =" + insc.ID.Replace("'", "''"));
            CommandQuerytosql(Deletestr);
            return RedirectToAction("SMSAlarmSystem", "FAM");
        }
        public ActionResult SMSBarriergate(int? page, int? PageSize, string searchBarriergateNumber)
        {
            List<SMSCCTVModel> list = new List<SMSCCTVModel>();//SMSCCTVModel的泛型空間
            string strCCTV = "";
            ViewBag.PageSize = new List<SelectListItem>()//數值放入ViewBag暫存
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
            if (!string.IsNullOrEmpty(searchBarriergateNumber))
            {
                strCCTV = "Select *from SMSBarriergate WHERE ProductNumber like '%" + searchBarriergateNumber + "%'";
            }
            else
            {
                strCCTV = "select *from SMSBarriergate order by ID DESC";
            }
            ReturnPermit ret = new ReturnPermit();
            String SMSCCTV1 = ret.returnSMSBarriergate(Session["Permit"].ToString(), sql);//取得驗證權限的資訊
            if (SMSCCTV1 == "授權")
            {


                DataTable dt = da(strCCTV);
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    SMSCCTVModel ins = new SMSCCTVModel();
                    ins.ID = dt.Rows[i]["ID"].ToString();
                    ins.RemoveClass = dt.Rows[i]["RemoveClass"].ToString();
                    ins.System = dt.Rows[i]["System"].ToString();
                    ins.SonSystem = dt.Rows[i]["SonSystem"].ToString();
                    ins.ProductNumber = dt.Rows[i]["ProductNumber"].ToString();
                    ins.ProductDescription = dt.Rows[i]["ProductDescription"].ToString();
                    ins.Factory = dt.Rows[i]["Factory"].ToString();
                    ins.Phase = dt.Rows[i]["Phase"].ToString();
                    ins.Buliding = dt.Rows[i]["Buliding"].ToString();
                    ins.Floor = dt.Rows[i]["Floor"].ToString();
                    ins.ClassNO = dt.Rows[i]["ClassNO"].ToString();
                    ins.DeviceClass = dt.Rows[i]["DeviceClass"].ToString();
                    ins.Brand = dt.Rows[i]["Brand"].ToString();
                    ins.Type = dt.Rows[i]["Type"].ToString();
                    ins.ProductClassful = dt.Rows[i]["ProductClassful"].ToString();
                    ins.RackPeople = dt.Rows[i]["RackPeople"].ToString();
                    ins.Remarks = dt.Rows[i]["Remarks"].ToString();
                    ins.AddDatetime = dt.Rows[i]["AddDatetime"].ToString();
                    ins.Machinestate = dt.Rows[i]["Machinestate"].ToString();
                    ins.DownDate = dt.Rows[i]["DownDate"].ToString();
                    ins.EditDatetime = dt.Rows[i]["EditDatetime"].ToString();
                    list.Add(ins);
                }
                var pagelist = list.ToPagedList(pageNumber, pagesize);
                return View(pagelist);
            }
            else
            {
                return RedirectToAction("NoAuthorityView", "Home");
            }

        }
        [HttpPost]
        public ActionResult InsertSMSBarriergate(SMSCCTVModel insc)//新增SMSCCTV的資訊
        {
            sc.ConnectionString = sql;
            string insertStr = "insert into SMSBarriergate(RemoveClass,System,SonSystem,ProductNumber,ProductDescription,Factory,Phase,Buliding,Floor,ClassNO,DeviceClass,Brand,Type,ProductClassful,RackPeople,Remarks,AddDatetime,Machinestate,DownDate)" +
                "VALUES('" + insc.RemoveClass + "','" + insc.System + "','" + insc.SonSystem + "','" + insc.ProductNumber + "','" + insc.ProductDescription + "','" + insc.Factory + "','" + insc.Phase + "','"
                + insc.Buliding + "','" + insc.Floor + "','" + insc.ClassNO + "','" + insc.DeviceClass + "','" + insc.Brand + "','" + insc.Type + "','" + insc.ProductClassful + "','" + insc.RackPeople + "','" + insc.Remarks + "','" + insc.AddDatetime + "','" + insc.Machinestate + "','" + insc.DownDate + "')";
            //string.Format("insert into DeviceSwitchPort(ID,SwitchName,SwitchPort,DeviceType,AddDatetime,EditDatetime" +
            //")VALUES(" + ins.ID + "," + ins.SwitchName.Replace("'","''") + "," + ins.SwitchPort.Replace("'","''") + "," + ins.DeviceType.Replace("'","''")
            //+ "," + ins.AddDatetime.Replace("'","''") + "," + ins.EditDatetime.Replace("'","''") + ")"
            //);
            CommandQuerytosql(insertStr);
            return RedirectToAction("SMSBarriergate");

        }
        public ActionResult InsertSMSBarriergate(FabModel fab)
        {
            string selectFab = string.Format("Select* from Fab");
            DataTable dt = da(selectFab);
            return View(dt);
        }
        public ActionResult SMSBarriergateUpload()
        {
            return View();
        }
        [HttpPost]
        public ActionResult SMSBarriergateUpload(SMSCCTVModel dev, HttpPostedFileBase file)
        {
            string extension =
                    System.IO.Path.GetExtension(file.FileName);

            if (extension == ".xls" || extension == ".xlsx")
            {
                string fileLocation = Server.MapPath("~/Upload/") + file.FileName;
                file.SaveAs(fileLocation); // 存放檔案到伺服器上               
            }
            //String filePath = @"D:\20201205廠區寫的Web\FactoryData - 複製\FactoryData\Upload\" + file.FileName;//測試用
            String filePath = @"C:\20201218廠區資料\20201218場內寫的WEB\20201205廠區寫的Web\FactoryData - 複製\FactoryData\Upload\" + file.FileName;//測試用

            //String filePath = @"D:\NewWeb\Upload\" + file.FileName;

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
                    dev.RemoveClass = curROw.GetCell(0).ToString();
                    dev.System = curROw.GetCell(1).ToString();
                    dev.SonSystem = curROw.GetCell(2).ToString();
                    dev.ProductNumber = curROw.GetCell(3).ToString();
                    dev.ProductDescription = curROw.GetCell(4).ToString();
                    dev.Factory = curROw.GetCell(5).ToString();
                    dev.Phase = curROw.GetCell(6).ToString();
                    dev.Buliding = curROw.GetCell(7).ToString();
                    dev.Floor = curROw.GetCell(8).ToString();
                    dev.ClassNO = curROw.GetCell(9).ToString();
                    dev.DeviceClass = curROw.GetCell(10).ToString();
                    dev.Brand = curROw.GetCell(11).ToString();
                    dev.Type = curROw.GetCell(12).ToString();
                    dev.ProductClassful = curROw.GetCell(13).ToString();
                    dev.RackPeople = curROw.GetCell(14).ToString();
                    dev.Remarks = curROw.GetCell(15).ToString();
                    dev.AddDatetime = curROw.GetCell(16).ToString();
                    dev.Machinestate = curROw.GetCell(17).ToString();
                    dev.DownDate = curROw.GetCell(18).ToString();
                    string insertStr = "insert into SMSBarriergate(RemoveClass,System,SonSystem,ProductNumber,ProductDescription,Factory,Phase,Buliding,Floor,ClassNO,DeviceClass,Brand,Type,ProductClassful,RackPeople,Remarks,AddDatetime,Machinestate,DownDate)" +
                "VALUES('" + dev.RemoveClass + "','" + dev.System + "','" + dev.SonSystem + "','" + dev.ProductNumber + "','" + dev.ProductDescription + "','" + dev.Factory + "','" + dev.Phase + "','"
                + dev.Buliding + "','" + dev.Floor + "','" + dev.ClassNO + "','" + dev.DeviceClass + "','" + dev.Brand + "','" + dev.Type + "','" + dev.ProductClassful + "','" + dev.RackPeople + "','" + dev.Remarks + "','" + dev.AddDatetime + "','" + dev.Machinestate + "','" + dev.DownDate + "')";
                    CommandQuerytosql(insertStr);
                }
            }
            return RedirectToAction("SMSBarriergate");

        }
        public ActionResult SMSBarriergateExcelExport(SMSCCTVModel ins)
        {
            //取得資料
            //var result = this.studentService.GetStudentDataList(request).ToList();
            String SqlselectDeviceText = "Select *from SMSBarriergate";
            //建立Excel
            HSSFWorkbook hssfworkbook = new HSSFWorkbook(); //建立活頁簿
            ISheet sheet = hssfworkbook.CreateSheet("資產"); //建立sheet叫資產

            //設定樣式
            ICellStyle headerStyle = hssfworkbook.CreateCellStyle();//標題樣式
            HSSFCellStyle CS = (HSSFCellStyle)hssfworkbook.CreateCellStyle();//內容樣式

            IFont headerfont = hssfworkbook.CreateFont();//標題文字樣式
            HSSFFont font = (HSSFFont)hssfworkbook.CreateFont();//內容文字樣式        
            font.FontName = "細明體";
            font.FontHeightInPoints = 12;
            headerStyle.Alignment = HorizontalAlignment.Left; //水平至左
            headerStyle.VerticalAlignment = VerticalAlignment.Center; //垂直置中
            //headerStyle.FillPattern = NPOI.SS.UserModel.FillPattern.SolidForeground;
            //headerStyle.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.Blue.Index;
            headerfont.FontName = "細明體";
            headerfont.FontHeightInPoints = 12;
            headerfont.Boldweight = (short)FontBoldWeight.Bold;
            headerStyle.SetFont(headerfont);//將表格內樣式套給字行
            CS.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.Blue.Index;
            CS.SetFont(font);
            //新增標題列
            sheet.CreateRow(0); //需先用CreateRow建立,才可通过GetRow取得該欄位


            //sheet.AddMergedRegion(new CellRangeAddress(0, 1, 0, 2)); //合併1~2列及A~C欄儲存格
            //sheet.GetRow(0).CreateCell(0).SetCellValue("昕力大學");

            sheet.CreateRow(0).CreateCell(0).SetCellValue("異動類別*");
            sheet.GetRow(0).CreateCell(1).SetCellValue("系統*");
            sheet.GetRow(0).CreateCell(2).SetCellValue("子系統*");
            sheet.GetRow(0).CreateCell(3).SetCellValue("資產編號*");
            sheet.GetRow(0).CreateCell(4).SetCellValue("資產描述*");
            sheet.GetRow(0).CreateCell(5).SetCellValue("廠區*");
            sheet.GetRow(0).CreateCell(6).SetCellValue("Phase*");
            sheet.GetRow(0).CreateCell(7).SetCellValue("棟別*");
            sheet.GetRow(0).CreateCell(8).SetCellValue("樓層*");
            sheet.GetRow(0).CreateCell(9).SetCellValue("課別*");
            sheet.GetRow(0).CreateCell(10).SetCellValue("設備分類*");
            sheet.GetRow(0).CreateCell(11).SetCellValue("廠牌*");
            sheet.GetRow(0).CreateCell(12).SetCellValue("型號*");
            sheet.GetRow(0).CreateCell(13).SetCellValue("資產分類*");
            sheet.GetRow(0).GetCell(0).CellStyle = headerStyle; //套用樣式
            sheet.GetRow(0).GetCell(1).CellStyle = headerStyle; //套用樣式
            sheet.GetRow(0).GetCell(2).CellStyle = headerStyle; //套用樣式
            sheet.GetRow(0).GetCell(3).CellStyle = headerStyle; //套用樣式
            sheet.GetRow(0).GetCell(4).CellStyle = headerStyle; //套用樣式
            sheet.GetRow(0).GetCell(5).CellStyle = headerStyle; //套用樣式
            sheet.GetRow(0).GetCell(6).CellStyle = headerStyle; //套用樣式
            sheet.GetRow(0).GetCell(7).CellStyle = headerStyle; //套用樣式
            sheet.GetRow(0).GetCell(8).CellStyle = headerStyle; //套用樣式
            sheet.GetRow(0).GetCell(9).CellStyle = headerStyle; //套用樣式
            sheet.GetRow(0).GetCell(10).CellStyle = headerStyle; //套用樣式
            sheet.GetRow(0).GetCell(11).CellStyle = headerStyle; //套用樣式
            sheet.GetRow(0).GetCell(12).CellStyle = headerStyle; //套用樣式
            sheet.GetRow(0).GetCell(13).CellStyle = headerStyle; //套用樣式

            DataTable dt = da(SqlselectDeviceText);
            //填入資料
            int rowIndex = 1;
            for (int row = 0; row < dt.Rows.Count; row++)
            {
                sheet.CreateRow(rowIndex).CreateCell(0).SetCellValue(dt.Rows[row]["RemoveClass"].ToString());
                sheet.GetRow(rowIndex).CreateCell(1).SetCellValue(dt.Rows[row]["System"].ToString());
                sheet.GetRow(rowIndex).CreateCell(2).SetCellValue(dt.Rows[row]["SonSystem"].ToString());
                sheet.GetRow(rowIndex).CreateCell(3).SetCellValue(dt.Rows[row]["ProductNumber"].ToString());
                sheet.GetRow(rowIndex).CreateCell(4).SetCellValue(dt.Rows[row]["ProductDescription"].ToString());
                sheet.GetRow(rowIndex).CreateCell(5).SetCellValue(dt.Rows[row]["Factory"].ToString());
                sheet.GetRow(rowIndex).CreateCell(6).SetCellValue(dt.Rows[row]["Phase"].ToString());
                sheet.GetRow(rowIndex).CreateCell(7).SetCellValue(dt.Rows[row]["Buliding"].ToString());
                sheet.GetRow(rowIndex).CreateCell(8).SetCellValue(dt.Rows[row]["Floor"].ToString());
                sheet.GetRow(rowIndex).CreateCell(9).SetCellValue("'" + dt.Rows[row]["ClassNO"].ToString());
                sheet.GetRow(rowIndex).CreateCell(10).SetCellValue(dt.Rows[row]["DeviceClass"].ToString());
                sheet.GetRow(rowIndex).CreateCell(11).SetCellValue(dt.Rows[row]["Brand"].ToString());
                sheet.GetRow(rowIndex).CreateCell(12).SetCellValue(dt.Rows[row]["Type"].ToString());
                if (dt.Rows[row]["ProductClassful"].ToString() == "N/A")
                {
                    sheet.GetRow(rowIndex).CreateCell(13).SetCellValue("");
                }
                else
                {
                    sheet.GetRow(rowIndex).CreateCell(13).SetCellValue(dt.Rows[row]["ProductClassful"].ToString());
                }
                sheet.GetRow(rowIndex).GetCell(0).CellStyle = CS; //套用樣式
                sheet.GetRow(rowIndex).GetCell(1).CellStyle = CS; //套用樣式
                sheet.GetRow(rowIndex).GetCell(2).CellStyle = CS; //套用樣式
                sheet.GetRow(rowIndex).GetCell(3).CellStyle = CS; //套用樣式
                sheet.GetRow(rowIndex).GetCell(4).CellStyle = CS; //套用樣式
                sheet.GetRow(rowIndex).GetCell(5).CellStyle = CS; //套用樣式
                sheet.GetRow(rowIndex).GetCell(6).CellStyle = CS; //套用樣式
                sheet.GetRow(rowIndex).GetCell(7).CellStyle = CS; //套用樣式
                sheet.GetRow(rowIndex).GetCell(8).CellStyle = CS; //套用樣式
                sheet.GetRow(rowIndex).GetCell(9).CellStyle = CS; //套用樣式
                sheet.GetRow(rowIndex).GetCell(10).CellStyle = CS; //套用樣式
                sheet.GetRow(rowIndex).GetCell(11).CellStyle = CS; //套用樣式
                sheet.GetRow(rowIndex).GetCell(12).CellStyle = CS; //套用樣式
                sheet.GetRow(rowIndex).GetCell(13).CellStyle = CS; //套用樣式
                sheet.SetColumnWidth(rowIndex - 1, 5000);
                rowIndex++;
            }
            var excelDatas = new MemoryStream();
            hssfworkbook.Write(excelDatas);
            return File(excelDatas.ToArray(), "application/vnd.ms-excel", string.Format($"FMCS_SMS-柵欄機_資產.xls"));
        }
        public ActionResult ExampleSMSBarriergate()
        {
            string path = Server.MapPath("~/Upload/SMS-柵欄機_範本.xls");
            string filenamepath = System.IO.Path.GetFileName(path);
            FileStream stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
            return File(stream, "application/octet-stream", filenamepath);
        }
        [HttpPost]
        public ActionResult SMSBarriergateEdit(SMSCCTVModel edit)
        {
            if (edit.RackPeople == Session["Remarks"].ToString())
            {
                string SqlEdit = string.Format("UPDATE SMSBarriergate SET RemoveClass ='" + edit.RemoveClass + "',System = '" + edit.System + "',SonSystem ='" + edit.SonSystem + "',ProductNumber = '" + edit.ProductNumber + "',ProductDescription = '" + edit.ProductDescription + "',Factory = '" +
                edit.Factory + "',Phase ='" + edit.Phase + "',Buliding ='" + edit.Buliding + "',Floor = '" + edit.Floor + "',ClassNO = '" + edit.ClassNO + "',DeviceClass = '" + edit.DeviceClass + "',Brand = '" + edit.Brand + "',Type = '" + edit.Type + "',ProductClassful = '" + edit.ProductClassful + "',Remarks = '" + edit.Remarks + "',AddDatetime ='" + edit.AddDatetime + "',Machinestate ='" + edit.Machinestate + "',DownDate ='" + edit.DownDate + "',EditDatetime = GETDATE()" + " WHERE ID = " + edit.ID);
                CommandQuerytosql(SqlEdit);
                return RedirectToAction("SMSBarriergate");
            }
            else
            {
                return RedirectToAction("NoAuthoritySMSCCTV", "Home");
            }
        }
        public ActionResult SMSBarriergateEdit(SMSCCTVModel DCTEdit, string id)
        {

            id = DCTEdit.ID;
            string sqlEditselect = string.Format("Select *from SMSBarriergate WHERE ID = " + DCTEdit.ID);
            DataTable dt = da(sqlEditselect);
            return View(dt);

        }
        public ActionResult SMSBarriergateDelete(SMSCCTVModel insc)
        {
            string Deletestr = string.Format("Delete from SMSBarriergate WHERE ID =" + insc.ID.Replace("'", "''"));
            CommandQuerytosql(Deletestr);
            return RedirectToAction("SMSBarriergate", "FAM");
        }
        public ActionResult SMSGate(int? page, int? PageSize, string searchGateNumber)
        {
            List<SMSCCTVModel> list = new List<SMSCCTVModel>();//SMSCCTVModel的泛型空間
            string strCCTV = "";
            ViewBag.PageSize = new List<SelectListItem>()//數值放入ViewBag暫存
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
            if (!string.IsNullOrEmpty(searchGateNumber))
            {
                strCCTV = "Select *from SMSGate WHERE ProductNumber like '%" + searchGateNumber + "%'";
            }
            else
            {
                strCCTV = "select *from SMSGate order by ID DESC";
            }
            ReturnPermit ret = new ReturnPermit();
            String SMSCCTV1 = ret.returnSMSGate(Session["Permit"].ToString(), sql);//取得驗證權限的資訊
            if (SMSCCTV1 == "授權")
            {


                DataTable dt = da(strCCTV);
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    SMSCCTVModel ins = new SMSCCTVModel();
                    ins.ID = dt.Rows[i]["ID"].ToString();
                    ins.RemoveClass = dt.Rows[i]["RemoveClass"].ToString();
                    ins.System = dt.Rows[i]["System"].ToString();
                    ins.SonSystem = dt.Rows[i]["SonSystem"].ToString();
                    ins.ProductNumber = dt.Rows[i]["ProductNumber"].ToString();
                    ins.ProductDescription = dt.Rows[i]["ProductDescription"].ToString();
                    ins.Factory = dt.Rows[i]["Factory"].ToString();
                    ins.Phase = dt.Rows[i]["Phase"].ToString();
                    ins.Buliding = dt.Rows[i]["Buliding"].ToString();
                    ins.Floor = dt.Rows[i]["Floor"].ToString();
                    ins.ClassNO = dt.Rows[i]["ClassNO"].ToString();
                    ins.DeviceClass = dt.Rows[i]["DeviceClass"].ToString();
                    ins.Brand = dt.Rows[i]["Brand"].ToString();
                    ins.Type = dt.Rows[i]["Type"].ToString();
                    ins.ProductClassful = dt.Rows[i]["ProductClassful"].ToString();
                    ins.RackPeople = dt.Rows[i]["RackPeople"].ToString();
                    ins.Remarks = dt.Rows[i]["Remarks"].ToString();
                    ins.AddDatetime = dt.Rows[i]["AddDatetime"].ToString();
                    ins.Machinestate = dt.Rows[i]["Machinestate"].ToString();
                    ins.DownDate = dt.Rows[i]["DownDate"].ToString();
                    ins.EditDatetime = dt.Rows[i]["EditDatetime"].ToString();
                    list.Add(ins);
                }
                var pagelist = list.ToPagedList(pageNumber, pagesize);
                return View(pagelist);
            }
            else
            {
                return RedirectToAction("NoAuthorityView", "Home");
            }

        }
        [HttpPost]
        public ActionResult InsertSMSGate(SMSCCTVModel insc)//新增SMSCCTV的資訊
        {
            sc.ConnectionString = sql;
            string insertStr = "insert into SMSGate(RemoveClass,System,SonSystem,ProductNumber,ProductDescription,Factory,Phase,Buliding,Floor,ClassNO,DeviceClass,Brand,Type,ProductClassful,RackPeople,Remarks,AddDatetime,Machinestate,DownDate)" +
                "VALUES('" + insc.RemoveClass + "','" + insc.System + "','" + insc.SonSystem + "','" + insc.ProductNumber + "','" + insc.ProductDescription + "','" + insc.Factory + "','" + insc.Phase + "','"
                + insc.Buliding + "','" + insc.Floor + "','" + insc.ClassNO + "','" + insc.DeviceClass + "','" + insc.Brand + "','" + insc.Type + "','" + insc.ProductClassful + "','" + insc.RackPeople + "','" + insc.Remarks + "','" + insc.AddDatetime + "','" + insc.Machinestate + "','" + insc.DownDate + "')";
            //string.Format("insert into DeviceSwitchPort(ID,SwitchName,SwitchPort,DeviceType,AddDatetime,EditDatetime" +
            //")VALUES(" + ins.ID + "," + ins.SwitchName.Replace("'","''") + "," + ins.SwitchPort.Replace("'","''") + "," + ins.DeviceType.Replace("'","''")
            //+ "," + ins.AddDatetime.Replace("'","''") + "," + ins.EditDatetime.Replace("'","''") + ")"
            //);
            CommandQuerytosql(insertStr);
            return RedirectToAction("SMSGate");

        }
        public ActionResult InsertSMSGate(FabModel fab)
        {
            string selectFab = string.Format("Select* from Fab");
            DataTable dt = da(selectFab);
            return View(dt);
        }
        public ActionResult SMSGateUpload()
        {
            return View();
        }
        [HttpPost]
        public ActionResult SMSGateUpload(SMSCCTVModel dev, HttpPostedFileBase file)
        {
            string extension =
                    System.IO.Path.GetExtension(file.FileName);

            if (extension == ".xls" || extension == ".xlsx")
            {
                string fileLocation = Server.MapPath("~/Upload/") + file.FileName;
                file.SaveAs(fileLocation); // 存放檔案到伺服器上               
            }
            //String filePath = @"D:\20201205廠區寫的Web\FactoryData - 複製\FactoryData\Upload\" + file.FileName;//測試用
            String filePath = @"C:\20201218廠區資料\20201218場內寫的WEB\20201205廠區寫的Web\FactoryData - 複製\FactoryData\Upload\" + file.FileName;//測試用

            //String filePath = @"D:\NewWeb\Upload\" + file.FileName;

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
                    dev.RemoveClass = curROw.GetCell(0).ToString();
                    dev.System = curROw.GetCell(1).ToString();
                    dev.SonSystem = curROw.GetCell(2).ToString();
                    dev.ProductNumber = curROw.GetCell(3).ToString();
                    dev.ProductDescription = curROw.GetCell(4).ToString();
                    dev.Factory = curROw.GetCell(5).ToString();
                    dev.Phase = curROw.GetCell(6).ToString();
                    dev.Buliding = curROw.GetCell(7).ToString();
                    dev.Floor = curROw.GetCell(8).ToString();
                    dev.ClassNO = curROw.GetCell(9).ToString();
                    dev.DeviceClass = curROw.GetCell(10).ToString();
                    dev.Brand = curROw.GetCell(11).ToString();
                    dev.Type = curROw.GetCell(12).ToString();
                    dev.ProductClassful = curROw.GetCell(13).ToString();
                    dev.RackPeople = curROw.GetCell(14).ToString();
                    dev.Remarks = curROw.GetCell(15).ToString();
                    dev.AddDatetime = curROw.GetCell(16).ToString();
                    dev.Machinestate = curROw.GetCell(17).ToString();
                    dev.DownDate = curROw.GetCell(18).ToString();
                    string insertStr = "insert into SMSGate(RemoveClass,System,SonSystem,ProductNumber,ProductDescription,Factory,Phase,Buliding,Floor,ClassNO,DeviceClass,Brand,Type,ProductClassful,RackPeople,Remarks,AddDatetime,Machinestate,DownDate)" +
                "VALUES('" + dev.RemoveClass + "','" + dev.System + "','" + dev.SonSystem + "','" + dev.ProductNumber + "','" + dev.ProductDescription + "','" + dev.Factory + "','" + dev.Phase + "','"
                + dev.Buliding + "','" + dev.Floor + "','" + dev.ClassNO + "','" + dev.DeviceClass + "','" + dev.Brand + "','" + dev.Type + "','" + dev.ProductClassful + "','" + dev.RackPeople + "','" + dev.Remarks + "','" + dev.AddDatetime + "','" + dev.Machinestate + "','" + dev.DownDate + "')";
                    CommandQuerytosql(insertStr);
                }
            }
            return RedirectToAction("SMSGate");

        }
        public ActionResult SMSGateExcelExport(SMSCCTVModel ins)
        {
            //取得資料
            //var result = this.studentService.GetStudentDataList(request).ToList();
            String SqlselectDeviceText = "Select *from SMSGate";
            //建立Excel
            HSSFWorkbook hssfworkbook = new HSSFWorkbook(); //建立活頁簿
            ISheet sheet = hssfworkbook.CreateSheet("資產"); //建立sheet叫資產

            //設定樣式
            ICellStyle headerStyle = hssfworkbook.CreateCellStyle();//標題樣式
            HSSFCellStyle CS = (HSSFCellStyle)hssfworkbook.CreateCellStyle();//內容樣式

            IFont headerfont = hssfworkbook.CreateFont();//標題文字樣式
            HSSFFont font = (HSSFFont)hssfworkbook.CreateFont();//內容文字樣式        
            font.FontName = "細明體";
            font.FontHeightInPoints = 12;
            headerStyle.Alignment = HorizontalAlignment.Left; //水平至左
            headerStyle.VerticalAlignment = VerticalAlignment.Center; //垂直置中
            //headerStyle.FillPattern = NPOI.SS.UserModel.FillPattern.SolidForeground;
            //headerStyle.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.Blue.Index;
            headerfont.FontName = "細明體";
            headerfont.FontHeightInPoints = 12;
            headerfont.Boldweight = (short)FontBoldWeight.Bold;
            headerStyle.SetFont(headerfont);//將表格內樣式套給字行
            CS.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.Blue.Index;
            CS.SetFont(font);
            //新增標題列
            sheet.CreateRow(0); //需先用CreateRow建立,才可通过GetRow取得該欄位


            //sheet.AddMergedRegion(new CellRangeAddress(0, 1, 0, 2)); //合併1~2列及A~C欄儲存格
            //sheet.GetRow(0).CreateCell(0).SetCellValue("昕力大學");

            sheet.CreateRow(0).CreateCell(0).SetCellValue("異動類別*");
            sheet.GetRow(0).CreateCell(1).SetCellValue("系統*");
            sheet.GetRow(0).CreateCell(2).SetCellValue("子系統*");
            sheet.GetRow(0).CreateCell(3).SetCellValue("資產編號*");
            sheet.GetRow(0).CreateCell(4).SetCellValue("資產描述*");
            sheet.GetRow(0).CreateCell(5).SetCellValue("廠區*");
            sheet.GetRow(0).CreateCell(6).SetCellValue("Phase*");
            sheet.GetRow(0).CreateCell(7).SetCellValue("棟別*");
            sheet.GetRow(0).CreateCell(8).SetCellValue("樓層*");
            sheet.GetRow(0).CreateCell(9).SetCellValue("課別*");
            sheet.GetRow(0).CreateCell(10).SetCellValue("設備分類*");
            sheet.GetRow(0).CreateCell(11).SetCellValue("廠牌*");
            sheet.GetRow(0).CreateCell(12).SetCellValue("型號*");
            sheet.GetRow(0).CreateCell(13).SetCellValue("資產分類*");
            sheet.GetRow(0).GetCell(0).CellStyle = headerStyle; //套用樣式
            sheet.GetRow(0).GetCell(1).CellStyle = headerStyle; //套用樣式
            sheet.GetRow(0).GetCell(2).CellStyle = headerStyle; //套用樣式
            sheet.GetRow(0).GetCell(3).CellStyle = headerStyle; //套用樣式
            sheet.GetRow(0).GetCell(4).CellStyle = headerStyle; //套用樣式
            sheet.GetRow(0).GetCell(5).CellStyle = headerStyle; //套用樣式
            sheet.GetRow(0).GetCell(6).CellStyle = headerStyle; //套用樣式
            sheet.GetRow(0).GetCell(7).CellStyle = headerStyle; //套用樣式
            sheet.GetRow(0).GetCell(8).CellStyle = headerStyle; //套用樣式
            sheet.GetRow(0).GetCell(9).CellStyle = headerStyle; //套用樣式
            sheet.GetRow(0).GetCell(10).CellStyle = headerStyle; //套用樣式
            sheet.GetRow(0).GetCell(11).CellStyle = headerStyle; //套用樣式
            sheet.GetRow(0).GetCell(12).CellStyle = headerStyle; //套用樣式
            sheet.GetRow(0).GetCell(13).CellStyle = headerStyle; //套用樣式

            DataTable dt = da(SqlselectDeviceText);
            //填入資料
            int rowIndex = 1;
            for (int row = 0; row < dt.Rows.Count; row++)
            {
                sheet.CreateRow(rowIndex).CreateCell(0).SetCellValue(dt.Rows[row]["RemoveClass"].ToString());
                sheet.GetRow(rowIndex).CreateCell(1).SetCellValue(dt.Rows[row]["System"].ToString());
                sheet.GetRow(rowIndex).CreateCell(2).SetCellValue(dt.Rows[row]["SonSystem"].ToString());
                sheet.GetRow(rowIndex).CreateCell(3).SetCellValue(dt.Rows[row]["ProductNumber"].ToString());
                sheet.GetRow(rowIndex).CreateCell(4).SetCellValue(dt.Rows[row]["ProductDescription"].ToString());
                sheet.GetRow(rowIndex).CreateCell(5).SetCellValue(dt.Rows[row]["Factory"].ToString());
                sheet.GetRow(rowIndex).CreateCell(6).SetCellValue(dt.Rows[row]["Phase"].ToString());
                sheet.GetRow(rowIndex).CreateCell(7).SetCellValue(dt.Rows[row]["Buliding"].ToString());
                sheet.GetRow(rowIndex).CreateCell(8).SetCellValue(dt.Rows[row]["Floor"].ToString());
                sheet.GetRow(rowIndex).CreateCell(9).SetCellValue("'" + dt.Rows[row]["ClassNO"].ToString());
                sheet.GetRow(rowIndex).CreateCell(10).SetCellValue(dt.Rows[row]["DeviceClass"].ToString());
                sheet.GetRow(rowIndex).CreateCell(11).SetCellValue(dt.Rows[row]["Brand"].ToString());
                sheet.GetRow(rowIndex).CreateCell(12).SetCellValue(dt.Rows[row]["Type"].ToString());
                if (dt.Rows[row]["ProductClassful"].ToString() == "N/A")
                {
                    sheet.GetRow(rowIndex).CreateCell(13).SetCellValue("");
                }
                else
                {
                    sheet.GetRow(rowIndex).CreateCell(13).SetCellValue(dt.Rows[row]["ProductClassful"].ToString());
                }
                sheet.GetRow(rowIndex).GetCell(0).CellStyle = CS; //套用樣式
                sheet.GetRow(rowIndex).GetCell(1).CellStyle = CS; //套用樣式
                sheet.GetRow(rowIndex).GetCell(2).CellStyle = CS; //套用樣式
                sheet.GetRow(rowIndex).GetCell(3).CellStyle = CS; //套用樣式
                sheet.GetRow(rowIndex).GetCell(4).CellStyle = CS; //套用樣式
                sheet.GetRow(rowIndex).GetCell(5).CellStyle = CS; //套用樣式
                sheet.GetRow(rowIndex).GetCell(6).CellStyle = CS; //套用樣式
                sheet.GetRow(rowIndex).GetCell(7).CellStyle = CS; //套用樣式
                sheet.GetRow(rowIndex).GetCell(8).CellStyle = CS; //套用樣式
                sheet.GetRow(rowIndex).GetCell(9).CellStyle = CS; //套用樣式
                sheet.GetRow(rowIndex).GetCell(10).CellStyle = CS; //套用樣式
                sheet.GetRow(rowIndex).GetCell(11).CellStyle = CS; //套用樣式
                sheet.GetRow(rowIndex).GetCell(12).CellStyle = CS; //套用樣式
                sheet.GetRow(rowIndex).GetCell(13).CellStyle = CS; //套用樣式
                sheet.SetColumnWidth(rowIndex - 1, 5000);
                rowIndex++;
            }
            var excelDatas = new MemoryStream();
            hssfworkbook.Write(excelDatas);
            return File(excelDatas.ToArray(), "application/vnd.ms-excel", string.Format($"FMCS_SMS-通關機_資產.xls"));
        }
        public ActionResult ExampleSMSGate()
        {
            string path = Server.MapPath("~/Upload/SMS-通關機_範本.xls");
            string filenamepath = System.IO.Path.GetFileName(path);
            FileStream stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
            return File(stream, "application/octet-stream", filenamepath);
        }
        [HttpPost]
        public ActionResult SMSGateEdit(SMSCCTVModel edit)
        {
            if (edit.RackPeople == Session["Remarks"].ToString())
            {
                string SqlEdit = string.Format("UPDATE SMSGate SET RemoveClass ='" + edit.RemoveClass + "',System = '" + edit.System + "',SonSystem ='" + edit.SonSystem + "',ProductNumber = '" + edit.ProductNumber + "',ProductDescription = '" + edit.ProductDescription + "',Factory = '" +
                edit.Factory + "',Phase ='" + edit.Phase + "',Buliding ='" + edit.Buliding + "',Floor = '" + edit.Floor + "',ClassNO = '" + edit.ClassNO + "',DeviceClass = '" + edit.DeviceClass + "',Brand = '" + edit.Brand + "',Type = '" + edit.Type + "',ProductClassful = '" + edit.ProductClassful + "',Remarks = '" + edit.Remarks + "',AddDatetime ='" + edit.AddDatetime + "',Machinestate ='" + edit.Machinestate + "',DownDate ='" + edit.DownDate + "',EditDatetime = GETDATE()" + " WHERE ID = " + edit.ID);
                CommandQuerytosql(SqlEdit);
                return RedirectToAction("SMSGate");
            }
            else
            {
                return RedirectToAction("NoAuthoritySMSCCTV", "Home");
            }
        }
        public ActionResult SMSGateEdit(SMSCCTVModel DCTEdit, string id)
        {

            id = DCTEdit.ID;
            string sqlEditselect = string.Format("Select *from SMSGate WHERE ID = " + DCTEdit.ID);
            DataTable dt = da(sqlEditselect);
            return View(dt);

        }
        public ActionResult SMSGateDelete(SMSCCTVModel insc)
        {
            string Deletestr = string.Format("Delete from SMSGate WHERE ID =" + insc.ID.Replace("'", "''"));
            CommandQuerytosql(Deletestr);
            return RedirectToAction("SMSGate", "FAM");
        }
        public ActionResult SMSMakeCard(int? page, int? PageSize, string searchMakeCardNumber)
        {
            List<SMSCCTVModel> list = new List<SMSCCTVModel>();//SMSCCTVModel的泛型空間
            string strCCTV = "";
            ViewBag.PageSize = new List<SelectListItem>()//數值放入ViewBag暫存
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
            if (!string.IsNullOrEmpty(searchMakeCardNumber))
            {
                strCCTV = "Select *from SMSMakeCard WHERE ProductNumber like '%" + searchMakeCardNumber + "%'";
            }
            else
            {
                strCCTV = "select *from SMSMakeCard order by ID DESC";
            }
            ReturnPermit ret = new ReturnPermit();
            String SMSCCTV1 = ret.returnSMSMakeCard(Session["Permit"].ToString(), sql);//取得驗證權限的資訊
            if (SMSCCTV1 == "授權")
            {


                DataTable dt = da(strCCTV);
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    SMSCCTVModel ins = new SMSCCTVModel();
                    ins.ID = dt.Rows[i]["ID"].ToString();
                    ins.RemoveClass = dt.Rows[i]["RemoveClass"].ToString();
                    ins.System = dt.Rows[i]["System"].ToString();
                    ins.SonSystem = dt.Rows[i]["SonSystem"].ToString();
                    ins.ProductNumber = dt.Rows[i]["ProductNumber"].ToString();
                    ins.ProductDescription = dt.Rows[i]["ProductDescription"].ToString();
                    ins.Factory = dt.Rows[i]["Factory"].ToString();
                    ins.Phase = dt.Rows[i]["Phase"].ToString();
                    ins.Buliding = dt.Rows[i]["Buliding"].ToString();
                    ins.Floor = dt.Rows[i]["Floor"].ToString();
                    ins.ClassNO = dt.Rows[i]["ClassNO"].ToString();
                    ins.DeviceClass = dt.Rows[i]["DeviceClass"].ToString();
                    ins.Brand = dt.Rows[i]["Brand"].ToString();
                    ins.Type = dt.Rows[i]["Type"].ToString();
                    ins.ProductClassful = dt.Rows[i]["ProductClassful"].ToString();
                    ins.RackPeople = dt.Rows[i]["RackPeople"].ToString();
                    ins.Remarks = dt.Rows[i]["Remarks"].ToString();
                    ins.AddDatetime = dt.Rows[i]["AddDatetime"].ToString();
                    ins.Machinestate = dt.Rows[i]["Machinestate"].ToString();
                    ins.DownDate = dt.Rows[i]["DownDate"].ToString();
                    ins.EditDatetime = dt.Rows[i]["EditDatetime"].ToString();
                    list.Add(ins);
                }
                var pagelist = list.ToPagedList(pageNumber, pagesize);
                return View(pagelist);
            }
            else
            {
                return RedirectToAction("NoAuthorityView", "Home");
            }

        }
        [HttpPost]
        public ActionResult InsertSMSMakeCard(SMSCCTVModel insc)//新增SMS-製證機的資訊
        {
            sc.ConnectionString = sql;
            string insertStr = "insert into SMSMakeCard(RemoveClass,System,SonSystem,ProductNumber,ProductDescription,Factory,Phase,Buliding,Floor,ClassNO,DeviceClass,Brand,Type,ProductClassful,RackPeople,Remarks,AddDatetime,Machinestate,DownDate)" +
                "VALUES('" + insc.RemoveClass + "','" + insc.System + "','" + insc.SonSystem + "','" + insc.ProductNumber + "','" + insc.ProductDescription + "','" + insc.Factory + "','" + insc.Phase + "','"
                + insc.Buliding + "','" + insc.Floor + "','" + insc.ClassNO + "','" + insc.DeviceClass + "','" + insc.Brand + "','" + insc.Type + "','" + insc.ProductClassful + "','" + insc.RackPeople + "','" + insc.Remarks + "','" + insc.AddDatetime + "','" + insc.Machinestate + "','" + insc.DownDate + "')";
            //string.Format("insert into DeviceSwitchPort(ID,SwitchName,SwitchPort,DeviceType,AddDatetime,EditDatetime" +
            //")VALUES(" + ins.ID + "," + ins.SwitchName.Replace("'","''") + "," + ins.SwitchPort.Replace("'","''") + "," + ins.DeviceType.Replace("'","''")
            //+ "," + ins.AddDatetime.Replace("'","''") + "," + ins.EditDatetime.Replace("'","''") + ")"
            //);
            CommandQuerytosql(insertStr);
            return RedirectToAction("SMSMakeCard");

        }
        public ActionResult InsertSMSMakeCard(FabModel fab)
        {
            string selectFab = string.Format("Select* from Fab");
            DataTable dt = da(selectFab);
            return View(dt);
        }
        public ActionResult SMSMakeCardUpload()
        {
            return View();
        }
        [HttpPost]
        public ActionResult SMSMakeCardUpload(SMSCCTVModel dev, HttpPostedFileBase file)
        {
            string extension =
                    System.IO.Path.GetExtension(file.FileName);

            if (extension == ".xls" || extension == ".xlsx")
            {
                string fileLocation = Server.MapPath("~/Upload/") + file.FileName;
                file.SaveAs(fileLocation); // 存放檔案到伺服器上               
            }
            //String filePath = @"D:\20201205廠區寫的Web\FactoryData - 複製\FactoryData\Upload\" + file.FileName;//測試用
            String filePath = @"C:\20201218廠區資料\20201218場內寫的WEB\20201205廠區寫的Web\FactoryData - 複製\FactoryData\Upload\" + file.FileName;//測試用

            //String filePath = @"D:\NewWeb\Upload\" + file.FileName;

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
                    dev.RemoveClass = curROw.GetCell(0).ToString();
                    dev.System = curROw.GetCell(1).ToString();
                    dev.SonSystem = curROw.GetCell(2).ToString();
                    dev.ProductNumber = curROw.GetCell(3).ToString();
                    dev.ProductDescription = curROw.GetCell(4).ToString();
                    dev.Factory = curROw.GetCell(5).ToString();
                    dev.Phase = curROw.GetCell(6).ToString();
                    dev.Buliding = curROw.GetCell(7).ToString();
                    dev.Floor = curROw.GetCell(8).ToString();
                    dev.ClassNO = curROw.GetCell(9).ToString();
                    dev.DeviceClass = curROw.GetCell(10).ToString();
                    dev.Brand = curROw.GetCell(11).ToString();
                    dev.Type = curROw.GetCell(12).ToString();
                    dev.ProductClassful = curROw.GetCell(13).ToString();
                    dev.RackPeople = curROw.GetCell(14).ToString();
                    dev.Remarks = curROw.GetCell(15).ToString();
                    dev.AddDatetime = curROw.GetCell(16).ToString();
                    dev.Machinestate = curROw.GetCell(17).ToString();
                    dev.DownDate = curROw.GetCell(18).ToString();
                    string insertStr = "insert into SMSMakeCard(RemoveClass,System,SonSystem,ProductNumber,ProductDescription,Factory,Phase,Buliding,Floor,ClassNO,DeviceClass,Brand,Type,ProductClassful,RackPeople,Remarks,AddDatetime,Machinestate,DownDate)" +
                "VALUES('" + dev.RemoveClass + "','" + dev.System + "','" + dev.SonSystem + "','" + dev.ProductNumber + "','" + dev.ProductDescription + "','" + dev.Factory + "','" + dev.Phase + "','"
                + dev.Buliding + "','" + dev.Floor + "','" + dev.ClassNO + "','" + dev.DeviceClass + "','" + dev.Brand + "','" + dev.Type + "','" + dev.ProductClassful + "','" + dev.RackPeople + "','" + dev.Remarks + "','" + dev.AddDatetime + "','" + dev.Machinestate + "','" + dev.DownDate + "')";
                    CommandQuerytosql(insertStr);
                }
            }
            return RedirectToAction("SMSMakeCard");

        }
        public ActionResult SMSMakeCardExcelExport(SMSCCTVModel ins)
        {
            //取得資料
            //var result = this.studentService.GetStudentDataList(request).ToList();
            String SqlselectDeviceText = "Select *from SMSMakeCard";
            //建立Excel
            HSSFWorkbook hssfworkbook = new HSSFWorkbook(); //建立活頁簿
            ISheet sheet = hssfworkbook.CreateSheet("資產"); //建立sheet叫資產

            //設定樣式
            ICellStyle headerStyle = hssfworkbook.CreateCellStyle();//標題樣式
            HSSFCellStyle CS = (HSSFCellStyle)hssfworkbook.CreateCellStyle();//內容樣式

            IFont headerfont = hssfworkbook.CreateFont();//標題文字樣式
            HSSFFont font = (HSSFFont)hssfworkbook.CreateFont();//內容文字樣式        
            font.FontName = "細明體";
            font.FontHeightInPoints = 12;
            headerStyle.Alignment = HorizontalAlignment.Left; //水平至左
            headerStyle.VerticalAlignment = VerticalAlignment.Center; //垂直置中
            //headerStyle.FillPattern = NPOI.SS.UserModel.FillPattern.SolidForeground;
            //headerStyle.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.Blue.Index;
            headerfont.FontName = "細明體";
            headerfont.FontHeightInPoints = 12;
            headerfont.Boldweight = (short)FontBoldWeight.Bold;
            headerStyle.SetFont(headerfont);//將表格內樣式套給字行
            CS.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.Blue.Index;
            CS.SetFont(font);
            //新增標題列
            sheet.CreateRow(0); //需先用CreateRow建立,才可通过GetRow取得該欄位


            //sheet.AddMergedRegion(new CellRangeAddress(0, 1, 0, 2)); //合併1~2列及A~C欄儲存格
            //sheet.GetRow(0).CreateCell(0).SetCellValue("昕力大學");

            sheet.CreateRow(0).CreateCell(0).SetCellValue("異動類別*");
            sheet.GetRow(0).CreateCell(1).SetCellValue("系統*");
            sheet.GetRow(0).CreateCell(2).SetCellValue("子系統*");
            sheet.GetRow(0).CreateCell(3).SetCellValue("資產編號*");
            sheet.GetRow(0).CreateCell(4).SetCellValue("資產描述*");
            sheet.GetRow(0).CreateCell(5).SetCellValue("廠區*");
            sheet.GetRow(0).CreateCell(6).SetCellValue("Phase*");
            sheet.GetRow(0).CreateCell(7).SetCellValue("棟別*");
            sheet.GetRow(0).CreateCell(8).SetCellValue("樓層*");
            sheet.GetRow(0).CreateCell(9).SetCellValue("課別*");
            sheet.GetRow(0).CreateCell(10).SetCellValue("設備分類*");
            sheet.GetRow(0).CreateCell(11).SetCellValue("廠牌*");
            sheet.GetRow(0).CreateCell(12).SetCellValue("型號*");
            sheet.GetRow(0).CreateCell(13).SetCellValue("資產分類*");
            sheet.GetRow(0).GetCell(0).CellStyle = headerStyle; //套用樣式
            sheet.GetRow(0).GetCell(1).CellStyle = headerStyle; //套用樣式
            sheet.GetRow(0).GetCell(2).CellStyle = headerStyle; //套用樣式
            sheet.GetRow(0).GetCell(3).CellStyle = headerStyle; //套用樣式
            sheet.GetRow(0).GetCell(4).CellStyle = headerStyle; //套用樣式
            sheet.GetRow(0).GetCell(5).CellStyle = headerStyle; //套用樣式
            sheet.GetRow(0).GetCell(6).CellStyle = headerStyle; //套用樣式
            sheet.GetRow(0).GetCell(7).CellStyle = headerStyle; //套用樣式
            sheet.GetRow(0).GetCell(8).CellStyle = headerStyle; //套用樣式
            sheet.GetRow(0).GetCell(9).CellStyle = headerStyle; //套用樣式
            sheet.GetRow(0).GetCell(10).CellStyle = headerStyle; //套用樣式
            sheet.GetRow(0).GetCell(11).CellStyle = headerStyle; //套用樣式
            sheet.GetRow(0).GetCell(12).CellStyle = headerStyle; //套用樣式
            sheet.GetRow(0).GetCell(13).CellStyle = headerStyle; //套用樣式

            DataTable dt = da(SqlselectDeviceText);
            //填入資料
            int rowIndex = 1;
            for (int row = 0; row < dt.Rows.Count; row++)
            {
                sheet.CreateRow(rowIndex).CreateCell(0).SetCellValue(dt.Rows[row]["RemoveClass"].ToString());
                sheet.GetRow(rowIndex).CreateCell(1).SetCellValue(dt.Rows[row]["System"].ToString());
                sheet.GetRow(rowIndex).CreateCell(2).SetCellValue(dt.Rows[row]["SonSystem"].ToString());
                sheet.GetRow(rowIndex).CreateCell(3).SetCellValue(dt.Rows[row]["ProductNumber"].ToString());
                sheet.GetRow(rowIndex).CreateCell(4).SetCellValue(dt.Rows[row]["ProductDescription"].ToString());
                sheet.GetRow(rowIndex).CreateCell(5).SetCellValue(dt.Rows[row]["Factory"].ToString());
                sheet.GetRow(rowIndex).CreateCell(6).SetCellValue(dt.Rows[row]["Phase"].ToString());
                sheet.GetRow(rowIndex).CreateCell(7).SetCellValue(dt.Rows[row]["Buliding"].ToString());
                sheet.GetRow(rowIndex).CreateCell(8).SetCellValue(dt.Rows[row]["Floor"].ToString());
                sheet.GetRow(rowIndex).CreateCell(9).SetCellValue("'" + dt.Rows[row]["ClassNO"].ToString());
                sheet.GetRow(rowIndex).CreateCell(10).SetCellValue(dt.Rows[row]["DeviceClass"].ToString());
                sheet.GetRow(rowIndex).CreateCell(11).SetCellValue(dt.Rows[row]["Brand"].ToString());
                sheet.GetRow(rowIndex).CreateCell(12).SetCellValue(dt.Rows[row]["Type"].ToString());
                if (dt.Rows[row]["ProductClassful"].ToString() == "N/A")
                {
                    sheet.GetRow(rowIndex).CreateCell(13).SetCellValue("");
                }
                else
                {
                    sheet.GetRow(rowIndex).CreateCell(13).SetCellValue(dt.Rows[row]["ProductClassful"].ToString());
                }
                sheet.GetRow(rowIndex).GetCell(0).CellStyle = CS; //套用樣式
                sheet.GetRow(rowIndex).GetCell(1).CellStyle = CS; //套用樣式
                sheet.GetRow(rowIndex).GetCell(2).CellStyle = CS; //套用樣式
                sheet.GetRow(rowIndex).GetCell(3).CellStyle = CS; //套用樣式
                sheet.GetRow(rowIndex).GetCell(4).CellStyle = CS; //套用樣式
                sheet.GetRow(rowIndex).GetCell(5).CellStyle = CS; //套用樣式
                sheet.GetRow(rowIndex).GetCell(6).CellStyle = CS; //套用樣式
                sheet.GetRow(rowIndex).GetCell(7).CellStyle = CS; //套用樣式
                sheet.GetRow(rowIndex).GetCell(8).CellStyle = CS; //套用樣式
                sheet.GetRow(rowIndex).GetCell(9).CellStyle = CS; //套用樣式
                sheet.GetRow(rowIndex).GetCell(10).CellStyle = CS; //套用樣式
                sheet.GetRow(rowIndex).GetCell(11).CellStyle = CS; //套用樣式
                sheet.GetRow(rowIndex).GetCell(12).CellStyle = CS; //套用樣式
                sheet.GetRow(rowIndex).GetCell(13).CellStyle = CS; //套用樣式
                sheet.SetColumnWidth(rowIndex - 1, 5000);
                rowIndex++;
            }
            var excelDatas = new MemoryStream();
            hssfworkbook.Write(excelDatas);
            return File(excelDatas.ToArray(), "application/vnd.ms-excel", string.Format($"FMCS_SMS-製證機_資產.xls"));
        }
        public ActionResult ExampleSMSMakeCard()
        {
            string path = Server.MapPath("~/Upload/SMS-製證機_範本.xls");
            string filenamepath = System.IO.Path.GetFileName(path);
            FileStream stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
            return File(stream, "application/octet-stream", filenamepath);
        }
        [HttpPost]
        public ActionResult SMSMakeCardEdit(SMSCCTVModel edit)
        {
            if (edit.RackPeople == Session["Remarks"].ToString())
            {
                string SqlEdit = string.Format("UPDATE SMSMakeCard SET RemoveClass ='" + edit.RemoveClass + "',System = '" + edit.System + "',SonSystem ='" + edit.SonSystem + "',ProductNumber = '" + edit.ProductNumber + "',ProductDescription = '" + edit.ProductDescription + "',Factory = '" +
                edit.Factory + "',Phase ='" + edit.Phase + "',Buliding ='" + edit.Buliding + "',Floor = '" + edit.Floor + "',ClassNO = '" + edit.ClassNO + "',DeviceClass = '" + edit.DeviceClass + "',Brand = '" + edit.Brand + "',Type = '" + edit.Type + "',ProductClassful = '" + edit.ProductClassful + "',Remarks = '" + edit.Remarks + "',AddDatetime ='" + edit.AddDatetime + "',Machinestate ='" + edit.Machinestate + "',DownDate ='" + edit.DownDate + "',EditDatetime = GETDATE()" + " WHERE ID = " + edit.ID);
                CommandQuerytosql(SqlEdit);
                return RedirectToAction("SMSGate");
            }
            else
            {
                return RedirectToAction("NoAuthoritySMSCCTV", "Home");
            }
        }
        public ActionResult SMSMakeCardEdit(SMSCCTVModel DCTEdit, string id)
        {

            id = DCTEdit.ID;
            string sqlEditselect = string.Format("Select *from SMSMakeCard WHERE ID = " + DCTEdit.ID);
            DataTable dt = da(sqlEditselect);
            return View(dt);

        }
        public ActionResult SMSMakeCardDelete(SMSCCTVModel insc)
        {
            string Deletestr = string.Format("Delete from SMSMakeCard WHERE ID =" + insc.ID.Replace("'", "''"));
            CommandQuerytosql(Deletestr);
            return RedirectToAction("SMSMakeCard", "FAM");
        }
    }
}
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Script.Serialization;
using System.Web.Mvc;
using Syncfusion.EJ.Export;
using Syncfusion.JavaScript.Models;
using Syncfusion.JavaScript.DataSources;
using Syncfusion.XlsIO;
using System.Collections;
using System.Reflection;
using ClosedXML.Excel;

namespace PicklistDemo.Controllers
{

    public class MultipleViewResult : ActionResult
    {
        public const string ChunkSeparator = "---|||---";
        public IList<PartialViewResult> PartialViewResults { get; private set; }
        

        public MultipleViewResult(params PartialViewResult[] views)
        {
            if (PartialViewResults == null)
                PartialViewResults = new List<PartialViewResult>();
            foreach (var v in views)
                PartialViewResults.Add(v);
        }

        public override void ExecuteResult(ControllerContext context)
        {
            if (context == null)
                throw new ArgumentNullException("context");
            var total = PartialViewResults.Count;
            for (var index = 0; index < total; index++)
            {
                var pv = PartialViewResults[index];
                pv.ExecuteResult(context);
                if (index < total - 1)
                    context.HttpContext.Response.Output.Write(ChunkSeparator);
            }
        }
    }
    public class DefaultController : Controller
    {
        DataTable dtExcelExport = new DataTable();
        DataSet dsPicklist = new DataSet();

        string firstPicklistDate;

        // GET: Default
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult ClosePicklist(string picklistDateToClose)
        {
            ClosePicklistDate(picklistDateToClose);
            return Redirect(Url.Content("~/")); 
        }

        public PartialViewResult DisplayResourceGroupsAndPriority(string strPickDate)
        {
            var cnSTR = ConfigurationManager.ConnectionStrings["LC_AppConnectionString"].ConnectionString;
            SqlConnection cn = new SqlConnection(cnSTR);
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = cn;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "TBEI_PicklistGetByDate";
            cmd.Parameters.Add("@startdate", SqlDbType.DateTime).Value = Convert.ToDateTime(strPickDate);

            SqlDataAdapter da = new SqlDataAdapter(cmd);
            DataSet ds = new DataSet();
            da.Fill(ds);

            DataTable dtPLGrps = new DataTable();
            dtPLGrps.Columns.AddRange(new DataColumn[2] {   new DataColumn("RGID", typeof(string)),
                                                         new DataColumn("Priority", typeof(string))});

            cmd.Connection = cn;
            cmd.CommandTimeout = 360;
            cmd.Parameters.Clear();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "TBEI_PicklistGetByResourceGroup";
            cmd.Parameters.Add("@startdate", SqlDbType.DateTime).Value = Convert.ToDateTime(strPickDate);
            cmd.Parameters.Add("@startresgrp", SqlDbType.NVarChar).Value = "";

            da.SelectCommand = cmd;
            da.Fill(dsPicklist);

            if (dsPicklist.Tables[0].Rows.Count == 0)
            {
                ClosePicklistDate(strPickDate);
                Session["PicklistDates"] = GetPicklistDates();
                Session["dataSourcePicklistDates"] = GetPicklistDates();
            }

            for (int x = 0; x <= ds.Tables[0].Rows.Count - 1; x++)
            {
                string strRGID = ds.Tables[0].Rows[x]["FirstGrp"].ToString();
                string strPriority = ds.Tables[0].Rows[x]["Priority"].ToString();

                string[] selectedColumns = new[] { "Item", "Description", "MatlDescription", "ItemQty", "LineLoc", "Priority", "FirstGrp", "StartDate" };
                DataTable dt = new DataView(dsPicklist.Tables[0]).ToTable(false, selectedColumns);

                int intDaysToSubtract = -1;
                int intDaysToAdd = 1;

                DateTime dtToday = Convert.ToDateTime(strPickDate);
                
                switch (dtToday.DayOfWeek)
                {
                    case DayOfWeek.Monday:
                        intDaysToSubtract = -4;
                        break;
                    case DayOfWeek.Friday:
                        intDaysToAdd = 3;
                        break;
                }

                // Get the string value of the selected date, selected date - 1 and selected data + 1
                DateTime dtSelectedDate = Convert.ToDateTime(strPickDate);
                DateTime dtSelectedDateMinus1 = dtSelectedDate.AddDays(intDaysToSubtract);

                string strPickDateMinus1 = dtSelectedDateMinus1.ToString("MM/dd/yyyy");
                DateTime dtSelectedDatePlus1 = dtSelectedDate.AddDays(intDaysToAdd);
                string strPickDatePlus1 = dtSelectedDatePlus1.ToString("MM/dd/yyyy");

                Session["PickDate"] = strPickDate;
                Session["PickDateMinus1"] = strPickDateMinus1;
                Session["PickDatePlus1"] = strPickDatePlus1;

                //DataRow[] drRows = dt.Select("StartDate = '" + strPickDate + "' AND Priority = '" + strPriority + "' AND FirstGrp = '" + strRGID + "'");
                DataRow[] drRows = dt.Select("Priority = '" + strPriority + "' AND FirstGrp = '" + strRGID + "'");
                //DataRow[] drRows = dt.Select("StartDate IN ('" + strPickDateMinus1 + "','" + strPickDate + "','" + strPickDatePlus1 + "') AND Priority = '" + strPriority + "' AND FirstGrp = '" + strRGID + "'");

                if (drRows.Count() > 0)
                {
                    dtPLGrps.Rows.Add(strRGID, strPriority);
                }
            }

            Session["dataSourcePicklistGroups"] = ConvertDataTableToJSON(dtPLGrps);
            Session["dsPicklist"] = (DataSet)dsPicklist;
            
            return PartialView("PicklistResourceGroups");
        }

        public PartialViewResult DisplayPicklists(string strRGID, string strPriority)
        {
            GetDataByLineLoc("1", strRGID, strPriority);
            GetDataByLineLoc("2", strRGID, strPriority);
            GetDataByLineLoc("3", strRGID, strPriority);
            GetDataByLineLoc("4", strRGID, strPriority);
            GetDataByLineLoc("5", strRGID, strPriority);
            GetDataByLineLoc("6", strRGID, strPriority);
            return PartialView("PicklistAssemblyLists");
        }

        public ActionResult Picklist(FormCollection form)
        {
            Session["displayTabPage"] = "0";
            Session["dataSourcePicklistDates"] = GetPicklistDates();

            if (form["resourceGroupText"] == null)
            {
                form["resourceGroupText"] = "";
                form["priorityText"] = "";
                form["strLocationType"] = "0";
                form["rgPriority"] = "";
                ViewBag.priorityText = "";
                ViewBag.resourceGroupText = "";
                Session["rgPriority"] = "";
                ViewBag.strLocationType = "0";
            }
            else
            {
                ViewBag.priorityText = form["priorityText"];
                ViewBag.resourceGroupText = form["resourceGroupText"];
                ViewBag.strLocationType = form["strLocationType"];
                Session["rgPriority"] = form["rgPriority"];
            }

            return View();
        }

        private static object ConvertDataTableToJSON(DataTable dt)
        {
            JavaScriptSerializer jSonString = new JavaScriptSerializer();
            List<Dictionary<string, object>> rows = new List<Dictionary<string, object>>();
            Dictionary<string, object> row;
            foreach (DataRow dr in dt.Rows)
            {
                row = new Dictionary<string, object>();
                foreach (DataColumn col in dt.Columns)
                {
                    row.Add(col.ColumnName, dr[col]);
                }
                rows.Add(row);
            }
            string serialize = jSonString.Serialize(rows);
            var data = jSonString.Deserialize<IEnumerable<object>>(serialize);
            return data;
        }

        public void ClosePicklistDate(string strPicklistDate)
        {
            var cnSTR = ConfigurationManager.ConnectionStrings["ShopfloorConnectionString"].ConnectionString;
            SqlConnection cn = new SqlConnection(cnSTR);
            cn.Open();

            SqlCommand cmd = new SqlCommand();
            cmd.Connection = cn;
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "UPDATE PickDate SET Active = 0 WHERE CONVERT(varchar,PickDate,101) = '" + strPicklistDate + "'";
            cmd.ExecuteNonQuery();
            cn.Close();
        }

        private List<object> GetPicklistDates()
        {
            DataTable dtDates = new DataTable();
            dtDates.Columns.AddRange(new DataColumn[1] { new DataColumn("SplitDate", typeof(string)) });
            dtDates.Rows.Add("-Select Pick Date-");

            var cnSTR = ConfigurationManager.ConnectionStrings["ShopfloorConnectionString"].ConnectionString;
            SqlConnection cn = new SqlConnection(cnSTR);
            cn.Open();

            SqlCommand cmd = new SqlCommand();
            cmd.Connection = cn;
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "SELECT CONVERT(varchar,PickDate,101) AS PickDate FROM PickDate WHERE Active = 1 ORDER BY PickDate";

            List<object> dropObj = new List<object>();

            using (SqlDataReader dr = cmd.ExecuteReader())
            {
                int x = 0;
                while (dr.Read())
                {
                    dropObj.Add(new
                    {
                        Text = dr.GetValue(0).ToString(),
                        Value = dr.GetValue(0).ToString()
                    });

                    if (x == 0)
                    {
                        firstPicklistDate = dr.GetValue(0).ToString();
                    }
                    x += 1;
                }
            }

            cn.Close();

            return dropObj;
        }

        private void GetDataByLineLoc(string strCnt, string strPriority, string strRGID)
        {
            DataSet dsPicklist = (DataSet)Session["dsPicklist"];
            string strNewRGID = "";

            if (strPriority == "") { strPriority = "0"; }
            if (strCnt == "") { strCnt = "1"; }

            DataTable dtPL = new DataTable();
            dtPL.Columns.AddRange(new DataColumn[9] {  new DataColumn("Item", typeof(string)),
                                                            new DataColumn("Description", typeof(string)),
                                                            new DataColumn("MatlDescription", typeof(string)),
                                                            new DataColumn("Qty", typeof(int)),
                                                            new DataColumn("Priority",typeof(int)),
                                                            new DataColumn("FirstGrp",typeof(string)),
                                                            new DataColumn("NewGrp",typeof(string)),
                                                            new DataColumn("LineLoc",typeof(string)),
                                                            new DataColumn("StartDate",typeof(string))});

            string[] selectedColumns = new[] { "Item", "Description", "MatlDescription", "ItemQty", "LineLoc", "Priority", "FirstGrp", "StartDate" };
            DataTable dt = new DataView(dsPicklist.Tables[0]).ToTable(false, selectedColumns);

            // The "LineLoc" session variable will become the final part of the Excel file name when selected to print/output by the user.
            //Session["LineLoc"] = strLineLoc;

            foreach (DataRow dr in dt.Rows)
            {
                string strItem = dr[0].ToString();
                string strDescription = dr[1].ToString();
                string strMatlDescription = dr[2].ToString();
                int intQty = Convert.ToInt32(dr[3].ToString());
                string strDate = dr[7].ToString();
                string strLineLoc = dr[4].ToString();
                string _LineLoc = "";
                string strCurrentRGID = dr[6].ToString();
                ViewBag.strLineLoc = dr[4].ToString();
                strNewRGID = strRGID;

                switch (strCnt)
                {
                    case "1":
                        _LineLoc = "Finish";
                        break;
                    case "2":
                        _LineLoc = "Frames";
                        break;
                    case "3":
                        _LineLoc = "Fronts";
                        break;
                    case "4":
                        _LineLoc = "Sides";
                        break;
                    case "5":
                        _LineLoc = "Tailgate";
                        break;
                    case "6":
                        _LineLoc = "Weldout";
                        break;

                }

                ViewBag.strLineLoc = _LineLoc;

                // Compare dates in order to reassign resource group.
                bool changeTo2006 = (strCurrentRGID == "200-2" && strDate == (string)Session["PickDatePlus1"] && strLineLoc == "Frames");
                if (changeTo2006) { strNewRGID = "200-6"; }

                bool changeTo2002 = (strCurrentRGID == "200-6" && strDate == (string)Session["PickDateMinus1"] && strLineLoc != "Frames");
                if (changeTo2002) { strNewRGID = "200-2"; }

                if (_LineLoc == strLineLoc && strPriority == dr[5].ToString() && (strNewRGID == dr[6].ToString() || changeTo2002 || changeTo2006))
                {
                    dtPL.Rows.Add(strItem, strDescription, strMatlDescription, intQty, strPriority, strCurrentRGID, strNewRGID, strLineLoc, strDate);
                }
                else if (_LineLoc == strLineLoc && strRGID != strNewRGID)
                {
                    dtPL.Rows.Add(strItem, strDescription, strMatlDescription, intQty, strPriority, strRGID, strNewRGID, strLineLoc, strDate);
                }
            }

            dtPL.DefaultView.Sort = "FirstGrp, Priority, StartDate, Item";
            DataTable dtX = dtPL.DefaultView.ToTable(true);
            dtPL = dtX;

            dtExcelExport = dtPL;
            Session["dtPL"] = dtPL;

            switch (strCnt)
            {
                case "1":
                    Session["LineLocDataSource1"] = dtPL;
                    Session["Grid1RowCount"] = dtPL.Rows.Count.ToString();
                    Session["dataSourcegrdData1"] = ConvertDataTableToJSON(dtPL);
                    if (strNewRGID == "200-6")
                    {
                        Session["LineLocDataSource1"] = null;
                        Session["dataSourcegrdData1"] = null;
                        Session["Grid1RowCount"] = "0";
                    }
                    break;
                case "2":
                    Session["LineLocDataSource2"] = dtPL;
                    Session["Grid2RowCount"] = dtPL.Rows.Count.ToString();
                    Session["dataSourcegrdData2"] = ConvertDataTableToJSON(dtPL);
                    if (strNewRGID != "200-6")
                    {
                        Session["LineLocDataSource2"] = null;
                        Session["dataSourcegrdData2"] = null;
                        Session["Grid2RowCount"] = "0";
                    }
                    break;
                case "3":
                    Session["LineLocDataSource3"] = dtPL;
                    Session["Grid3RowCount"] = dtPL.Rows.Count.ToString();
                    Session["dataSourcegrdData3"] = ConvertDataTableToJSON(dtPL);
                    if (strNewRGID == "200-6")
                    {
                        Session["LineLocDataSource3"] = null;
                        Session["dataSourcegrdData3"] = null;
                        Session["Grid3RowCount"] = "0";
                    }
                    break;
                case "4":
                    Session["LineLocDataSource4"] = dtPL;
                    Session["Grid4RowCount"] = dtPL.Rows.Count.ToString();
                    Session["dataSourcegrdData4"] = ConvertDataTableToJSON(dtPL);
                    if (strNewRGID == "200-6")
                    {
                        Session["LineLocDataSource4"] = null;
                        Session["dataSourcegrdData4"] = null;
                        Session["Grid4RowCount"] = "0";
                    }
                    break;
                case "5":
                    Session["LineLocDataSource5"] = dtPL;
                    Session["Grid5RowCount"] = dtPL.Rows.Count.ToString();
                    Session["dataSourcegrdData5"] = ConvertDataTableToJSON(dtPL);
                    if (strNewRGID == "200-6")
                    {
                        Session["LineLocDataSource5"] = null;
                        Session["dataSourcegrdData5"] = null;
                        Session["Grid5RowCount"] = "0";
                    }
                    break;
                case "6":
                    Session["LineLocDataSource6"] = dtPL;
                    Session["Grid6RowCount"] = dtPL.Rows.Count.ToString();
                    Session["dataSourcegrdData6"] = ConvertDataTableToJSON(dtPL);
                    if (strNewRGID == "200-6")
                    {
                        Session["LineLocDataSource6"] = null;
                        Session["dataSourcegrdData6"] = null;
                        Session["Grid6RowCount"] = "0";
                    }
                    break;
            }

        }

      public void ExportToExcel(string GridModel)
      {
         // ClosedXML method of exporting to Excel

         // Create an instance of a workbook.
         var wb = new XLWorkbook();

         // Get datatables used to fill each grid.
         var dtFinish = (DataTable)Session["LineLocDataSource1"];
         var dtFrames = (DataTable)Session["LineLocDataSource2"];
         var dtFronts = (DataTable)Session["LineLocDataSource3"];
         var dtTailgate = (DataTable)Session["LineLocDataSource4"];
         var dtSides = (DataTable)Session["LineLocDataSource5"];
         var dtWeldout = (DataTable)Session["LineLocDataSource6"];

        // Add datatable data to a worksheet tab and name accordingly.
        if (dtFinish !=null) { wb.Worksheets.Add(dtFinish, "Finish"); }
        if (dtFrames != null) { wb.Worksheets.Add(dtFrames, "Frames"); }
        if (dtFronts != null) { wb.Worksheets.Add(dtFronts, "Fronts"); }
        if (dtTailgate != null) { wb.Worksheets.Add(dtTailgate, "Tailgate"); }
        if (dtSides != null) { wb.Worksheets.Add(dtSides, "Sides"); }
        if (dtWeldout != null) { wb.Worksheets.Add(dtWeldout, "Weldout"); }

         // Prepare the file to be downloaded.
         Response.Clear();
         Response.Buffer = true;
         Response.AddHeader("content-disposition", "attachment; filename=Picklist.xlsx");
         Response.ContentType = "application/vnd.ms-excel";

         // Flush the workbook to the Response.OutputStream
         using (MemoryStream memoryStream = new MemoryStream())
         {
            wb.SaveAs(memoryStream);
            memoryStream.WriteTo(Response.OutputStream);
            memoryStream.Close();
         }

         // File has been downloaded.
         Response.End();
      }

   }
}

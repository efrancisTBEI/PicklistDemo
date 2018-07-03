#region Copyright Syncfusion Inc. 2001-2018.
// Copyright Syncfusion Inc. 2001-2018. All rights reserved.
// Use of this code is subject to the terms of our license.
// A copy of the current license can be obtained at any time by e-mailing
// licensing@syncfusion.com. Any infringement will be prosecuted under
// applicable laws. 
#endregion
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

namespace PicklistDemo.Controllers
{
    public class HomeController : Controller
    {

        DataSet dsPicklist = new DataSet();

        public ActionResult Index(string strPickDate)
        {

            if (strPickDate == null)
            {
                strPickDate = "06/14/2018";
                ViewBag.strRGID = "";
            }


            //ViewBag.strPickDate = strPickDate;

            ViewBag.PicklistDates = GetPicklistDates();

            var cnSTR = ConfigurationManager.ConnectionStrings["LC_AppConnectionString"].ConnectionString;
            SqlConnection cn = new SqlConnection(cnSTR);
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = cn;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "TBEI_PicklistGetByDate";
            if (strPickDate == "") { strPickDate = "06/14/2018"; };
            cmd.Parameters.Add("@startdate", SqlDbType.DateTime).Value = Convert.ToDateTime(strPickDate);

            SqlDataAdapter da = new SqlDataAdapter(cmd);
            DataSet ds = new DataSet();
            da.Fill(ds);

            DataTable dtPLGrps = new DataTable();
            dtPLGrps.Columns.AddRange(new DataColumn[2] {   new DataColumn("RGID", typeof(string)),
                                                            new DataColumn("Priority", typeof(string))});
            cmd.Connection = cn;
            cmd.Parameters.Clear();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "TBEI_PicklistGetByResourceGroup";
            cmd.Parameters.Add("@startdate", SqlDbType.DateTime).Value = Convert.ToDateTime(strPickDate);
            cmd.Parameters.Add("@startresgrp", SqlDbType.NVarChar).Value = ViewBag.strRGID;

            da.SelectCommand = cmd;
            da.Fill(dsPicklist);

            Session["dsPicklist"] = dsPicklist;

            for (int x = 1; x <= ds.Tables[0].Rows.Count - 1; x++)
            {
                string strRGID = ds.Tables[0].Rows[x]["FirstGrp"].ToString();
                string strPriority = ds.Tables[0].Rows[x]["Priority"].ToString();

                string[] selectedColumns = new[] { "Item", "Description", "MatlDescription", "ItemQty", "LineLoc", "Priority", "FirstGrp", "StartDate" };
                DataTable dt = new DataView(dsPicklist.Tables[0]).ToTable(false, selectedColumns);
                DataRow[] drRows = dt.Select("StartDate = '" + strPickDate + "' AND Priority = '" + strPriority + "' AND FirstGrp = '" + strRGID + "'");

                if (drRows.Count() > 0)
                {
                    dtPLGrps.Rows.Add(strRGID, strPriority);
                }
            }

            ViewBag.dtPLGrps = ConvertDataTableToJSON(dtPLGrps);
            TempData["strPickDates"] = strPickDate;
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

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
                while (dr.Read())
                {
                    dropObj.Add(new
                    {
                        Text = dr.GetValue(0).ToString(),
                        Value = dr.GetValue(0).ToString()
                    });
                }
            }

            cn.Close();

            return dropObj;
        }

    }
}
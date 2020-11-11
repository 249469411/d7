using Bll;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.UI.WebControls;

namespace Hopu
{
    public class ToExcel
    {
        public void LogToExcel()
        {
            GridView gvNew = new GridView();

            gvNew.DataSource = GetData();
            gvNew.DataBind();

            System.Web.UI.Control ctl = gvNew;
            HttpContext.Current.Response.AppendHeader("Content-Disposition", "attachment;filename=HopuLog.xls");
            HttpContext.Current.Response.Charset = "UTF-8";
            HttpContext.Current.Response.ContentEncoding = System.Text.Encoding.Default;
            HttpContext.Current.Response.ContentType = "application/ms-excel";
            //ctl.Page.EnableViewState = false;
            System.IO.StringWriter tw = new System.IO.StringWriter();
            System.Web.UI.HtmlTextWriter hw = new System.Web.UI.HtmlTextWriter(tw);
            ctl.RenderControl(hw);
            HttpContext.Current.Response.Write(tw.ToString());
            HttpContext.Current.Response.End();
        }

        private DataTable GetData()
        {
            DataTable dt = CreateStructure();
            DataTable gvMsg = new LogsBll().GetLogLists("");

            for (int i = 0; i < gvMsg.Rows.Count; i++)
            {


                DataRow dr = dt.NewRow();
                dr["编号"] = Convert.ToInt32(gvMsg.Rows[i]["logid"]);
                dr["操作者"] = gvMsg.Rows[i]["adminname"].ToString();

                dt.Rows.Add(dr);

            }

            return dt;
        }

        private DataTable CreateStructure()
        {

            DataTable dt = new DataTable();
            dt.Columns.Add(new DataColumn("编号", typeof(int)));
            dt.Columns.Add(new DataColumn("操作者", typeof(string)));

            return dt;
        }
    }
}
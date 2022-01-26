using System;
using XMasReader.Services;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.HtmlControls;
using System.Collections;

namespace XMasReader
{
    public partial class Xmas : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            // I started working on having it display on the webpage after hitting the button, but abandonded.
            // This portion is irrelevant for usage, only was going to be used for testing.
        
            //ExcelReader eReader = new ExcelReader();
            //eReader.UploadExcel();


            //ArrayList eList = eReader.UploadExcel();

            //HtmlTable table = new HtmlTable();

            //table.Border = 1;
            //table.CellPadding = 3;
            //table.CellSpacing = 3;
            //table.BorderColor = "red";

            //HtmlTableRow row;
            //HtmlTableCell cell;
            //int i = 1;
            //foreach (ArrayList list in eList)
            //{
            //    row = new HtmlTableRow();
            //    row.BgColor = (i % 2 == 0 ? "lightyellow" : "lightcyan");
            //    i++;

            //    for (int j = 0; i < 10; j++)
            //    {
            //        cell = new HtmlTableCell();
            //        cell.InnerHtml = list.FirstName;
            //    }
            //}
        }

        protected void Btn_Click(Object sender, EventArgs e)
        {
            ExcelReader eReader = new ExcelReader();
            eReader.UploadExcel();
        }
    }
}

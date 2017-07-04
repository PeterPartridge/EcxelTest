using Microsoft.Office.Interop;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Excel = Microsoft.Office.Interop.Excel;


namespace EcxelTest
{
    public partial class WebForm1 : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

            var itemsList = play();
        }
        public List<MyItems> play()
        {
            List<MyItems> items = new List<MyItems>();
            Excel.Application exclApp = new Microsoft.Office.Interop.Excel.Application();
            Excel.Workbook wkbook = exclApp.Workbooks.Open("D:/Coding/excelTest.xlsx");
            Excel._Worksheet wkSheet = wkbook.Worksheets[1];
            Excel.Range usedRange = wkSheet.UsedRange;
            int rowCounter = usedRange.Rows.Count;
            int colCounter = usedRange.Columns.Count;
            int rowCount = 2;
            int colCount = 1;
            while (rowCount <= rowCounter)
            {
                items.Add(new MyItems { Thingy = usedRange.Cells[rowCount, colCount].Value.ToString(), Thingy2 = (usedRange.Cells[rowCount, colCount += 1].Value == null) ? "Empty" : usedRange.Cells[rowCount, colCount].Value.ToString() });
                rowCount++;
                colCount = 1;
            }
            wkbook.Close();
            exclApp.Quit();
            return items;
        }


        public class MyItems
        {


            public string Thingy { get; set; }
            public string Thingy2 { get; set; }



        }


    }
}


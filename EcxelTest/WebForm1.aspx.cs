using Microsoft.Office.Interop;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
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
            //list of the object myItems 
            List<MyItems> items = new List<MyItems>();
            // open the excel application
            Excel.Application exclApp = new Microsoft.Office.Interop.Excel.Application();
            //open a workbook in a defined file path.
            Excel.Workbook wkbook = exclApp.Workbooks.Open("D:/Coding/excelTest.xlsx");
            // this will open the first worksheet
            Excel._Worksheet wkSheet = wkbook.Worksheets[1];
            //get the range of used cells
            Excel.Range usedRange = wkSheet.UsedRange;
            //get max number of rows
            int rowCounter = usedRange.Rows.Count;
            //get max number of columns 
            int colCounter = usedRange.Columns.Count;
            //this will start on the second row of the excel sheet and expect their to be a heading
            int rowCount = 2;
            //start on the first column
            int colCount = 1;
            //loopiung through rows.
            while (rowCount <= rowCounter)
            {
                items.Add(new MyItems { Thingy = usedRange.Cells[rowCount, colCount].Value.ToString(), Thingy2 = (usedRange.Cells[rowCount, colCount += 1].Value == null) ? "Empty" : usedRange.Cells[rowCount, colCount].Value.ToString() });
                
                rowCount++;
                colCount = 1;
            }
            //close the workbook
            wkbook.Close();
            //exit the excel applicaiton
            
            exclApp.Quit();
            Marshal.ReleaseComObject(wkSheet);
            Marshal.ReleaseComObject(wkbook);
            Marshal.ReleaseComObject(exclApp);
            exclApp = null;
            return items;
        }

        public class MyItems
        {
            public string Thingy { get; set; }
            public string Thingy2 { get; set; }
        }


    }
}


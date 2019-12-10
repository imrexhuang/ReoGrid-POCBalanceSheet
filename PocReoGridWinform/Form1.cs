using System;
using System.Windows.Forms;
using unvell.ReoGrid;
using unvell.ReoGrid.DataFormat;

namespace PocReoGridWinform
{
    public partial class Form1 : Form
    {
        /*
         官方文件:
        https://reogrid.net/document/

        DEMO
        https://github.com/unvell/ReoGrid/tree/master/Demo

        ReoScript
        https://reogrid.net/document/script-execution/
             */

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //https://www.codeproject.com/Articles/691749/Free-NET-Spreadsheet-Control
            // create instance
            var grid = new ReoGridControl()
            {
                // put to fill the parent control
                Dock = DockStyle.Fill,

            };

            var worksheet = grid.CurrentWorksheet;

            worksheet["A1"] = "會計科目編號";
            worksheet["B1"] = "會計科目名稱";
            worksheet["C1"] = "子目編號";
            worksheet["D1"] = "子目名稱";
            worksheet["E1"] = "金額";
            worksheet.AutoFitColumnWidth(0, false);
            worksheet.AutoFitColumnWidth(1, false);
            worksheet.AutoFitColumnWidth(2, false);
            worksheet.AutoFitColumnWidth(3, false);
            worksheet["A2"] = new object[,] {
              { "2123", "應付代收款", null, null ,"=SUM(E3:E5)"},
              { "2123", "應付代收款",  "A00006", "公庫利息" , 109827},
              { "2123", "應付代收款",  "C00016", "畢業旅行經費" , 250000},
              { "2123", "應付代收款",  "C00017", "家長會費" , 5889000},
            };
            worksheet.HideRows(2, 3);

            // https://reogrid.net/document/row-and-column/
            //worksheet.HideColumns(7, 3); //H、I、J欄被隱藏了

            //worksheet.HideRows(8, 5); //9到13列被隱藏了

            this.Controls.Add(grid);

            
        }
    }
}

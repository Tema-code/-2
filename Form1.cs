using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using E = Microsoft.Office.Interop.Excel;


namespace Работа2
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            
            E.Application oXls = new E.Application();
            oXls.Visible = true;
            E.Workbook oWb = oXls.Workbooks.Add();
            oXls.SheetsInNewWorkbook = 1;
            E.Worksheet sheet = (E.Worksheet)oXls.Worksheets.get_Item(1);

            
            E.Range oRng = oXls.Range["A1", "L1"];
            oRng.VerticalAlignment = E.XlVAlign.xlVAlignCenter;
            oRng.HorizontalAlignment = E.XlHAlign.xlHAlignCenter;
            oRng.Cells.Font.Size = 10;
            oRng.Cells.Font.Bold = 1;
            oRng.Rows.RowHeight = 50;
            oRng.Columns.ColumnWidth = 15;
            oRng.Borders.Weight = 2;
            oRng.Cells.WrapText = true;
            sheet.Cells[1] = "Линии развития";
            E.Range rng2 = oXls.Range["B1", "E1"];
            rng2.Cells.Merge();
            rng2.Columns.ColumnWidth = 7;
            rng2.Cells.Font.Bold = 0;
            rng2.VerticalAlignment = E.XlVAlign.xlVAlignTop;
            rng2.HorizontalAlignment = E.XlHAlign.xlHAlignLeft;
            sheet.Cells[1, 2] = "1. Производить вычисления для принятия решений в различных жизненных ситуация";
            E.Range rng3 = oXls.Range["F1", "L1"];
            rng3.Cells.Merge();
            rng3.Cells.Font.Bold = 0;
            rng3.Columns.ColumnWidth = 7;
            rng3.VerticalAlignment = E.XlVAlign.xlVAlignTop;
            rng3.HorizontalAlignment = E.XlHAlign.xlHAlignLeft;
            sheet.Cells[1, 6] = "2. Читать и записывать сведения об окружающем мире на языке математики";

            
            E.Range oRng2 = oXls.Range["A2", "L2"];
            oRng2.VerticalAlignment = E.XlVAlign.xlVAlignBottom;
            oRng2.HorizontalAlignment = E.XlHAlign.xlHAlignLeft;
            oRng2.Orientation = 90;
            oRng2.Cells.Font.Size = 10;
            oRng2.Cells.Font.Bold = 0;
            oRng2.Rows.RowHeight = 190;
            oRng2.Borders.Weight = 2;
            oRng2.Cells.WrapText = true;
            sheet.Cells[2, 2] = "• читать, записывать и сравнивать числа в пределах 1 000 000";
            sheet.Cells[2, 3] = "• складывать, вычитать, умножать и делить числа в пределах 1 000 000";
            sheet.Cells[2, 4] = "• находить значения выражений в 2-4 действия";
            sheet.Cells[2, 5] = "• сравнивать именованные числа и выполнять 4 арифметических действия с ними";
            sheet.Cells[2, 6] = "• читать и записывать именованные числа (длина, площадь, масса, объём)";
            sheet.Cells[2, 7] = "• читать информацию, заданную с помощью столбчатых, линейных и круговых диаграмм, таблиц, графов";
            sheet.Cells[2, 8] = "• переносить информацию из таблицы в линейные и столбчатые диаграммы";
            sheet.Cells[2, 9] = "• находить значение выражений с переменной (изученных видов)";
            sheet.Cells[2, 10] = "• находить среднее арифметическое двух чисел";
            sheet.Cells[2, 11] = "• определять время по часам (до минуты)";
            sheet.Cells[2, 12] = "• сравнивать и упорядочивать объекты по разным признакам (длина, масса, объём)";

            E.Range rng4 = oXls.Range["A2"];
            rng4.Borders[E.XlBordersIndex.xlDiagonalDown].Weight = 2;
            rng4.Orientation = 0;
            rng4.Cells.Font.Bold = 1;
        }
    }
}

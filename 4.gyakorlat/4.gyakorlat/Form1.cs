using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Drawing.Text;

namespace _4.gyakorlat
{

    public partial class Form1 : Form
    {
        
        RealEstateEntities context = new RealEstateEntities();
        List<Flat> Flats;
        Excel.Application xlApp; 
        Excel.Workbook xlWB; 
        Excel.Worksheet xlSheet; 
        public Form1()
        {
            InitializeComponent();
            LoadData();
            CreateExcel();
            
        }


        private void LoadData()
        {
           Flats = context.Flat.ToList();
        }

        private void CreateExcel()
        {
            try
            {
                
                xlApp = new Excel.Application();

               
                xlWB = xlApp.Workbooks.Add(Missing.Value);

                
                xlSheet = xlWB.ActiveSheet;

                
                CreateTable(); 

                
                xlApp.Visible = true;
                xlApp.UserControl = true;
            }
            catch (Exception ex) 
            {
                string errMsg = string.Format("Error: {0}\nLine: {1}", ex.Message, ex.Source);
                MessageBox.Show(errMsg, "Error");

                
                xlWB.Close(false, Type.Missing, Type.Missing);
                xlApp.Quit();
                xlWB = null;
                xlApp = null;
            }

            
        }

        private void CreateTable()
        {
            string[] headers = new string[]
            {
                "Kód",
                "Eladó",
                "Oldal",
                "Kerület",
                "Lift",
                "Szobák száma",
                "Alapterület (m2)",
                "Ár (mFt)",
                "Egy szobára jutó ár"
            };

            for (int x = 0; x < headers.Length; x++)
            {
                xlSheet.Cells[1, 1] = headers[0];
                xlSheet.Cells[1, 2] = headers[1];
                xlSheet.Cells[1, 3] = headers[2];
                xlSheet.Cells[1, 4] = headers[3];
                xlSheet.Cells[1, 5] = headers[4];
                xlSheet.Cells[1, 6] = headers[5];
                xlSheet.Cells[1, 7] = headers[6];
                xlSheet.Cells[1, 8] = headers[7];
                xlSheet.Cells[1, 9] = headers[8];
            }

            object[,] values = new object[Flats.Count, headers.Length];

            int counter = 0;

            foreach (var x in Flats)
            {
                values[counter, 0] = x.Code;
                values[counter, 1] = x.Vendor;
                values[counter, 2] = x.Side;
                values[counter, 3] = x.District;
                values[counter, 4] = x.Elevator;
                values[counter, 5] = x.NumberOfRooms;
                values[counter, 6] = x.FloorArea;
                values[counter, 7] = x.Price;
                values[counter, 8] ="="+GetCell(counter+2,8)+"/"+GetCell(counter+2,6);
                counter++;

            }

            xlSheet.get_Range(
             GetCell(2, 1),
             GetCell(1 + values.GetLength(0), values.GetLength(1))).Value2 = values;

            int lastRowID = xlSheet.UsedRange.Rows.Count;

            Excel.Range headerRange = xlSheet.get_Range(GetCell(1, 1), GetCell(1, headers.Length));
            Excel.Range sheet= xlSheet.get_Range(GetCell(1, 1), GetCell(lastRowID, headers.Length));
            Excel.Range utolsooszlop = xlSheet.get_Range(GetCell(2, 9), GetCell(lastRowID, headers.Length));
            Excel.Range elsooszlop = xlSheet.get_Range(GetCell(2, 1), GetCell(lastRowID, 1));

            sheet.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThick);

            utolsooszlop.Interior.Color = Color.LightGreen;
            utolsooszlop.NumberFormat= "##0.00";

            elsooszlop.Interior.Color = Color.LightYellow;
            elsooszlop.Font.Bold = true;

            headerRange.Font.Bold = true;
            headerRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            headerRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            headerRange.EntireColumn.AutoFit();
            headerRange.RowHeight = 40;
            headerRange.Interior.Color = Color.LightBlue;
            headerRange.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThick);

        }

        private string GetCell(int x, int y)
        {
            string ExcelCoordinate = "";
            int dividend = y;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                ExcelCoordinate = Convert.ToChar(65 + modulo).ToString() + ExcelCoordinate;
                dividend = (int)((dividend - modulo) / 26);
            }
            ExcelCoordinate += x.ToString();

            return ExcelCoordinate;
        }

    }
}

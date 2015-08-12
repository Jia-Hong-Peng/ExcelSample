using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using NPOI.HPSF;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace ExcelSample
{
    public partial class Form1 : Form
    {

        //static IWorkbook hssfworkbook;

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

           // LoadTemplate();


            //FileStream file = new FileStream(@"template/Astro2_Sys WHCK Status - 0712-SA2.xlsx", FileMode.Open, FileAccess.ReadWrite);
            FileStream file = new FileStream(@"template/Astro2_Sys WHCK Status - 0712-SA2.xlsx", FileMode.Open, FileAccess.ReadWrite);
            IWorkbook wb = new XSSFWorkbook(file);


            MessageBox.Show(wb.NumberOfSheets+"", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            if (wb.NumberOfSheets >= 2)
            {
                ISheet sheet1 = wb.GetSheetAt(1);

                sheet1.GetRow(7).GetCell(6).SetCellValue("Failed");

                //Write the stream data of workbook to the root directory
                //FileStream sw = new FileStream(@"test.xls", FileMode.Create);                
                FileStream sw = new FileStream(@"test.xls", FileMode.OpenOrCreate);
                wb.Write(file);
                sw.Close();

            }
            //create cell on rows, since rows do already exist,it's not necessary to create rows again.
         //   sheet1.GetRow(8).GetCell(7).SetCellValue("Failed");
            /*
            sheet1.GetRow(2).GetCell(1).SetCellValue(300);
            sheet1.GetRow(3).GetCell(1).SetCellValue(500050);
            sheet1.GetRow(4).GetCell(1).SetCellValue(8000);
            sheet1.GetRow(5).GetCell(1).SetCellValue(110);
            sheet1.GetRow(6).GetCell(1).SetCellValue(100);
            sheet1.GetRow(7).GetCell(1).SetCellValue(200);
            sheet1.GetRow(8).GetCell(1).SetCellValue(210);
            sheet1.GetRow(9).GetCell(1).SetCellValue(2300);
            sheet1.GetRow(10).GetCell(1).SetCellValue(240);
            sheet1.GetRow(11).GetCell(1).SetCellValue(180123);
            sheet1.GetRow(12).GetCell(1).SetCellValue(150);
             * */

            //Force excel to recalculate all the formula while open
            //sheet1.ForceFormulaRecalculation = true;

            
            //Response.AddHeader("Content-Disposition", string.Format("attachment; filename=EmptyWorkbook.xls"));
            //Response.BinaryWrite(ms.ToArray());
           // WriteToFile();
          


          //  wb = null;
            //ms.Close();
            //ms.Dispose();
        }


        static void WriteToFile()
        {
            //Write the stream data of workbook to the root directory
            FileStream file = new FileStream(@"test.xls", FileMode.Create);
           // hssfworkbook.Write(file);
            file.Close();
        }


        static void LoadTemplate()
        {
            //read the template via FileStream, it is suggested to use FileAccess.Read to prevent file lock.
            //book1.xls is an Excel-2007-generated file, so some new unknown BIFF records are added. 
            FileStream file = new FileStream(@"template/1.xlsx", FileMode.Open, FileAccess.Read);

         //   hssfworkbook = new XSSFWorkbook(file);

            //create a entry of DocumentSummaryInformation
            DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
            dsi.Company = "Compal";
            //hssfworkbook.DocumentSummaryInformation = dsi;
            

            //create a entry of SummaryInformation
            SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
            si.Subject = "Compal Excel Example";
           // hssfworkbook.SummaryInformation = si;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string src = @"template/2.xlsx";

            //OpenXML SDK 2.5
            //REF: http://msdn.microsoft.com/en-us/library/office/cc850837.aspx
            string dst = src.Replace(Path.GetFileName(src), "Astro2_Sys WHCK Status - 0712-SA2.xlsx"); //另存目的檔
            File.Copy(src, dst, true);
            using (var shtDoc = SpreadsheetDocument.Open(dst, true))
            {
                //var sht = shtDoc.WorkbookPart.Workbook.Descendants<Sheet>().First();
                var sht = shtDoc.WorkbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == "Astro 2 TestResult 9431").FirstOrDefault();
                var shtPart = shtDoc.WorkbookPart.GetPartById(sht.Id) as WorksheetPart;
                //var cell = shtPart.Worksheet.Descendants<Row>().First().Descendants<Cell>().First();

 // 3. 建立 Cell 物件，設定寫入位置，格式，資料

                Cell cell = InsertCellInWorksheet("G", 8, shtPart);
              

                //REF: InlineString http://bit.ly/ZpUf18
                var ins = new InlineString();
                ins.AppendChild(new Text("Failed"));
                cell.AppendChild(ins);
                cell.DataType = new DocumentFormat.OpenXml.EnumValue<CellValues>(CellValues.InlineString);
                shtPart.Worksheet.Save();
                shtDoc.WorkbookPart.Workbook.Save();
                shtDoc.Close();
            }
        }



        /// <summary>
        /// 創建一個SpreadsheetDocument對像
        /// </summary>
        /// <param name="excelFileName"></param>
        /// <returns></returns>
        static SpreadsheetDocument CreateSpreadsheetDocument(string excelFileName)
        {
            SpreadsheetDocument excel = SpreadsheetDocument.Create(excelFileName, SpreadsheetDocumentType.Workbook, true);
            WorkbookPart workbookpart = excel.AddWorkbookPart();
            workbookpart.Workbook = new Workbook();
            return excel;
        }
 
        /// <summary>
        /// 插入worksheet
        /// </summary>
        /// <param name="workbookPart"></param>
        /// <returns></returns>
        private static WorksheetPart InsertWorksheet(WorkbookPart workbookPart, string sheetName = null)
        {
            //創建一個新的WorkssheetPart（後面將用它來容納具體的Sheet）
            WorksheetPart newWorksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            newWorksheetPart.Worksheet = new Worksheet(new SheetData());
            newWorksheetPart.Worksheet.Save();
 
            //取得Sheet集合
            Sheets sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
            if (sheets == null)
            {
                sheets = workbookPart.Workbook.AppendChild<Sheets>(new Sheets());
            }
 
            string relationshipId = workbookPart.GetIdOfPart(newWorksheetPart);
 
            //得到Sheet的唯一序號
            uint sheetId = 1;
            if (sheets.Elements<Sheet>().Count() > 0)
            {
                sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
            }
 
            string sheetTempName = "Sheet" + sheetId;
 
            if (sheetName != null)
            {
                bool hasSameName = false;
                //檢測是否有重名
                foreach (var item in sheets.Elements<Sheet>())
                {
                    if (item.Name == sheetName)
                    {
                        hasSameName = true;
                        break;
                    }
                }
                if (!hasSameName)
                {
                    sheetTempName = sheetName;
                }
            }
 
            //創建Sheet實例並將它與sheets關聯
            Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetTempName };
            sheets.Append(sheet);
            workbookPart.Workbook.Save();
 
            return newWorksheetPart;
        }
 
        /// <summary>
        /// 創建一個SharedStringTablePart(相當於各Sheet共用的存放字符串的容器)
        /// </summary>
        /// <param name="workbookPart"></param>
        /// <returns></returns>
        private static SharedStringTablePart CreateSharedStringTablePart(WorkbookPart workbookPart)
        {
            SharedStringTablePart shareStringPart = null;
            if (workbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
            {
                shareStringPart = workbookPart.GetPartsOfType<SharedStringTablePart>().First();
            }
            else
            {
                shareStringPart = workbookPart.AddNewPart<SharedStringTablePart>();
            }
            return shareStringPart;
        }
 
        /// <summary>
        /// 向工作表插入一個單元格
        /// </summary>
        /// <param name="columnName">列名稱</param>
        /// <param name="rowIndex">行索引</param>
        /// <param name="worksheetPart"></param>
        /// <returns></returns>
        private static Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
        {
            Worksheet worksheet = worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            string cellReference = columnName + rowIndex;//列的引用字符串，類似:"A3"或"B5"
 
            //如果指定的行存在，則直接返回該行，否則插入新行
            Row row;
            if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
            {
                row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
            }
            else
            {
                row = new Row() { RowIndex = rowIndex };
                sheetData.Append(row);
            }
 
            //如果該行沒有指定ColumnName的列，則插入新列，否則直接返回該列
            if (row.Elements<Cell>().Where(c => c.CellReference.Value == columnName + rowIndex).Count() > 0)
            {
                return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
            }
            else
            {
                //列必須按(字母)順序插入，因此要先根據"列引用字符串"查找插入的位置
                Cell refCell = null;
                foreach (Cell cell in row.Elements<Cell>())
                {
                    if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                    {
                        refCell = cell;
                        break;
                    }
                }
 
                Cell newCell = new Cell() { CellReference = cellReference };
                row.InsertBefore(newCell, refCell);
 
                worksheet.Save();
                return newCell;
            }
        }
 
        /// 向SharedStringTablePart添加字符串
        /// </summary>
        /// <param name="text">字符串內容</param>
        /// <param name="shareStringPart">sharedStringTablePart內容</param>
        /// <returns>如果要添加的字符串已經存在，則直接返回該字符串的索引</returns>
        private static int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart)
        {
            //檢測SharedStringTable是否存在，如果不存在，則創建一個
            if (shareStringPart.SharedStringTable == null)
            {
                shareStringPart.SharedStringTable = new SharedStringTable();
            }
 
            int i = 0;
 
            //遍歷SharedStringTable中所有的Elements，查看目標字符串是否存在
            foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
            {
                if (item.InnerText == text)
                {
                    return i;
                }
                i++;
            }
 
            //如果目標字符串不存在，則創建一個，同時把SharedStringTable的最後一個Elements的"索引+1"返回
            shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));
            shareStringPart.SharedStringTable.Save();
 
            return i;
        }
    


    }
}

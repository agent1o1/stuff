using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Linq;
using System.Text.RegularExpressions;
using System.Data;

namespace OPENXML
{
    public class ExcelReport
    {
        private MemoryStream ms = new MemoryStream();
        private List<DataTable> dataList = new List<DataTable>();
        private int rownumber = 0;
        private int colnumber = 0;

        public ExcelReport(DataInterpreterReport input)
        {
            dataList = input.DataList;
            GenerateDocument();
        }

        public MemoryStream GetReport()
        {
            ms.Position = 0;
            return ms;
        }

        private void GenerateDocument()
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(ms, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart wbp = document.AddWorkbookPart();
                Workbook workbook = new Workbook();
                Worksheet ws = new Worksheet();
                WorksheetPart wsp = wbp.AddNewPart<WorksheetPart>("rId1");
                Sheets sheets = new Sheets();
                Sheet sheet = new Sheet() { Name = "Отчёт", Id = "rId1", SheetId = 1 };
                SheetData sheetdata = new SheetData();

                foreach (DataTable dt in dataList)
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        rownumber++;
                        Row NewRow = new Row();
                        foreach (object obj in row.ItemArray)
                        {
                            NewRow.AppendChild<Cell>(CreateCell(obj));
                        }
                        sheetdata.AppendChild<Row>(NewRow);
                    }
                }


                ws.AppendChild<SheetData>(sheetdata);
                wsp.Worksheet = ws;
                sheets.Append(sheet);
                workbook.Append(sheets);
                wbp.Workbook = workbook;
            }
        }


        private string GetCellName()
        {
            int dividend = colnumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar('A' + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName + rownumber;
        }

        private Cell CreateCell(object value)
        {
            Cell cell;
            switch (value.GetType().Name)
            {
                case "Int16":
                case "Int32":
                case "Int64":
                    {
                        cell = new Cell() { CellValue = new CellValue(value.ToString()), DataType = CellValues.Number, CellReference = GetCellName() };
                        var a = GetCellName();
                    }
                    break;
                //case "DateTime":
                //    {
                //        cell = new Cell() { CellValue = new CellValue(value.ToString()), DataType = CellValues.Date, CellReference = GetCellName() };
                //        var a = GetCellName();
                //    }
                //    break;
                default:
                    {
                        cell = new Cell() { CellValue = new CellValue(value.ToString()), DataType = CellValues.String, CellReference = GetCellName() };
                        var a = GetCellName();
                    }
                    break;
            }
            return cell;
        }
    }
}
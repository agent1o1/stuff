using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text.RegularExpressions;

namespace OPENXML
{
    public class DataInterpreterReport
    {
        private SpreadsheetDocument document;
        private List<SharedStringItem> sharedStringsList = new List<SharedStringItem>();
        private DataTable dataTable = new DataTable();
        private List<Row> rowsList = new List<Row>();
        private SheetData sheetData = new SheetData();
        private SharedStringTablePart SSTP;
        private WorksheetPart wp;

        //Поля для работы алгоритма
        private PropertyCollection Properties = new PropertyCollection();
        private DataRow NewRow;
        private DataTable CurrentTable = new DataTable();
        private bool IsTable = false;
        private Regex regex = new Regex(@"{*}");
        private List<string> ColumnNames = new List<string>();
        private int colindex;

        private List<DataTable> dataList = new List<DataTable>();
        private string path = AppDomain.CurrentDomain.BaseDirectory + @"Templates\";

        public List<DataTable> DataList { get { return dataList; } }
        
        public DataInterpreterReport(DataSet DataSet)
        {
            Properties = DataSet.ExtendedProperties;
            FillData();
            dataList.Add(CurrentTable);
            foreach (DataTable dt in DataSet.Tables)
            {
                var res = new DataInterpreterReport(dt);
                this.dataList.Add(res.CurrentTable);
            }

        }

        private DataInterpreterReport(DataTable dt)
        {
            Properties = dt.ExtendedProperties;
            this.dataTable = dt;
            FillData();
        }

        private void FillDocument()
        {
            document = SpreadsheetDocument.Open(path + Properties["DocumentName"] + ".xlsx", true);
            //Получение доп свойств
            //Получение коллекции строк
            var sheet = document.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().FirstOrDefault();
            wp = (WorksheetPart)document.WorkbookPart.GetPartById(sheet.Id.Value);
            sheetData = wp.Worksheet.GetFirstChild<SheetData>();
            rowsList = sheetData.Elements<Row>().ToList();

            //Получение коллекции SharedStrings
            SSTP = document.WorkbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
            sharedStringsList = SSTP.SharedStringTable.Elements<SharedStringItem>().ToList();
        }

        private void FillData()
        {
            FillDocument();
            foreach (Row item in rowsList)
            {
                NewRow = CurrentTable.NewRow();
                colindex = 0;
                foreach (Cell cell in item)
                {

                    if (cell.CellValue != null && cell.DataType == CellValues.SharedString)
                    {
                        string sharedstring = sharedStringsList.GetRange(Convert.ToInt32(cell.CellValue.InnerText), 1).FirstOrDefault().InnerText;

                        if (regex.Match(sharedstring).Success)
                        {
                            TemplateDetector(sharedstring);
                        }
                        else
                        {

                            InsertValue(sharedstring);
                        }
                    }
                    else if (cell.CellValue != null && cell.DataType == CellValues.Number)
                    {
                        InsertValue((Convert.ToInt32(cell.CellValue.InnerText)));
                    }

                    else if (cell.CellValue != null && cell.DataType == CellValues.Date)
                    {
                        InsertValue(Convert.ToDateTime(cell.CellValue.InnerText));
                    }
                    //если есть проблемы ищи здесь
                    colindex++;
                }
                if (IsTable)
                {
                    FillTable(dataTable);
                    ColumnNames.Clear();
                    IsTable = false;
                }
                else
                {
                    CurrentTable.Rows.Add(NewRow);
                }

            }
        }

        private void TemplateDetector(string input)
        {
            string[] command = Regex.Replace(input, @"{|}", "").Split(':');
            if (command[0] == "Table")
            {
                IsTable = true;
                ColumnNames.Add(command[1]);
            }

            if (command[0] == "Parameters")
            {
                InsertValue(Properties[command[1]]);
            }

            if (command[0] == "Expression")
            {
                //todo
            }
        }

        private void InsertValue(object value)
        {
            try
            {
                NewRow.SetField(colindex, value);
            }
            catch
            {
                CurrentTable.Columns.Add("");

                NewRow.SetField(colindex, value);
            }
        }
        private void FillTable(DataTable dt)
        {
            colindex = 0;
            if (ColumnNames.Count > 1)
            {
                foreach (var item in ColumnNames)
                {
                    InsertValue(dt.Columns[item].Caption);
                    colindex++;
                }
                CurrentTable.Rows.Add(NewRow);
                NewRow = CurrentTable.NewRow();
                colindex = 0;
            }

            if (ColumnNames.Count == 1)
            {
                InsertValue(dt.Columns[ColumnNames.First()].Caption);
                colindex++;
            }

            foreach (DataRow row in dt.Rows)
            {
                foreach (var item in ColumnNames)
                {
                    InsertValue(row[item]);
                    colindex++;
                }
                CurrentTable.Rows.Add(NewRow);
                NewRow = CurrentTable.NewRow();
                colindex = 0;
            }

        }
    }
}
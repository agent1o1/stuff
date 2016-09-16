using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using System.IO;
using System.Data;

namespace OPENXML
{
    public class WordReport
    {
        private List<DataTable> dataList = new List<DataTable>();
        private MemoryStream ms = new MemoryStream();
        private Body body;
        private Table table;

        public WordReport(DataInterpreterReport input)
        {
            dataList = input.DataList;
            BuildDocument();
        }

        public MemoryStream GetReport()
        {
            ms.Position = 0;
            return ms;
        }

        private void BuildDocument()
        {
            var document = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document);
            MainDocumentPart mainpart = document.AddMainDocumentPart();
            body = new Body();
            Document doc = new Document();
            foreach (DataTable dt in dataList)
            {
                table = GenerateTable(dt);
                foreach (DataRow row in dt.Rows)
                {
                    GenerateRow(row);
                }

                body.Append(table);
            }
            doc.Append(body);
            mainpart.Document = doc;
            document.Close();
        }


        private Table GenerateTable(DataTable dt)
        {
            Table table = new Table();

            TableProperties tableProperties = new TableProperties();
            TableWidth tableWidth = new TableWidth() { Width = "15000", Type = TableWidthUnitValues.Auto };
            TableIndentation tableIndentation1 = new TableIndentation() { Width = -10, Type = TableWidthUnitValues.Dxa };

            TableBorders tableBorders = new TableBorders();
            TopBorder topBorder = new TopBorder() { Val = BorderValues.Single, Color = "FFFFFF", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            LeftBorder leftBorder = new LeftBorder() { Val = BorderValues.Single, Color = "FFFFFF", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder = new BottomBorder() { Val = BorderValues.Single, Color = "FFFFFF", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder = new RightBorder() { Val = BorderValues.Single, Color = "FFFFFF", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableBorders.Append(topBorder, leftBorder, bottomBorder, rightBorder);

            tableProperties.Append(tableWidth, tableBorders);

            TableGrid tableGrid = new TableGrid();
            foreach (DataTable item in dataList)
            {
                tableGrid.Append(new GridColumn());
            }

            table.Append(tableProperties, tableGrid);

            return table;
        }

        private void GenerateRow(DataRow row)
        {
            TableRow tableRow = new TableRow();

            foreach (object obj in row.ItemArray)
            {
                tableRow.Append(GenerateCell(obj));
            }

            table.Append(tableRow);
        }

        private TableCell GenerateCell(object value)
        {
            TableCell tableCell = new TableCell();
            TableCellProperties tableCellProperties = CreateCellProp("15000");
            GridSpan gridSpan = new GridSpan() { Val = 4 };
            tableCellProperties.Append(gridSpan);
            Paragraph paragraph1 = new Paragraph();

            ParagraphProperties paragraphProperties1 = new ParagraphProperties();
            Justification justification1 = new Justification() { Val = JustificationValues.Center };

            paragraphProperties1.Append(justification1);

            Run run1 = new Run();

            RunProperties runProperties = new RunProperties();

            run1.Append(runProperties, new Text(value.ToString()));

            paragraph1.Append(paragraphProperties1, run1);

            tableCell.Append(tableCellProperties, paragraph1);

            return tableCell;

        }

        private TableCellProperties CreateCellProp(string width)
        {
            TableCellProperties tableCellProperties = new TableCellProperties();
            TableCellWidth tableCellWidth = new TableCellWidth() { Width = width, Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders = new TableCellBorders();
            TopBorder topBorder = new TopBorder() { Val = BorderValues.Single, Color = "FFFFFF", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            LeftBorder leftBorder = new LeftBorder() { Val = BorderValues.Single, Color = "FFFFFF", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder = new BottomBorder() { Val = BorderValues.Single, Color = "FFFFFF", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder = new RightBorder() { Val = BorderValues.Single, Color = "FFFFFF", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            //delete some borders
            tableCellBorders.Append(topBorder, leftBorder, bottomBorder, rightBorder);

            TableCellMargin tableCellMargin = new TableCellMargin();
            TopMargin topMargin = new TopMargin() { Width = "75", Type = TableWidthUnitValues.Dxa };
            LeftMargin leftMargin = new LeftMargin() { Width = "75", Type = TableWidthUnitValues.Dxa };
            BottomMargin bottomMargin = new BottomMargin() { Width = "75", Type = TableWidthUnitValues.Dxa };
            RightMargin rightMargin = new RightMargin() { Width = "75", Type = TableWidthUnitValues.Dxa };

            tableCellMargin.Append(topMargin, leftMargin, bottomMargin, rightMargin);

            tableCellProperties.Append(new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center });
            tableCellProperties.Append(tableCellWidth, tableCellBorders, tableCellMargin);
            return tableCellProperties;
        }
    }
}
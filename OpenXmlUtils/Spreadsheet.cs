#region File Information
//
// File: "Spreadsheet.cs"
// Purpose: "Create xlxs spreadsheet files"
// Author: "Geoplex"
// 
#endregion

#region (c) Copyright 2014 Geoplex
//
// THE SOFTWARE IS PROVIDED "AS-IS" AND WITHOUT WARRANTY OF ANY KIND,
// EXPRESS, IMPLIED OR OTHERWISE, INCLUDING WITHOUT LIMITATION, ANY
// WARRANTY OF MERCHANTABILITY OR FITNESS FOR A PARTICULAR PURPOSE.
//
// IN NO EVENT SHALL GEOPLEX BE LIABLE FOR ANY SPECIAL, INCIDENTAL,
// INDIRECT OR CONSEQUENTIAL DAMAGES OF ANY KIND, OR ANY DAMAGES WHATSOEVER
// RESULTING FROM LOSS OF USE, DATA OR PROFITS, WHETHER OR NOT ADVISED OF THE
// POSSIBILITY OF DAMAGE, AND ON ANY THEORY OF LIABILITY, ARISING OUT OF OR IN
// CONNECTION WITH THE USE OR PERFORMANCE OF THIS SOFTWARE.
//
#endregion

using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXmlUtils
{
    public static class Spreadsheet
    {
        /// <summary>
        /// Write xlsx spreadsheet file of a list of T objects
        /// Maximum of 24 columns
        /// </summary>
        /// <typeparam name="T">Type of objects passed in</typeparam>
        /// <param name="fileName">Full path filename for the new spreadsheet</param>
        /// <param name="def">A sheet definition used to create the spreadsheet</param>
        public static void Create<T>(
            string fileName,
            SheetDefinition<T> def)
        {
            // open a template workbook
            using (var myWorkbook = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook))
            {
                // create workbook part
                var workbookPart = myWorkbook.AddWorkbookPart();

                // add stylesheet to workbook part
                var stylesPart = myWorkbook.WorkbookPart.AddNewPart<WorkbookStylesPart>();
                Stylesheet styles = new CustomStylesheet();
                styles.Save(stylesPart);

                // create workbook
                var workbook = new Workbook();

                // add work sheet
                var sheets = new Sheets();
                sheets.AppendChild(CreateSheet(1, def, workbookPart));
                workbook.AppendChild(sheets);

                // add workbook to workbook part
                myWorkbook.WorkbookPart.Workbook = workbook;
                myWorkbook.WorkbookPart.Workbook.Save();
                myWorkbook.Close();
            }
        }

        /// <summary>
        /// Write xlsx spreadsheet file of a list of T objects
        /// Maximum of 24 columns
        /// </summary>
        /// <typeparam name="T">Type of objects passed in</typeparam>
        /// <param name="fileName">Full path filename for the new spreadsheet</param>
        /// <param name="defs">A list of sheet definitions used to create the spreadsheet</param>
        public static void Create<T>(
            string fileName,
            IEnumerable<SheetDefinition<T>> defs)
        {
            // open a template workbook
            using (var myWorkbook = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook))
            {
                // create workbook part
                var workbookPart = myWorkbook.AddWorkbookPart();

                // add stylesheet to workbook part
                var stylesPart = myWorkbook.WorkbookPart.AddNewPart<WorkbookStylesPart>();
                Stylesheet styles = new CustomStylesheet();
                styles.Save(stylesPart);

                // create workbook
                var workbook = new Workbook();

                // add work sheets
                var sheets = new Sheets();
                var list = defs.ToList();
                for (var i = 0; i < list.Count; i++)
                    sheets.AppendChild(CreateSheet(i + 1, list[i], workbookPart));
                workbook.AppendChild(sheets);

                // add workbook to workbook part
                myWorkbook.WorkbookPart.Workbook = workbook;
                myWorkbook.WorkbookPart.Workbook.Save();
                myWorkbook.Close();
            }
        }

        private static Sheet CreateSheet<T>(int sheetIndex, SheetDefinition<T> def, WorkbookPart workbookPart)
        {
            // create worksheet part
            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            var worksheetId = workbookPart.GetIdOfPart(worksheetPart);

            // variables
            var numCols = def.Fields.Count;
            var numRows = def.Objects.Count;
            var az = new List<char>(Enumerable.Range('A', 'Z' - 'A' + 1).Select(i => (char) i).ToArray());
            var headerCols = az.GetRange(0, numCols);
            var hasTitleRow = def.Title != null;
            var hasSubtitleRow = def.SubTitle != null;
            var titleRowCount = hasTitleRow ? 1 + (hasSubtitleRow ? 1 : 0) : hasSubtitleRow ? 1 : 0;

            // get the worksheet data
            int firstTableRow;
            var sheetData = CreateSheetData(def.Objects, def.Fields, headerCols, def.IncludeTotalsRow, def.Title,
                def.SubTitle,
                out firstTableRow);

            // populate column metadata
            var columns = new Columns();
            for (var col = 0; col < numCols; col++)
            {
                var width = ColumnWidth(sheetData, col, titleRowCount);
                columns.AppendChild(CreateColumnMetadata((uint) col + 1, (uint) numCols + 1, width));
            }

            // populate worksheet
            var worksheet = new Worksheet();
            worksheet.AppendChild(columns);
            worksheet.AppendChild(sheetData);

            // add an auto filter
            worksheet.AppendChild(new AutoFilter
            {
                Reference =
                    $"{headerCols.First()}{firstTableRow - 1}:{headerCols.Last()}{numRows + titleRowCount + 1}"
            });

            // add worksheet to worksheet part
            worksheetPart.Worksheet = worksheet;
            worksheetPart.Worksheet.Save();

            return new Sheet {Name = def.Name, SheetId = (uint) sheetIndex, Id = worksheetId};
        }

        private static Cell NumberCell(string header, string text, int index)
        {
            return new Cell
            {
                DataType = CellValues.Number,
                CellReference = header + index,
                CellValue = new CellValue(text)
            };
        }

        private static Cell DateCell(string header, DateTime dateTime, int index)
        {
            return new Cell
            {
                DataType = CellValues.Date,
                CellReference = header + index,
                StyleIndex = (uint) CustomStylesheet.CustomCellFormats.DefaultDate,
                CellValue = new CellValue(dateTime.ToString("yyyy-MM-dd"))
            };
        }

        private static Cell FormulaCell(string header, string text, int index)
        {
            return new Cell
            {
                CellFormula = new CellFormula {CalculateCell = true, Text = text},
                DataType = CellValues.Number,
                CellReference = header + index
            };
        }

        private static Cell TextCell(string header, string text, int index)
        {
            return new Cell
            {
                DataType = CellValues.InlineString,
                CellReference = header + index,
                InlineString = new InlineString {Text = new Text {Text = text}}
            };
        }

        private static Cell WithStyleIndex(this Cell cell, UInt32Value styleIndex)
        {
            cell.StyleIndex = styleIndex;
            return cell;
        }

        private static double ColumnWidth(SheetData sheetData, int col, int titleRowCount)
        {
            var rows = col == 0
                ? sheetData.ChildElements.ToList()
                    .GetRange(titleRowCount, sheetData.ChildElements.Count - titleRowCount)
                : sheetData.ChildElements.ToList();

            var maxLength =
            (from row in rows
                where row.ChildElements.Count > col
                select row.ChildElements[col] as Cell
                into cell
                where cell?.CellFormula != null //cell.GetType() != typeof (FormulaCell)
                select cell.InnerText.Length).Concat(new[] {0}).Max();
            var width = maxLength * 0.9 + 5;
            return width;
        }

        private static SheetData CreateSheetData<T>(IList<T> objects, List<SpreadsheetField> fields,
            List<char> headerCols, bool includedTotalsRow, string sheetTitle, string sheetSubTitle,
            out int firstTableRow)
        {
            var sheetData = new SheetData();
            var fieldNames = fields.Select(f => f.Title).ToList();
            var numCols = headerCols.Count;
            var rowIndex = 0;
            firstTableRow = 0;
            Row row;

            // create title
            if (sheetTitle != null)
            {
                rowIndex++;
                row = CreateTitle(sheetTitle, headerCols, ref rowIndex);
                sheetData.AppendChild(row);
            }

            // create subtitle
            if (sheetSubTitle != null)
            {
                rowIndex++;
                row = CreateSubTitle(sheetSubTitle, headerCols, ref rowIndex);
                sheetData.AppendChild(row);
            }

            // create the header
            rowIndex++;
            row = CreateHeader(fieldNames, headerCols, ref rowIndex);
            sheetData.AppendChild(row);

            if (objects.Count == 0)
                return sheetData;

            // create a row for each object and set the columns for each field
            firstTableRow = rowIndex + 1;
            CreateTable(objects, ref rowIndex, numCols, fields, headerCols, sheetData);

            // create an additional row with summed totals
            if (includedTotalsRow)
            {
                rowIndex++;
                AppendTotalsRow(objects, rowIndex, firstTableRow, numCols, fields, headerCols, sheetData);
            }

            return sheetData;
        }

        private static Row CreateTitle(string title, List<char> headerCols, ref int rowIndex)
        {
            var header = new Row {RowIndex = (uint) rowIndex, Height = 40, CustomHeight = true};
            header.AppendChild(TextCell(headerCols[0].ToString(), title, rowIndex)
                .WithStyleIndex((uint) CustomStylesheet.CustomCellFormats.TitleText));

            return header;
        }

        private static Row CreateSubTitle(string title, List<char> headerCols, ref int rowIndex)
        {
            var header = new Row {RowIndex = (uint) rowIndex, Height = 28, CustomHeight = true};

            var c = TextCell(headerCols[0].ToString(), title, rowIndex)
                .WithStyleIndex((uint) CustomStylesheet.CustomCellFormats.SubtitleText);
            header.AppendChild(c);

            return header;
        }

        private static Row CreateHeader(IList<string> headerNames, List<char> headerCols, ref int rowIndex)
        {
            var header = new Row {RowIndex = (uint) rowIndex};

            for (var col = 0; col < headerCols.Count; col++)
            {
                var c = TextCell(headerCols[col].ToString(), headerNames[col], rowIndex)
                    .WithStyleIndex((uint) CustomStylesheet.CustomCellFormats.HeaderText);
                header.AppendChild(c);
            }
            return header;
        }

        private static void CreateTable<T>(IList<T> objects, ref int rowIndex, int numCols,
            List<SpreadsheetField> fields, List<char> headers, SheetData sheetData, bool hidden = false, int outline = 0)
        {
            // for each object
            foreach (var rowObj in objects)
            {
                // row group?
                var list = rowObj as IList<object>;
                if (list != null)
                {
                    CreateTable(list, ref rowIndex, numCols, fields, headers, sheetData, true, outline + 1);
                    continue;
                }

                rowIndex++;

                // create a row
                var row = new Row
                {
                    RowIndex = (uint) rowIndex,
                    Collapsed = new BooleanValue(false),
                    OutlineLevel = new ByteValue((byte) outline),
                    Hidden = new BooleanValue(hidden)
                };

                int col;

                // populate columns using supplied objects
                for (col = 0; col < numCols; col++)
                {
                    var field = fields[col];
                    var columnObj = GetColumnObject(field.FieldName, rowObj);
                    if (columnObj == null || columnObj == DBNull.Value) continue;

                    Cell cell;

                    if (field.GetType() == typeof(HyperlinkField))
                    {
                        var displayColumnObj = GetColumnObject(((HyperlinkField) field).DisplayFieldName, rowObj);
                        cell = CreateHyperlinkCell(rowIndex, headers, columnObj, displayColumnObj, col);
                    }
                    else if (field.GetType() == typeof(DecimalNumberField))
                        cell = CreateDecimalNumberCell(rowIndex, headers, columnObj,
                            ((DecimalNumberField) field).DecimalPlaces, col);
                    else
                        cell = CreateCell(rowIndex, headers, columnObj, col);

                    row.AppendChild(cell);

                } // for each column

                sheetData.AppendChild(row);
            }
        }

        private static Cell CreateHyperlinkCell(int rowIndex, List<char> headers, object columnObj,
            object displayColumnObj, int col)
        {
            return FormulaCell(headers[col].ToString(),
                    $@"HYPERLINK(""{columnObj}"", ""{displayColumnObj}"")", rowIndex)
                .WithStyleIndex((uint) CustomStylesheet.CustomCellFormats.Hyperlink);
        }

        private static Cell CreateDecimalNumberCell(int rowIndex, List<char> headers, object columnObj,
            int decimalPlaces, int col)
        {
            var decStyle = decimalPlaces == 5
                ? (uint) CustomStylesheet.CustomCellFormats.DefaultNumber5DecimalPlace
                : (uint) CustomStylesheet.CustomCellFormats.DefaultNumber2DecimalPlace;
            return NumberCell(headers[col].ToString(), columnObj.ToString(), rowIndex)
                .WithStyleIndex(decStyle);
        }

        private static Cell CreateCell(int rowIndex, List<char> headers, object columnObj, int col)
        {
            Cell cell;
            if (columnObj is string)
                cell = TextCell(headers[col].ToString(), columnObj.ToString(), rowIndex);
            else if (columnObj is bool)
                cell = TextCell(headers[col].ToString(), (bool) columnObj ? "Yes" : "No", rowIndex);
            else if (columnObj is DateTime)
                cell = DateCell(headers[col].ToString(), (DateTime) columnObj, rowIndex);
            else if (columnObj is TimeSpan)
                // excel stores time as "fraction of hours in a day"
                cell = NumberCell(headers[col].ToString(), (((TimeSpan) columnObj).TotalHours / 24).ToString(), rowIndex)
                    .WithStyleIndex((uint) CustomStylesheet.CustomCellFormats.Duration);
            else if (columnObj is decimal || columnObj is double)
                cell = NumberCell(headers[col].ToString(), columnObj.ToString(), rowIndex)
                    .WithStyleIndex((uint) CustomStylesheet.CustomCellFormats.DefaultNumber2DecimalPlace);
            else
            {
                long value;
                cell = long.TryParse(columnObj.ToString(), out value)
                    ? NumberCell(headers[col].ToString(), columnObj.ToString(), rowIndex)
                    : TextCell(headers[col].ToString(), columnObj.ToString(), rowIndex);
            }
            return cell;
        }

        private static object GetColumnObject<T>(string fieldName, T rowObj)
        {
            // is the object a dictionary?
            var dict = rowObj as IDictionary<string, object>;
            if (dict != null)
            {
                object value;
                return dict.TryGetValue(fieldName, out value) ? value : null;
            }

            // get the properties for this object type
            var properties = GetPropertyInfo<T>();
            if (!properties.Contains(fieldName))
                return null;

            var myf = rowObj.GetType().GetProperty(fieldName);
            if (myf == null)
                return null;

            var obj = myf.GetValue(rowObj, null);
            return obj;
        }

        private static void AppendTotalsRow<T>(IList<T> objects, int rowIndex, int firstTableRow, int numCols,
            List<SpreadsheetField> fields,
            List<char> headers,
            SheetData sheetData)
        {
            var fieldNames = fields.Select(f => f.FieldName).ToList();
            var rowObj = objects[0];
            var total = new Row {RowIndex = (uint) rowIndex};

            for (var col = 0; col < numCols; col++)
            {
                var field = fields[col];
                if (field.IgnoreFromTotals)
                {
                    total.AppendChild(
                        TextCell(headers[col].ToString(), string.Empty, rowIndex)
                            .WithStyleIndex((uint) CustomStylesheet.CustomCellFormats.TotalsText)
                    );
                    continue;
                }

                var columnObject = GetColumnObject(fieldNames[col], rowObj);

                // look through objects until we have a value for this column
                var row = 0;
                while (columnObject == null || columnObject == DBNull.Value)
                {
                    if (objects.Count <= ++row)
                        break;
                    columnObject = GetColumnObject(fieldNames[col], objects[row]);
                }

                if (field.CountNoneNullRowsForTotal)
                    total.AppendChild(CreateRowTotalFomulaCell(rowIndex, firstTableRow, headers, col,
                        (uint) CustomStylesheet.CustomCellFormats.TotalsNumber, true));

                if (col == 0)
                    total.AppendChild(TextCell(headers[col].ToString(), "Total", rowIndex)
                        .WithStyleIndex((uint) CustomStylesheet.CustomCellFormats.TotalsText));
                else if (columnObject is decimal || columnObject is double)
                    total.AppendChild(CreateRowTotalFomulaCell(rowIndex, firstTableRow, headers, col,
                        (uint) CustomStylesheet.CustomCellFormats.TotalsNumber2DecimalPlace));
                else if (columnObject is TimeSpan)
                    total.AppendChild(CreateRowTotalFomulaCell(rowIndex, firstTableRow, headers, col,
                        (uint) CustomStylesheet.CustomCellFormats.TotalsDuration));
                else
                {
                    long value;
                    if (columnObject != null &&
                        long.TryParse(columnObject.ToString(), out value))
                    {
                        total.AppendChild(CreateRowTotalFomulaCell(rowIndex, firstTableRow, headers, col,
                            (uint) CustomStylesheet.CustomCellFormats.TotalsNumber));
                    }
                    else
                    {
                        total.AppendChild(TextCell(headers[col].ToString(), string.Empty, rowIndex)
                            .WithStyleIndex((uint) CustomStylesheet.CustomCellFormats.TotalsText));
                    }
                }
            } // for each column
            sheetData.AppendChild(total);
        }

        private static Cell CreateRowTotalFomulaCell(int rowIndex, int firstTableRow, List<char> headers, int col,
            uint styleIndex, bool countNonBlank = false)
        {
            var headerCol = headers[col].ToString();
            var firstRow = headerCol + firstTableRow;
            var lastRow = headerCol + (rowIndex - 1);
            return CreateFormulaCell(rowIndex, headers, col, styleIndex, countNonBlank, firstRow, lastRow);
        }

        private static Cell CreateFormulaCell(int rowIndex, List<char> headers, int col, uint styleIndex,
            bool countNonBlank, string firstCell, string lastCell)
        {
            var formula = (countNonBlank ? "COUNTA" : "SUM") + "(" + firstCell + ":" + lastCell + ")";
            return FormulaCell(headers[col].ToString(), formula, rowIndex)
                .WithStyleIndex(styleIndex);
        }

        private static List<string> GetPropertyInfo<T>()
        {
            var propertyInfos = typeof(T).GetProperties();
            return propertyInfos.Select(propertyInfo => propertyInfo.Name).ToList();
        }

        private static Column CreateColumnMetadata(uint startColumnIndex, uint endColumnIndex, double width)
        {
            var column = new Column
            {
                Min = startColumnIndex,
                Max = endColumnIndex,
                BestFit = true,
                Width = width,
            };
            return column;
        }
    }
}
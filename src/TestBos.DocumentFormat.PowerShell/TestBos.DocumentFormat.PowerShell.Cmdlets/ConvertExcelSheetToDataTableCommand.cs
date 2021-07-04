using System;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.Globalization;
using System.IO;
using System.Management.Automation;
// nuget
using CsvHelper;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;


namespace TestBos.DocumentFormat.PowerShell.Cmdlets
{
    [Cmdlet("Convert", "ExcelSheetToCSV")]
    [OutputType(typeof(String))]
    public class ConvertExcelSheetToDataTableCommand : Cmdlet
    {
        [Parameter(
            Position = 0,
            ValueFromPipelineByPropertyName = true,
            ValueFromPipeline = true)]
        public string Path { get; set; }

        [Parameter(
                Position = 1,
                ValueFromPipelineByPropertyName = true,
                ValueFromPipeline = true)]
        public string SheetName { get; set; }

        [Parameter(Position = 2)]
        public int StartRow { get; set; } = 1;

        [Parameter(Position = 3)]
        public int EndRow { get; set; }

        [Parameter(Position = 4)]
        public bool IncludeNullValue { get; set; } = true;

        // internal
        private static DataTable ReadSheetAsDataTable(
            string FilePath, 
            string SheetName, 
            int StartRow, 
            int? EndRow, 
            bool IncludeNullValue
        )
        {
            if (string.IsNullOrEmpty(FilePath))
            {
                throw new ArgumentException($"'{nameof(FilePath)}' cannot be null or empty.", nameof(FilePath));
            }

            if (string.IsNullOrEmpty(SheetName))
            {
                throw new ArgumentException($"'{nameof(SheetName)}' cannot be null or empty.", nameof(SheetName));
            }

            try
            {
                DataTable dtTable = new DataTable();

                using (SpreadsheetDocument doc = SpreadsheetDocument.Open(FilePath, false))
                {
                    WorkbookPart workbookPart = doc.WorkbookPart;
                    Sheets thesheetcollection = workbookPart.Workbook.GetFirstChild<Sheets>();

                    string relId = workbookPart.Workbook.Descendants<Sheet>().First(s => SheetName.Equals(s.Name)).Id;
                    Worksheet theWorksheet = ((WorksheetPart)workbookPart.GetPartById(relId)).Worksheet;
                    SheetData thesheetdata = theWorksheet.GetFirstChild<SheetData>();
                    int totalCnt = EndRow ?? thesheetdata.ChildElements.Count;

                    for (int rCnt = StartRow-1; rCnt < totalCnt; rCnt++)
                    {
                        List<string> rowList = new List<string>();
                        for (int rCnt1 = 0; rCnt1 < thesheetdata.ElementAt(rCnt).ChildElements.Count; rCnt1++)
                        {
                            Cell thecurrentcell = (Cell)thesheetdata.ElementAt(rCnt).ChildElements.ElementAt(rCnt1);
                            string currentcellvalue = string.Empty;
                            if (thecurrentcell.DataType != null)
                            {
                                switch (thecurrentcell.DataType.Value)
                                {
                                    case CellValues.SharedString:
                                        int id;
                                        if (int.TryParse(thecurrentcell.InnerText, out id))
                                        {
                                            SharedStringItem item = workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(id);
                                            if (item.Text != null)
                                            {
                                                // header
                                                if (rCnt == StartRow - 1)
                                                {
                                                    dtTable.Columns.Add(item.Text.Text);
                                                }
                                                else
                                                {
                                                    if (rowList.Count < dtTable.Columns.Count)
                                                    {
                                                        var value = item.Text.Text;
                                                        if (IncludeNullValue == true)
                                                        {
                                                            rowList.Add(value);
                                                        }
                                                        else
                                                        {
                                                            if (string.IsNullOrEmpty(value) == false)
                                                            {
                                                                rowList.Add(value);
                                                            }
                                                        }
                                                    }
                                                    else
                                                        break;
                                                }
                                            }
                                        };
                                        break;

                                    case CellValues.String:
                                        if (rowList.Count < dtTable.Columns.Count)
                                        {
                                            var value = thecurrentcell.CellValue.InnerText;
                                            if (IncludeNullValue == true)
                                            {
                                                rowList.Add(value);
                                            }
                                            else
                                            {
                                                if (string.IsNullOrEmpty(value) == false)
                                                {
                                                    rowList.Add(value);
                                                }
                                            }
                                        }
                                        break;

                                    default:
                                        break;
                                }

                            }
                            else
                            {
                                if (rCnt != StartRow-1)
                                {
                                    if (rowList.Count < dtTable.Columns.Count)
                                    {
                                        //var value = thecurrentcell.InnerText;
                                        string value;
                                        if (thecurrentcell.CellValue != null)
                                            value = thecurrentcell.CellValue.InnerText.ToString();
                                        else
                                            value = "";

                                        if (IncludeNullValue == true)
                                        {
                                            rowList.Add(value);
                                        }
                                        else
                                        {
                                            if (string.IsNullOrEmpty(value) == false)
                                            {
                                                rowList.Add(value);
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        if (rCnt != StartRow-1)
                        {
                            if (rowList.Count > 0)
                            {
                                dtTable.Rows.Add(rowList.ToArray());
                            }
                        }
                    }

                    return dtTable;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                throw;
            }
        }
      
        private static string WriteDataTableToCSVString(DataTable Table)
        {
            if (Table is null)
            {
                throw new ArgumentNullException(nameof(Table));
            }

            try
            {
                StringWriter csvString = new StringWriter();
                using (var csv = new CsvWriter(csvString, CultureInfo.InvariantCulture))
                {
                    foreach (DataColumn column in Table.Columns)
                    {
                        csv.WriteField(column.ColumnName);
                    }
                    csv.NextRecord();

                    foreach (DataRow row in Table.Rows)
                    {
                        for (var i = 0; i < Table.Columns.Count; i++)
                        {
                            csv.WriteField(row[i]);
                        }
                        csv.NextRecord();
                    }

                }
                return csvString.ToString();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                throw;
            }
        }

        protected override void ProcessRecord()
        {
            WriteObject(WriteDataTableToCSVString(ReadSheetAsDataTable(Path, SheetName, StartRow, EndRow, IncludeNullValue)));
        }
    }
}
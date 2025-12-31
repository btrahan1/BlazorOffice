using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text;

namespace BlazorOffice.Services
{
    public class CellData
    {
        public string Value { get; set; } = "";
        public string Formula { get; set; } = "";

        // Display logic: Show formula if present, else value
        public string DisplayValue => !string.IsNullOrEmpty(Formula) ? $"={Formula}" : Value;
    }

    public class ExcelService
    {
        // Reads the first sheet of an XLSX file into a grid of CellData
        public async Task<List<List<CellData>>> ReadXlsxAsync(Stream stream)
        {
            var rows = new List<List<CellData>>();
            
            using var memStream = new MemoryStream();
            await stream.CopyToAsync(memStream);
            memStream.Position = 0;

            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(memStream, false))
            {
                WorkbookPart? workbookPart = doc.WorkbookPart;
                WorksheetPart? worksheetPart = workbookPart?.WorksheetParts.FirstOrDefault();
                if (workbookPart == null || worksheetPart == null) return rows;

                SharedStringTablePart? sstPart = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                SharedStringTable? sst = sstPart?.SharedStringTable;

                SheetData? sheetData = worksheetPart.Worksheet.Elements<SheetData>().FirstOrDefault();
                if (sheetData == null) return rows;

                foreach (Row r in sheetData.Elements<Row>())
                {
                    var rowData = new List<CellData>();
                    
                    foreach (Cell c in r.Elements<Cell>())
                    {
                        rowData.Add(GetCellData(c, sst, workbookPart));
                    }
                    rows.Add(rowData);
                }
            }

            return rows;
        }

        private CellData GetCellData(Cell cell, SharedStringTable? sst, WorkbookPart workbookPart)
        {
            var data = new CellData();

            // 1. Get Formula
            if (cell.CellFormula != null)
            {
                data.Formula = cell.CellFormula.Text;
            }

            // 2. Get Value
            if (cell.CellValue != null)
            {
                string value = cell.CellValue.InnerXml;
                if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
                {
                    if (sst != null && int.TryParse(value, out int index))
                    {
                        value = sst.ElementAt(index).InnerText;
                    }
                }
                else
                {
                    // Check for Date Formatting
                    if (cell.StyleIndex != null && workbookPart.WorkbookStylesPart?.Stylesheet?.CellFormats != null)
                    {
                        uint styleIndex = cell.StyleIndex.Value;
                        var cellFormats = workbookPart.WorkbookStylesPart.Stylesheet.CellFormats;
                        
                        if (styleIndex < cellFormats.Count())
                        {
                            var cellFormat = (CellFormat)cellFormats.ElementAt((int)styleIndex);
                            if (cellFormat.NumberFormatId != null)
                            {
                                uint numberFormatId = cellFormat.NumberFormatId.Value;
                                
                                // Standard Date Formats (14-22)
                                // or Custom Formats (164+) which we can't easily validate without parsing NumberingFormats
                                // For now, we assume simple date detection (ID 14 is default Date)
                                if (numberFormatId >= 14 && numberFormatId <= 22)
                                {
                                    if (double.TryParse(value, out double oaDate))
                                    {
                                        value = DateTime.FromOADate(oaDate).ToShortDateString();
                                    }
                                }
                            }
                        }
                    }
                }
                data.Value = value;
            }

            return data;
        }

        // Saves a grid of CellData to an XLSX file
        public async Task<byte[]> SaveXlsxAsync(List<List<CellData>> data)
        {
            using var memStream = new MemoryStream();
            using (SpreadsheetDocument doc = SpreadsheetDocument.Create(memStream, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookPart = doc.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());

                Sheets sheets = doc.WorkbookPart.Workbook.AppendChild(new Sheets());
                Sheet sheet = new Sheet() 
                { 
                    Id = doc.WorkbookPart.GetIdOfPart(worksheetPart), 
                    SheetId = 1, 
                    Name = "Sheet1" 
                };
                sheets.Append(sheet);

                SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

                for (int i = 0; i < data.Count; i++)
                {
                    Row row = new Row();
                    for (int j = 0; j < data[i].Count; j++)
                    {
                        var cellData = data[i][j];
                        Cell cell = new Cell();

                        // Write Formula if present
                        if (!string.IsNullOrEmpty(cellData.Formula))
                        {
                            cell.CellFormula = new CellFormula(cellData.Formula);
                            // We don't calculate values here (Excel will do it on open)
                        }
                        else
                        {
                            cell.DataType = CellValues.String;
                            cell.CellValue = new CellValue(cellData.Value);
                        }

                        row.Append(cell);
                    }
                    sheetData.Append(row);
                }
                
                workbookPart.Workbook.Save();
            }

            return memStream.ToArray();
        }
    }
}

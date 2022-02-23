using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace Phone_book_Formatter
{
    internal class Program
    {
        private static List<Tuple<string, string>> identifiers;
        private static int CheckLength;

        private static void LoadCodes()
        {
            string text = "null";
            if (File.Exists(Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + "\\conf1.json"))
                text = File.ReadAllText(Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + "\\conf1.json");
            identifiers = new List<Tuple<string, string>>();
            if (text != "null")
                Newtonsoft.Json.JsonConvert.PopulateObject(text, identifiers);
            text = "x";
            if (File.Exists(Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + "\\conf2.json"))
                text = File.ReadAllText(Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + "\\conf2.json");
            while (!int.TryParse(text, out CheckLength))
            {
                Console.Write("A valid ID length was not provided. ID Length is the number of digits at the start of a number that should be used to define a phones location.\nProvide ID Length\n?>");
                text = Console.ReadLine();
            }
            identifiers ??= new List<Tuple<string, string>>(0);
        }

        private static void SaveCodes()
        {
            TextWriter text = File.CreateText(Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + "\\conf1.json");
            text.Write(Newtonsoft.Json.JsonConvert.SerializeObject(identifiers));
            text.Flush();
            text.Close();
            text = File.CreateText(Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + "\\conf2.json");
            text.Write(Newtonsoft.Json.JsonConvert.SerializeObject(CheckLength));
            text.Flush();
            text.Close();
        }

        private static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                Console.WriteLine("Drop file on the EXE to reformat.");
                Console.ReadLine();
                Environment.Exit(0);
            }

            LoadCodes();

            int lines = 45;

            if (args.Length > 1)
                lines = Int32.Parse(args[1]) + 1;

            Console.WriteLine($"Parsing file {args[0]}");

            // Stores all the extensions
            List<Extension> extensions = new(0);
            StreamReader reader = new(args[0]);
            reader.ReadLine();
            while (!reader.EndOfStream)
            {
                string line = reader.ReadLine();
                extensions.Add(new Extension()
                {
                    No = line.Split(',')[0],
                    Name = line.Split(',')[1] + " " + line.Split(',')[2],
                    IDLength = CheckLength
                });
            }

            // Sorts the numbers based on the first two digits
            List<Tuple<List<Extension>, string>> extensionsSplit = new(0);

            foreach (Extension e in extensions)
            {
                bool added = false;
                for (int i = 0; i < extensionsSplit.Count; i++)
                    if (extensionsSplit[i].Item2 == e.ID)
                    {
                        extensionsSplit[i].Item1.Add(e);
                        added = true;
                        break;
                    }
                if (!added)
                {
                    extensionsSplit.Add(new Tuple<List<Extension>, string>(new List<Extension> { e }, e.ID));
                }
            }
            List<Tuple<string, string>> cats = new(0);
            // Gets some details from user input
            foreach (Tuple<List<Extension>, string> tuple in extensionsSplit)
            {
                Console.Clear();
                foreach (Tuple<string, string> tuple1 in cats)
                {
                    Console.WriteLine($"{tuple1.Item1} -> {tuple1.Item2}");
                }
                if (identifiers.Any(value => value.Item1 == tuple.Item2))
                {
                    cats.Add(new Tuple<string, string>(tuple.Item2, identifiers.Find(value => value.Item1 == tuple.Item2).Item2));
                    continue;
                }
                Console.Write($"Ignore ID (/I)\n{tuple.Item2}>");
                string resp = Console.ReadLine();
                cats.Add(new Tuple<string, string>(tuple.Item2, resp));
                if (tuple.Item1.Count > 0)
                    identifiers.Add(new Tuple<string, string>(tuple.Item1[0].ID, resp));
            }

            SaveCodes();

            foreach (Tuple<string, string> cat in cats)
            {
                if (cat.Item2.ToUpper() == "/I")
                {
                    extensionsSplit.RemoveAll(value => value.Item2 == cat.Item1);
                }
            }

            // Prepares the layout of the sheet.
            List<List<Tuple<string, string>>> printData = new()
            {
                new List<Tuple<string, string>>(lines)
            };
            int rows = 0;
            int column = 0;
            foreach (Tuple<List<Extension>, string> extensionSorted in extensionsSplit)
            {
                if (rows + 10 >= lines)
                {
                    column++;
                    rows = 0;
                    printData.Add(new List<Tuple<string, string>>(lines));
                }
                foreach (Tuple<string, string> title in cats)
                {
                    if (title.Item1 == extensionSorted.Item2)
                    {
                        if (rows != 0)
                        {
                            printData[column].Add(new Tuple<string, string>("***FILL***", "***FILL***"));
                            rows++;
                        }
                        printData[column].Add(new Tuple<string, string>(title.Item2, title.Item1));
                        rows++;
                    }
                }
                foreach (Extension extensionToAdd in extensionSorted.Item1)
                {
                    printData[column].Add(new Tuple<string, string>(extensionToAdd.Name, extensionToAdd.No));
                    rows++;
                    if (rows >= lines)
                    {
                        column++;
                        rows = 0;
                        printData.Add(new List<Tuple<string, string>>(lines));
                    }
                }
            }

            // Works using OpenXML
            //OpenXML(printData, lines);

            // Uses NPOI
            Npoi(printData, lines);
        }

        public static void Npoi(List<List<Tuple<string, string>>> printData, int lines)
        {
            IWorkbook workBook = new XSSFWorkbook();
            ISheet excelSheet = workBook.CreateSheet("Extensions");
            IFont TitleFont = workBook.CreateFont();
            IFont BoldFont = workBook.CreateFont();
            TitleFont.IsBold = true;
            BoldFont.IsBold = true;
            TitleFont.Color = NPOI.HSSF.Util.HSSFColor.Yellow.Index;
            ICellStyle TitleCell = workBook.CreateCellStyle();
            ICellStyle BoldCell = workBook.CreateCellStyle();
            ICellStyle BlackCell = workBook.CreateCellStyle();
            TitleCell.SetFont(TitleFont);
            BoldCell.SetFont(BoldFont);
            TitleCell.FillPattern = FillPattern.SolidForeground;
            BlackCell.FillPattern = FillPattern.SolidForeground;
            TitleCell.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            BlackCell.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            int columnIndex;
            IRow cRow = excelSheet.CreateRow(0);
            for (int q = 0; q < (printData.Count * 3) + 1; q++)
            {
                cRow.CreateCell(q).CellStyle = BlackCell;
            }
            for (int i = 0; i < lines; i++)
            {
                cRow = excelSheet.CreateRow(i + 1);
                columnIndex = 0;
                foreach (List<Tuple<string, string>> _column in printData)
                {
                    cRow.CreateCell(columnIndex).CellStyle = BlackCell;
                    excelSheet.SetColumnWidth(columnIndex, 250);
                    columnIndex++;
                    cRow.CreateCell(columnIndex);
                    excelSheet.SetColumnWidth(columnIndex, 4000);
                    columnIndex++;
                    cRow.CreateCell(columnIndex);
                    excelSheet.SetColumnWidth(columnIndex, 2000);
                    if (i < _column.Count)
                    {
                        cRow.GetCell(columnIndex - 1).SetCellValue(_column[i].Item1);
                        cRow.GetCell(columnIndex).SetCellValue(_column[i].Item2);
                        if (_column[i].Item2.Length == 2)
                        {
                            cRow.GetCell(columnIndex - 1).CellStyle = TitleCell;
                            cRow.GetCell(columnIndex).CellStyle = TitleCell;
                        }
                        else if (_column[i].Item1.ToString().Equals("***FILL***"))
                        {
                            cRow.CreateCell(columnIndex - 1).CellStyle = BlackCell;
                            cRow.CreateCell(columnIndex).CellStyle = BlackCell;
                        }
                        else
                        {
                            cRow.GetCell(columnIndex - 1).CellStyle = BoldCell;
                            cRow.GetCell(columnIndex).CellStyle = BoldCell;
                        }
                    }
                    else
                    {
                        cRow.GetCell(columnIndex - 1).CellStyle = BlackCell;
                        cRow.GetCell(columnIndex).CellStyle = BlackCell;
                    }
                    columnIndex++;
                }
                cRow.CreateCell(columnIndex).CellStyle = BlackCell;
                excelSheet.SetColumnWidth(columnIndex, 250);
            }
            cRow = excelSheet.CreateRow(lines);
            for (int q = 0; q < (printData.Count * 3) + 1; q++)
            {
                cRow.CreateCell(q).CellStyle = BlackCell;
            }
            string path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\Phonebook.xlsx";
            if (File.Exists(path))
                File.Delete(path);
            FileStream fs = new FileStream(path, FileMode.CreateNew, FileAccess.Write);
            workBook.Write(fs);
        }

        public static void OpenXML(List<List<Tuple<string, string>>> printData, int lines)
        {
            // Interface with Excel
            string path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\Phonebook.xlsx";
            using SpreadsheetDocument dococument = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
            WorkbookPart workBookPart = dococument.AddWorkbookPart();
            workBookPart.Workbook = new Workbook();

            WorksheetPart workSheetPart = workBookPart.AddNewPart<WorksheetPart>();

            var sheetData = new SheetData();
            workSheetPart.Worksheet = new Worksheet(sheetData);

            Sheets sheets = workBookPart.Workbook.AppendChild(new Sheets());
            Sheet sheet = new() { Id = workBookPart.GetIdOfPart(workSheetPart), SheetId = 1, Name = "Sheet1" };

            sheets.Append(sheet);
            sheetData.Append(new Row());
            for (int i = 0; i < lines; i++)
            {
                Row newRow = new();
                foreach (List<Tuple<string, string>> _column in printData)
                {
                    newRow.AppendChild(new Cell());
                    Cell cell1 = new();
                    Cell cell2 = new();
                    cell1.DataType = CellValues.String;
                    cell2.DataType = CellValues.String;
                    if (i < _column.Count)
                    {
                        cell1.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(_column[i].Item1);
                        cell2.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(_column[i].Item2);
                    }
                    newRow.AppendChild(cell1);
                    newRow.AppendChild(cell2);
                }
                sheetData.AppendChild(newRow);
            }

            workBookPart.Workbook.Save();
        }
    }
}
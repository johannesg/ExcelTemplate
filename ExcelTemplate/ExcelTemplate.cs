using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Spreadsheet;
using System.IO;
using DocumentFormat.OpenXml;

namespace ExcelTemplate
{
    public class ExcelTemplateException : Exception
    {
        public ExcelTemplateException(string message)
          : base(message)
        {

        }
    }

    public class ExcelTemplate : IDisposable
    {
        SpreadsheetDocument _doc;

        private WorkbookPart _workbookPart;

        private MemoryStream _inMemoryStream;
        private readonly Stream _outputStream;
        private IDictionary<string, DefinedNameValue> _definedNames;

        public ExcelTemplate(string templatePath, Stream output)
        {
            OpenTemplate(templatePath);
            _outputStream = output;
        }

        public void OpenTemplate(string filename)
        {
            _inMemoryStream = GetMemoryStream(filename);

            _doc = SpreadsheetDocument.Open(_inMemoryStream, true);

            _workbookPart = _doc.WorkbookPart;

            _definedNames = GetDefinedNames();
        }

        private DefinedNameValue GetDefinedName(string name)
        {
            DefinedNameValue definedName;
            if (!_definedNames.TryGetValue(name, out definedName))
                throw new ExcelTemplateException("Template row not found");

            return definedName;
        }

        private MemoryStream GetMemoryStream(string filename)
        {
            var stream = new MemoryStream();

            using (var templateStream = File.OpenRead(filename))
            {
                var buf = new byte[8 * 1024];

                var readLen = templateStream.Read(buf, 0, buf.Length);

                while (readLen > 0)
                {
                    stream.Write(buf, 0, readLen);
                    readLen = templateStream.Read(buf, 0, buf.Length);
                }

                stream.Position = 0;
            }

            return stream;
        }

        private RowTemplate GetRowTemplate(string templateName)
        {
            var templateRowName = GetDefinedName(templateName);

            if (templateRowName == null)
                throw new ExcelTemplateException(String.Format("Template {0} not found", templateName));

            var worksheetPart = GetWorksheetPart(templateRowName);

            return new RowTemplate(_workbookPart, worksheetPart, templateRowName);
        }

        private WorksheetPart GetWorksheetPart(DefinedNameValue definedName)
        {
            var sheet = _workbookPart.Workbook.Descendants<Sheet>().SingleOrDefault(x => x.Name == definedName.SheetName);

            if (sheet == null)
                throw new ExcelTemplateException(String.Format("Worksheet {0} not found", definedName.SheetName));

            return (WorksheetPart)_workbookPart.GetPartById(sheet.Id);
        }

        private IDictionary<string, DefinedNameValue> GetDefinedNames()
        {
            var r = new Regex(@"(?<SheetName>.*)!\$(?<StartCol>[A-Z]+)\$(?<StartRow>\d+)(:\$(?<EndCol>[A-Z]+)\$(?<EndRow>\d+))?");

            var definedNames = new Dictionary<string, DefinedNameValue>();

            foreach (DefinedName definedName in _workbookPart.Workbook.GetFirstChild<DefinedNames>())
            {
                var m = r.Match(definedName.InnerText);
                var sheetName = m.Groups["SheetName"].Value;
                var startCol = m.Groups["StartCol"].Value;
                var startRow = m.Groups["StartRow"].Value;
                var endCol = m.Groups["EndCol"].Value;
                var endRow = m.Groups["EndRow"].Value;

                definedNames.Add(definedName.Name, new DefinedNameValue { SheetName = sheetName, StartCol = startCol, StartRow = startRow, EndCol = endCol, EndRow = endRow });
            }

            return definedNames;
        }

        public void Dispose()
        {
            Save();
        }

        public void WriteField(string definedName, string value)
        {
            var dfv = GetDefinedName(definedName);

            if (dfv == null)
                throw new ExcelTemplateException(String.Format("Template {0} not found", definedName));

            var cellRef = $"{dfv.StartCol}{dfv.StartRow}";
            var worksheetPart = GetWorksheetPart(dfv);
            var cell = worksheetPart.Worksheet.Descendants<Cell>().SingleOrDefault(x => x.CellReference.Value == cellRef);
            if (cell == null)
                return;

            cell.CellValue = new CellValue(value);
            cell.DataType = new EnumValue<CellValues>(CellValues.String);
        }

        public void WriteObjects<T>(IEnumerable<T> objects)
        {
            WriteObjects("TemplateRow", objects);
        }

        public void WriteObjects<T>(string templateName, IEnumerable<T> objects)
        {
            var rowTemplate = GetRowTemplate(templateName);

            foreach (var o in objects)
                rowTemplate.InsertRow(o);

            rowTemplate.Save();
        }

        private void CloseDocument()
        {
            if (_doc != null)
            {
                _doc.Close();
                _doc.Dispose();
                _doc = null;
            }
        }

        private void Save()
        {
            CloseDocument();

            _inMemoryStream.WriteTo(_outputStream);
        }
    }
}

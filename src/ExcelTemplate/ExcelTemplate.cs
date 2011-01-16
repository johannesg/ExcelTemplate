using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Spreadsheet;
using System.IO;

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

    public RowTemplate RowTemplate { get; set; }

    private SheetData _sheetData;
    private WorkbookPart _workbookPart;
    private WorksheetPart _worksheetPart;

    private DefinedNameValue _definedNameValue;
    private MemoryStream _inMemoryStream;
    private Stream _outputStream;
    private int _currentRowIndex;

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

      var definedNames = GetDefinedNames();

      if (!definedNames.TryGetValue("TemplateRow", out _definedNameValue))
        throw new ExcelTemplateException("Template row not found");

      _worksheetPart = GetWorksheetPart();

      _sheetData = _worksheetPart.Worksheet.Descendants<SheetData>().SingleOrDefault();

      RowTemplate = GetRowTemplate();

      if (RowTemplate == null)
        throw new ExcelTemplateException(String.Format("Sheet {0} not found", _definedNameValue.SheetName));
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

    private RowTemplate GetRowTemplate()
    {
      _currentRowIndex = Int32.Parse(_definedNameValue.StartRow);
      var templateRow = _sheetData.Descendants<Row>().SingleOrDefault(x => x.RowIndex == _currentRowIndex);
      
      return new RowTemplate(_workbookPart, templateRow);
    }

    private WorksheetPart GetWorksheetPart()
    {
      var sheet = _workbookPart.Workbook.Descendants<Sheet>().SingleOrDefault(x => x.Name == _definedNameValue.SheetName);

      if (sheet == null)
        throw new ExcelTemplateException(String.Format("Worksheet {0} not found", _definedNameValue.SheetName));

      return (WorksheetPart)_workbookPart.GetPartById(sheet.Id);
    }

    private Row GetTemplateRow(DefinedNameValue definedNameValue)
    {
      return null;
    }

    private IDictionary<string, DefinedNameValue> GetDefinedNames()
    {
      var r = new Regex(@"(?<SheetName>.*)!\$(?<StartCol>[A-Z]+)\$(?<StartRow>\d+):\$(?<EndCol>[A-Z]+)\$(?<EndRow>\d+)");

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

    public void WriteObjects<T>(IEnumerable<T> objects)
    {
      foreach (var o in objects)
      {
        var row = RowTemplate.CreateRow(_currentRowIndex, o);

        RowTemplate.Row.InsertBeforeSelf(row);
        _currentRowIndex++;
      }

      _worksheetPart.Worksheet.Save();
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
      RowTemplate.Row.Remove();

      CloseDocument();

      _inMemoryStream.WriteTo(_outputStream);
    }
  }
}

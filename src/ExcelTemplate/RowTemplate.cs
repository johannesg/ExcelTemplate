using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using System.ComponentModel;
using System.Globalization;

namespace ExcelTemplate
{
  public class RowTemplate
  {
    public Row Row { get; set; }
    private readonly WorkbookPart _workbookPart;
    private readonly WorksheetPart _worksheetPart;
    private readonly SheetData _sheetData;
    private int _currentRowIndex;

    public List<Cell> Cells { get; set; }

    public RowTemplate(WorkbookPart workbookPart, WorksheetPart worksheetPart, DefinedNameValue templateRowName)
    {
      //      Row = templateRow;
      _workbookPart = workbookPart;
      _worksheetPart = worksheetPart;

      _sheetData = worksheetPart.Worksheet.Descendants<SheetData>().SingleOrDefault();

      _currentRowIndex = Int32.Parse(templateRowName.StartRow);
      Row = _sheetData.Descendants<Row>().SingleOrDefault(x => x.RowIndex == _currentRowIndex);
    }

    public void Save()
    {
      Row.Remove();
      _worksheetPart.Worksheet.Save();
    }

    public Row CreateRow(int index, object value)
    {
      var regex = new Regex(@"\[(\w+)\]");

      var row = new Row();
      row.RowIndex = (uint)index;
      row.CustomHeight = Row.CustomHeight;
      row.CustomFormat = Row.CustomFormat;
      row.Height = Row.Height;

      foreach (var cell in Row.Descendants<Cell>())
      {
        var cellValue = GetCellValue(cell);

        var m = regex.Match(cellValue);

        if (!m.Success)
          continue;

        var propertyName = m.Groups[1].Value;

        //        Console.WriteLine("propertyname: {0}", propertyName);

        var propertyValue = GetPropertyValue(value, propertyName);

        if (propertyValue == null)
          continue;

        //        Console.WriteLine("propertyValue: {0}", propertyValue);

        Cell newCell = CreateCell(index, cell, propertyValue);

        row.AppendChild(newCell);
      }

      return row;
    }

    private Cell CreateCell(int index, Cell cellTemplate, object propertyValue)
    {
      var newCell = new Cell();
      newCell.StyleIndex = cellTemplate.StyleIndex;
      newCell.CellReference = GetNewCellReference(index, cellTemplate.CellReference.InnerText);

      CellValue cellValue;

      if (propertyValue is string)
      {
        newCell.DataType = CellValues.SharedString;
        int sharedStringId = CreateSharedString((string)propertyValue);
        cellValue = new CellValue(sharedStringId.ToString());
      }
      else
        cellValue = new CellValue(Convert.ToString(propertyValue, CultureInfo.InvariantCulture));

      newCell.AppendChild(cellValue);

      return newCell;
    }

    private int CreateSharedString(string p)
    {
      var stringTablePart = _workbookPart.SharedStringTablePart;

      if (stringTablePart.SharedStringTable == null)
        stringTablePart.SharedStringTable = new SharedStringTable();

      int index = 0;

      foreach (var stringItem in stringTablePart.SharedStringTable.Elements<SharedStringItem>())
      {
        if (stringItem.InnerText == p)
          return index;

        index++;
      }

      stringTablePart.SharedStringTable.AppendChild<SharedStringItem>(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(p)));
      stringTablePart.SharedStringTable.Save();

      return index;
    }

    private string GetNewCellReference(int index, string cellReference)
    {
      var match = Regex.Match(cellReference, @"([A-Z]+)\d+");

      var column = match.Groups[1];

      return string.Format("{0}{1}", column, index);
    }

    private object GetPropertyValue(object value, string propertyName)
    {
      var properties = TypeDescriptor.GetProperties(value);

      var descriptor = properties.Cast<PropertyDescriptor>().SingleOrDefault(x => x.Name == propertyName);

      if (descriptor == null)
        return null;

      return descriptor.GetValue(value);
    }

    private string GetCellValue(Cell cell)
    {
      var value = cell.CellValue.InnerText;

      if (cell.DataType == CellValues.SharedString)
        value = _workbookPart.SharedStringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)].InnerText;

      return value;
    }

    public OpenXmlElement InsertRow(object o)
    {
      var row = CreateRow(_currentRowIndex, o);

      MoveAllRowsAfter(_currentRowIndex);

      Row.InsertBeforeSelf(row);
      _currentRowIndex++;

      return row;
    }

    private void MoveAllRowsAfter(int currentRowIndex)
    {
      uint newRowIndex;

      IEnumerable<Row> rows = _sheetData.Descendants<Row>().Where(r => r.RowIndex.Value >= currentRowIndex);
      foreach (Row row in rows)
      {
        newRowIndex = System.Convert.ToUInt32(row.RowIndex.Value + 1);

        foreach (Cell cell in row.Elements<Cell>())
        {
          string cellReference = cell.CellReference.Value;
          cell.CellReference = new StringValue(cellReference.Replace(row.RowIndex.Value.ToString(), newRowIndex.ToString()));
        }

        row.RowIndex = new UInt32Value(newRowIndex);
      }
    }
  }
}


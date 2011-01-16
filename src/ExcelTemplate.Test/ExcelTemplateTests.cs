using System;
using System.Collections.Generic;
using NUnit.Framework;
using System.IO;

namespace ExcelTemplate.Test
{
  [TestFixture]
  public class ExcelTemplateTests
  {
    [Test]
    public void Test()
    {
      using (var stream = File.Create("Output.xlsx"))
      using (var generator = new ExcelTemplate(@"Template.xlsx", stream))
      {
        generator.WriteObjects(GetTestData());
      }
    }

    [Test]
    public void TestMultipleRows()
    {
      using (var stream = File.Create("OutputMultipleRows.xlsx"))
      using (var generator = new ExcelTemplate(@"TemplateMultipleRows.xlsx", stream))
      {
        generator.WriteObjects("TemplateRow2", GetTestData2());
        generator.WriteObjects("TemplateRow", GetTestData());
      }
    }

    private static List<object> GetTestData()
    {
      var data = new List<object>();
      var rand = new Random();

      for (int i = 0; i < 20; i++)
      {
        data.Add(new
        {
          DistanceFrom = i * 20,
          DistanceTo = i * 20,
          WeightGroup1 = rand.Next(5, 50),
          WeightGroup2 = rand.Next(5, 50),
          WeightGroup3 = rand.Next(5, 50),
          WeightGroup4 = rand.Next(5, 50),
          WeightGroup5 = rand.Next(5, 50),
          WeightGroup6 = rand.Next(5, 50),
          WeightGroup7 = rand.Next(5, 50),
          WeightGroup8 = rand.Next(5, 50)
        });
      }

      return data;
    }

    private static List<object> GetTestData2()
    {
      var data = new List<object>();
      var rand = new Random();

      for (int i = 0; i < 20; i++)
      {
        data.Add(new
        {
          C1 = string.Format("Row C1: {0}", i),
          C2 = rand.Next(5, 50),
          C3 = rand.Next(5, 50),
          C4 = rand.Next(5, 50),
          C5 = rand.Next(5, 50),
          C6 = rand.Next(5, 50),
        });
      }

      return data;
    }
  }
}

using NUnit.Framework;
using SharpTestsEx;
using System.IO;

namespace ExcelTemplate.Test
{
  [TestFixture]
  class RowTemplateTest
  {

    [Test]
    public void Test()
    {

      using (var stream = new MemoryStream())
      using (var generator = new ExcelTemplate(@"Template.xlsx", stream))
      {
        var row = generator.RowTemplate.CreateRow(6, new { DistanceFrom = "123", DistanceTo = "321" });

        row.Should().Not.Be.Null();
      }
    }
  }
}

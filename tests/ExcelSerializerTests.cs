using System.Collections.Generic;
using System.IO;
using Xunit;
using ExcelImportExport;

namespace ExcelImportExport.Tests
{
    public class ExampleModel
    {
        public int Id { get; set; }
        public string Name { get; set; } = string.Empty;
    }

    public class ExcelSerializerTests
    {
        [Fact]
        public void ExportThenImport_ShouldReturnEquivalentObjects()
        {
            // Arrange
            var serializer = new ExcelSerializer();
            var originalList = new List<ExampleModel>
            {
                new ExampleModel { Id = 1, Name = "Alice" },
                new ExampleModel { Id = 2, Name = "Bob" }
            };

            //var tempFile = Path.GetTempFileName();
            var tempFile = Path.Combine(Directory.GetCurrentDirectory(), "teste-output.xlsx");

            File.Delete(tempFile);
            tempFile = Path.ChangeExtension(tempFile, ".xlsx");

            try
            {
                // Act
                serializer.ExportToExcel(originalList, tempFile);
                var importedList = serializer.ImportFromExcel<ExampleModel>(tempFile);

                // Assert
                Assert.Equal(originalList.Count, importedList.Count);
                for (int i = 0; i < originalList.Count; i++)
                {
                    Assert.Equal(originalList[i].Id, importedList[i].Id);
                    Assert.Equal(originalList[i].Name, importedList[i].Name);
                }
            }
            finally
            {
                //if (File.Exists(tempFile))
                //    File.Delete(tempFile);
            }
        }
    }
}

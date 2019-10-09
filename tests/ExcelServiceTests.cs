using System;
using Xunit;
using Ltwlf.Functions.Excel;
using System.IO;
using Newtonsoft.Json;
using System.Collections.Generic;

namespace Api.Tests
{
    public class ExcelServiceTests
    {
        [Fact]
        public void GetCellData()
        {
            var binExcel = File.ReadAllBytes(@"test.xlsx");
            var base64Excel = Convert.ToBase64String(binExcel);

            var template = @"{cell1:${Tabelle1!A1},cell2:'${Tabelle1!B2},${Tabelle2!C8}'}";

            var expectedResult = template
                .Replace("${Tabelle1!A1}", "4711")
                .Replace("${Tabelle1!B2}", "This is a text");

            using(var excelService = new ExcelService(base64Excel))
            {
                var actualResult = excelService.GetCellData(template);

                Assert.Equal(expectedResult, actualResult);
            }
        }

        [Fact]
        public void WriteCellData()
        {
            var binExcel = File.ReadAllBytes(@"test.xlsx");
            var base64Excel = Convert.ToBase64String(binExcel);

            var json = @"{
                ""Tabelle1!B3"":""This is a test too"",
                ""Tabelle2!A2"": 1234
            }";

            var data = JsonConvert.DeserializeObject<Dictionary<string, object>>(json);

            using(var excelService = new ExcelService(base64Excel))
            {
                var actualResult = excelService.WriteCellData(data);
                // not a very goot assert
                Assert.NotEmpty(actualResult);

                var tempPath = Path.GetTempFileName() + ".xlsx";
                File.WriteAllBytes(tempPath, Convert.FromBase64String(actualResult));

                // check manually
                Console.WriteLine(tempPath);
            }
        }
    }
}

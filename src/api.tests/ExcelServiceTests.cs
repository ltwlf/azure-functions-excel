using System;
using Xunit;
using Ltwlf.Functions.Excel;
using System.IO;

namespace Api.Tests
{
    public class ExcelServiceTests
    {
        [Fact]
        public void ReplaceTokens()
        {
            var binExcel = File.ReadAllBytes(@"test.xlsx");
            var base64Excel = Convert.ToBase64String(binExcel);

            var template = @"{cell1:${Tabelle1!A1},cell2:'${Tabelle1!B2},${Tabelle2!C8}'}";

            var expectedResult = template
                .Replace("${Tabelle1!A1}", "4711")
                .Replace("${Tabelle1!B2}", "This is a text");

            using(var excelService = new ExcelService(base64Excel))
            {
                var actualResult = excelService.ReplaceTokens(template);

                Assert.Equal(expectedResult, actualResult);
            }
        }
    }
}

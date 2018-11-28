using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;

namespace Ltwlf.Functions.Excel
{
    public static class ReplaceValues
    {
        [FunctionName("ReplaceCellValues")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)] HttpRequest req,
            ILogger log)
        {
            log.LogInformation("HttpTrigger ReplaceCellValues is processing a request...");

            var requestBody = await new StreamReader(req.Body).ReadToEndAsync();
            dynamic data = JsonConvert.DeserializeObject(requestBody);
            
            var excelAsBase64 = data.excelAsBase64.ToString();
            var template = data.template.ToString();

            var excelService = new ExcelService(excelAsBase64);
            var result = excelService.ReplaceTokens(template);

            return new OkObjectResult(new { result = result} );
        }
    }
}

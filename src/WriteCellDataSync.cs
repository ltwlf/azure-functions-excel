using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using System.Collections.Generic;

namespace Ltwlf.Functions.Excel
{
    [Serializable]
    public class WriteCellDataMessageInputSync
    {
        [JsonProperty("excelAsBase64")]
        public string ExcelAsBase64 { get; set; }

        [JsonProperty("mapping")]
        public Dictionary<string, object> Mapping { get; set; }
    }
    
    public static class WriteCellDataSync
    {
        [FunctionName("WriteCellDataSync")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Function, "post", Route = null)] HttpRequest req,
            ILogger log)
        {
            string data = await req.ReadAsStringAsync();

            var msg = JsonConvert.DeserializeObject<WriteCellDataMessageInputSync>(data);

            var excelService = new ExcelService(msg.ExcelAsBase64);

            var result = excelService.WriteCellData(msg.Mapping);

            return new OkObjectResult(new GetCellDataMessageOutput()
            {
                Result = result
            });
        }
    }
}

using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Collections.Generic;

namespace Ltwlf.Functions.Excel
{
    [Serializable]
    public class WriteCellDataMessageInput
    {
        [JsonProperty("excelAsBase64")]
        public string ExcelAsBase64 { get; set; }

        [JsonProperty("mapping")]
        public string Mapping { get; set; }
    }

    public class WriteCellDataMessageOutput
    {
        [JsonProperty("result")]
        public string Result { get; set; }
    }
    public static class WriteCellDataFunction
    {

        [FunctionName("WriteCellData")]
        public static async Task<HttpResponseMessage> HttpStart(
            [HttpTrigger(AuthorizationLevel.Function, "post")] HttpRequestMessage req,
            [OrchestrationClient] DurableOrchestrationClient starter,
            ILogger log)
        {
            var data = await req.Content.ReadAsStringAsync();

            var msg = JsonConvert.DeserializeObject<WriteCellDataMessageInput>(data);

            string instanceId = await starter.StartNewAsync("WriteCellDataOrchestrator", msg);

            log.LogInformation($"Started orchestration with ID = '{instanceId}'.");

            var res = starter.CreateCheckStatusResponse(req, instanceId);

            res.Headers.RetryAfter = new RetryConditionHeaderValue(TimeSpan.FromSeconds(5));

            return res;
        }

        [FunctionName("WriteCellDataOrchestrator")]
        public static async Task<WriteCellDataMessageOutput> WriteCellDataOrchestrator(
            [OrchestrationTrigger] DurableOrchestrationContext context)
        {
            var msg = context.GetInput<WriteCellDataMessageInput>();

            return await context.CallActivityAsync<WriteCellDataMessageOutput>("WriteCellDataWorker", msg);
        }

        [FunctionName("WriteCellDataWorker")]
        public static object WriteCellDataWorker([ActivityTrigger]WriteCellDataMessageInput msg, ILogger log)
        {
            var excelAsBase64 = msg.ExcelAsBase64;

            try
            {
                var excelService = new ExcelService(excelAsBase64);

                var mapping = JsonConvert.DeserializeObject<Dictionary<string, object>>(msg.Mapping);

                var result = excelService.WriteCellData(mapping);

                return new GetCellDataMessageOutput()
                {
                    Result = result
                };
            }
            catch (Exception ex)
            {
                log.LogError(ex, "Error while writting to Excel");
                throw;
            }
        }
    }
}

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
    public class GetCellDataMessageInput
    {
        [JsonProperty("excelAsBase64")]
        public string ExcelAsBase64 { get; set; }

        [JsonProperty("template")]
        public string Template { get; set; }
    }

    public class GetCellDataMessageOutput
    {
        [JsonProperty("result")]
        public string Result { get; set; }
    }
    public static class GetCellDataFunction
    {

        [FunctionName("GetCellData")]
        public static async Task<HttpResponseMessage> HttpStart(
            [HttpTrigger(AuthorizationLevel.Function, "post")] HttpRequestMessage req,
            [OrchestrationClient] DurableOrchestrationClient starter,
            ILogger log)
        {
            var data = await req.Content.ReadAsStringAsync();

            var msg = JsonConvert.DeserializeObject<GetCellDataMessageInput>(data);

            string instanceId = await starter.StartNewAsync("GetCellDataOrchestrator", msg);

            log.LogInformation($"Started orchestration with ID = '{instanceId}'.");

            var res = starter.CreateCheckStatusResponse(req, instanceId);

            res.Headers.RetryAfter = new RetryConditionHeaderValue(TimeSpan.FromSeconds(5));

            return res;
        }

        [FunctionName("GetCellDataOrchestrator")]
        public static async Task<GetCellDataMessageOutput> GetCellDataOrchestrator(
            [OrchestrationTrigger] DurableOrchestrationContext context)
        {
            var msg = context.GetInput<GetCellDataMessageInput>();

            return await context.CallActivityAsync<GetCellDataMessageOutput>("GetCellDataWorker", msg);
        }

        [FunctionName("GetCellDataWorker")]
        public static object GetCellDataWorker([ActivityTrigger]GetCellDataMessageInput msg, ILogger log)
        {
            var excelAsBase64 = msg.ExcelAsBase64;
            var template = msg.Template;

            try
            {
                var excelService = new ExcelService(excelAsBase64);
                var result = excelService.GetCellData(template);

                return new GetCellDataMessageOutput()
                {
                    Result = result
                };
            }
            catch (Exception ex)
            {
                log.LogError(ex, "Error while reading to Excel");
                throw;
            }
        }
    }
}

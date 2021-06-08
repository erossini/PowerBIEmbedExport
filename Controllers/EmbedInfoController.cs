// ----------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
// ----------------------------------------------------------------------------

namespace AppOwnsData.Controllers
{
    using AppOwnsData.Models;
    using AppOwnsData.Services;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Extensions.Options;
	using Microsoft.PowerBI.Api;
	using Microsoft.PowerBI.Api.Models;
	using Microsoft.Rest;
	using System;
	using System.Collections.Generic;
	using System.IO;
	using System.Linq;
	using System.Text.Json;
	using System.Threading;
	using System.Threading.Tasks;

	public class EmbedInfoController : Controller
    {
        private readonly PbiEmbedService pbiEmbedService;
        private readonly IOptions<AzureAd> azureAd;
        private readonly IOptions<PowerBI> powerBI;
        private PowerBIClient Client;

        public EmbedInfoController(PbiEmbedService pbiEmbedService, IOptions<AzureAd> azureAd, IOptions<PowerBI> powerBI)
        {
            this.pbiEmbedService = pbiEmbedService;
            this.azureAd = azureAd;
            this.powerBI = powerBI;

            Client = pbiEmbedService.GetPowerBIClient();
        }

        /// <summary>
        /// Returns Embed token, Embed URL, and Embed token expiry to the client
        /// </summary>
        /// <returns>JSON containing parameters for embedding</returns>
        [HttpGet]
        public string GetEmbedInfo()
        {
            try
            {
                // Validate whether all the required configurations are provided in appsettings.json
                string configValidationResult = ConfigValidatorService.ValidateConfig(azureAd, powerBI);
                if (configValidationResult != null)
                {
                    HttpContext.Response.StatusCode = 400;
                    return configValidationResult;
                }

                EmbedParams embedParams = pbiEmbedService.GetEmbedParams(new Guid(powerBI.Value.WorkspaceId), new Guid(powerBI.Value.ReportId));
                return JsonSerializer.Serialize<EmbedParams>(embedParams);
            }
            catch (Exception ex)
            {
                HttpContext.Response.StatusCode = 500;
                return ex.Message + "\n\n" + ex.StackTrace;
            }
        }

        #region Export

        public async Task<IActionResult> GetReport()
		{
            Pages pages = pbiEmbedService.GetPages(new Guid(powerBI.Value.WorkspaceId), new Guid(powerBI.Value.ReportId));
            List<string> pgs = pages.Value.Select(p => p.Name).ToList();
            CancellationToken tk = new CancellationToken();
            ExportedFile file = await ExportPowerBIReport(new Guid(powerBI.Value.ReportId), new Guid(powerBI.Value.WorkspaceId), FileFormat.PDF, 1, tk, pgs);

            string mimeType = "application/pdf";
            return new FileStreamResult(file.FileStream, mimeType)
            {
                FileDownloadName = "Test.pdf"
            };
		}

        private async Task<string> PostExportRequest(
            Guid reportId,
            Guid groupId,
            FileFormat format,
            IList<string> pageNames = null, /* Get the page names from the GetPages REST API */
            string urlFilter = null)
        {
            var powerBIReportExportConfiguration = new PowerBIReportExportConfiguration
            {
                Settings = new ExportReportSettings
                {
                    Locale = "en-us",
                },
                // Note that page names differ from the page display names
                // To get the page names use the GetPages REST API
                Pages = pageNames?.Select(pn => new ExportReportPage(pageName: pn)).ToList(),
                // ReportLevelFilters collection needs to be instantiated explicitly
                ReportLevelFilters = !string.IsNullOrEmpty(urlFilter) ? new List<ExportFilter>() { new ExportFilter(urlFilter) } : null,

            };

            var exportRequest = new ExportReportRequest
            {
                Format = format,
                PowerBIReportConfiguration = powerBIReportExportConfiguration,
            };

            // The 'Client' object is an instance of the Power BI .NET SDK
            var export = await Client.Reports.ExportToFileInGroupAsync(groupId, reportId, exportRequest);

            // Save the export ID, you'll need it for polling and getting the exported file
            return export.Id;
        }

        private async Task<HttpOperationResponse<Export>> PollExportRequest(
            Guid reportId,
            Guid groupId,
            string exportId /* Get from the PostExportRequest response */,
            int timeOutInMinutes,
            CancellationToken token)
        {
            HttpOperationResponse<Export> httpMessage = null;
            Export exportStatus = null;
            DateTime startTime = DateTime.UtcNow;
            const int c_secToMillisec = 1000;
            do
            {
                if (DateTime.UtcNow.Subtract(startTime).TotalMinutes > timeOutInMinutes || token.IsCancellationRequested)
                {
                    // Error handling for timeout and cancellations 
                    return null;
                }

                // The 'Client' object is an instance of the Power BI .NET SDK
                httpMessage = await Client.Reports.GetExportToFileStatusInGroupWithHttpMessagesAsync(groupId, reportId, exportId);
                exportStatus = httpMessage.Body;

                // You can track the export progress using the PercentComplete that's part of the response
                Console.WriteLine(string.Format("{0} (Percent Complete : {1}%)", exportStatus.Status.ToString(), exportStatus.PercentComplete));
                if (exportStatus.Status == ExportState.Running || exportStatus.Status == ExportState.NotStarted)
                {
                    // The recommended waiting time between polling requests can be found in the RetryAfter header
                    // Note that this header is not always populated
                    var retryAfter = httpMessage.Response.Headers.RetryAfter;
                    var retryAfterInSec = retryAfter.Delta.Value.Seconds;
                    await Task.Delay(retryAfterInSec * c_secToMillisec);
                }
            }
            // While not in a terminal state, keep polling
            while (exportStatus.Status != ExportState.Succeeded && exportStatus.Status != ExportState.Failed);

            return httpMessage;
        }

        private async Task<ExportedFile> GetExportedFile(
            Guid reportId,
            Guid groupId,
            Export export /* Get from the PollExportRequest response */)
        {
            if (export.Status == ExportState.Succeeded)
            {
                // The 'Client' object is an instance of the Power BI .NET SDK
                var fileStream = await Client.Reports.GetFileOfExportToFileAsync(groupId, reportId, export.Id);
                return new ExportedFile
                {
                    FileStream = fileStream,
                    FileSuffix = export.ResourceFileExtension,
                };
            }
            return null;
        }

        public class ExportedFile
        {
            public Stream FileStream;
            public string FileSuffix;
        }

        private async Task<ExportedFile> ExportPowerBIReport(
            Guid reportId,
            Guid groupId,
            FileFormat format,
            int pollingtimeOutInMinutes,
            CancellationToken token,
            IList<string> pageNames = null,  /* Get the page names from the GetPages REST API */
            string urlFilter = null)
        {
            const int c_maxNumberOfRetries = 3; /* Can be set to any desired number */
            const int c_secToMillisec = 1000;
            try
            {
                Export export = null;
                int retryAttempt = 1;
                do
                {
                    var exportId = await PostExportRequest(reportId, groupId, format, pageNames, urlFilter);
                    var httpMessage = await PollExportRequest(reportId, groupId, exportId, pollingtimeOutInMinutes, token);
                    export = httpMessage.Body;
                    if (export == null)
                    {
                        // Error, failure in exporting the report
                        return null;
                    }
                    if (export.Status == ExportState.Failed)
                    {
                        // Some failure cases indicate that the system is currently busy. The entire export operation can be retried after a certain delay
                        // In such cases the recommended waiting time before retrying the entire export operation can be found in the RetryAfter header
                        var retryAfter = httpMessage.Response.Headers.RetryAfter;
                        if (retryAfter == null)
                        {
                            // Failed state with no RetryAfter header indicates that the export failed permanently
                            return null;
                        }

                        var retryAfterInSec = retryAfter.Delta.Value.Seconds;
                        await Task.Delay(retryAfterInSec * c_secToMillisec);
                    }
                }
                while (export.Status != ExportState.Succeeded && retryAttempt++ < c_maxNumberOfRetries);

                if (export.Status != ExportState.Succeeded)
                {
                    // Error, failure in exporting the report
                    return null;
                }

                var exportedFile = await GetExportedFile(reportId, groupId, export);

                // Now you have the exported file stream ready to be used according to your specific needs
                // For example, saving the file can be done as follows:
                /*
                    var pathOnDisk = @"C:\temp\" + export.ReportName + exportedFile.FileSuffix;

                    using (var fileStream = File.Create(pathOnDisk))
                    {
                        exportedFile.FileStream.CopyTo(fileStream);
                    }
                */

                return exportedFile;
            }
            catch
            {
                // Error handling
                throw;
            }
        }

        #endregion
    }
}

using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Extensions.Logging;
using Microsoft.SharePoint.Client;
using PnP.Core.Admin.Model.SharePoint;
using PnP.Core.Model.SharePoint;
using PnP.Core.Services;
using PnP.Framework;

//using PnP.Core.Services;
//using Microsoft.SharePoint.Client;
using PnP.Framework.Provisioning.Connectors;
using PnP.Framework.Provisioning.Model;
using PnP.Framework.Provisioning.ObjectHandlers;
using PnP.Framework.Provisioning.Providers.Xml;



using System;
using System.Collections.Specialized;
using System.IO;
using System.Net;
using System.Text.Json;
using System.Threading.Tasks;
using System.Web;

namespace ProvisioningDemo
{
    public class ExtractTemplate
    {
        private readonly ILogger logger;
        private readonly IPnPContextFactory contextFactory;
        private readonly AzureFunctionSettings azureFunctionSettings;

        public ExtractTemplate(IPnPContextFactory pnpContextFactory, ILoggerFactory loggerFactory, AzureFunctionSettings settings)
        {
            logger = loggerFactory.CreateLogger<ExtractTemplate>();
            contextFactory = pnpContextFactory;
            azureFunctionSettings = settings;
        }

        /// <summary>
        /// Demo function that creates a site collection, uploads an image to site assets and creates a page with an image web part
        /// GET/POST url: http://localhost:7071/api/ExtractTemplate?owner=bert.jansen@bertonline.onmicrosoft.com&sitename=deleteme1844
        /// </summary>
        /// <param name="req"></param>
        /// <returns></returns>
        [Function("ExtractTemplate")]
        public async Task<HttpResponseData> Run([HttpTrigger(AuthorizationLevel.Function, "get", "post")] HttpRequestData req)
        {
            logger.LogInformation("ExtractTemplate function starting...");

            // Parse the url parameters
            NameValueCollection parameters = HttpUtility.ParseQueryString(req.Url.Query);
            var siteName = parameters["siteName"];
            var siteUrl = parameters["siteUrl"];
            var owner = parameters["owner"];

            HttpResponseData response = null;

            try
            {

                using (var pnpContext = await contextFactory.CreateAsync(new Uri(siteUrl)))
                {
                    response = req.CreateResponse(HttpStatusCode.OK);
                    response.Headers.Add("Content-Type", "application/json");

                    // Load the root folder to get all the properties
                    await pnpContext.Web.LoadAsync(p => p.Title);

                    logger.LogInformation($"The source site to create the tempalte is: {pnpContext.Web.Title}");

                    using (ClientContext ctx = PnPCoreSdk.Instance.GetClientContext(pnpContext))
                    {

                        

                        ProvisioningTemplateCreationInformation ptci = new ProvisioningTemplateCreationInformation(ctx.Web);
                        // Create FileSystemConnector to store a temporary copy of the template
                        ptci.FileConnector = new FileSystemConnector(@"c:\temp\PratusProvisioningDemo", "");
                        ptci.PersistBrandingFiles = true;
                        ptci.ProgressDelegate = delegate (String message, Int32 progress, Int32 total)
                        {
                            // Only to output progress for console UI
                            Console.WriteLine("{0:00}/{1:00} - {2}", progress, total, message);
                        };
                        // Extract the template
                        ProvisioningTemplate template = ctx.Web.GetProvisioningTemplate(ptci);
                        // We can serialize this template to save and reuse it
                        XMLTemplateProvider provider = new XMLFileSystemTemplateProvider(@"c:\temp\PratusProvisioningDemo", "");
                        provider.SaveAs(template, "SourceSiteTemplate.xml");
                    }




                    // Return the URL of the templated site
                    await response.WriteStringAsync(JsonSerializer.Serialize(new { siteUrl = pnpContext.Uri.AbsoluteUri }));                 

                    return response;
                }
            }
            catch (Exception ex)
            {
                response = req.CreateResponse(HttpStatusCode.OK);
                response.Headers.Add("Content-Type", "application/json");
                await response.WriteStringAsync(JsonSerializer.Serialize(new { error = ex.Message }));
                return response;
            }
        }
    }
}

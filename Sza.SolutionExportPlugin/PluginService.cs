using McTools.Xrm.Connection;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.IO;
using System.Threading.Tasks;
using System.Web.Services.Protocols;
using Sza.SolutionExportPlugin.Shared.Enum;
using XrmToolBox.Extensibility;

namespace Sza.SolutionExportPlugin
{
    public class PluginService
    {
        public static EntityCollection GetUnmanagedSolutions(IOrganizationService service)
        {

            var query = new QueryExpression("solution");
            var filter = new FilterExpression();
            filter.AddCondition(new ConditionExpression("ismanaged", ConditionOperator.Equal, false));
            filter.AddCondition(new ConditionExpression("uniquename", ConditionOperator.NotIn, Constants.friendlyName));
            query.Criteria.AddFilter(filter);
            query.ColumnSet.AddColumns("uniquename", "version", "friendlyname");
            var solutions = service.RetrieveMultiple(query);
            return solutions;
        }
        public static string ExportSolution(IOrganizationService service, ConnectionDetail ConnectionDetail, ExportSolutionRequest exportSolutionRequest, string version, string outputDir)
        {
            outputDir = outputDir + @"\";
            try
            {

                if (ConnectionDetail != null)
                    ConnectionDetail.Timeout = new TimeSpan(1, 0, 0);
                //if (client.OrganizationWebProxyClient != null)
                //    ConnectionDetail.OperationTimeout = new TimeSpan(1, 0, 0);

               // ExportSolutionResponse exportSolutionResponse = await Task.Run(() => (ExportSolutionResponse)service.Execute(exportSolutionRequest));
                ExportSolutionResponse exportSolutionResponse =  (ExportSolutionResponse)service.Execute(exportSolutionRequest);
                byte[] exportXml = exportSolutionResponse.ExportSolutionFile;
                string filename;
                if (exportSolutionRequest.Managed)
                {
                    filename = exportSolutionRequest.SolutionName + "_" + version + "_managed.zip";
                }
                else
                {
                    filename = exportSolutionRequest.SolutionName + "_" + version + ".zip";
                }

                File.WriteAllBytes(outputDir + filename, exportXml);
                return "Successfully exported";
            }
            catch (SoapException es)
            {
                return es.Message;
            }
            catch (Exception e)
            {
                return e.Message;
            }

        }
        public static string RetrieveVersionRequest(IOrganizationService service)
        {
            var request = new RetrieveVersionRequest();
            var response = (RetrieveVersionResponse)service.Execute(request);
            return response.Version;
        }
        //internal int Version(string number)
        //{

        //    //return null;
        //}
    }
}

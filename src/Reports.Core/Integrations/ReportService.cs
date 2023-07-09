using Integrations.Reports.Core.Dto;
using Integrations.Reports.Core.Integrations.Report;
using Integrations.Reports.Core.Integrations.Report.Dto;
using Reports.Core.Integrations.Report;

namespace Reports.Core.Integrations
{
    public class ReportService : IReportService
    {
        public async Task<EmployeeReportDto> GenerateEmployeeReport(List<EmployeeDto> employees, ReportTypes type, CancellationToken token)
        {

            var outputPath = "path";//$"{AppSettings.ReportDirectory.TrimEndPath()}\\BackOfficeReport\\";
            var report = new EmployeeReport(new ReportingServiceEmployeeReportDto
            {
                Data = employees,
                ApplicationName = "IntegrationReports",
                OutputPath = outputPath
            });

            var reportResult = report.GenerateReport();
            return new EmployeeReportDto
            {
                Report = reportResult
            };
        }
    }
}

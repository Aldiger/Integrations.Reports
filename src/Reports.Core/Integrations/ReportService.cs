using Integrations.Reports.Core.Dto;
using Integrations.Reports.Core.Helpers;
using Integrations.Reports.Core.Integrations.Report;
using Integrations.Reports.Core.Integrations.Report.Dto;
using Microsoft.Extensions.Options;
using Reports.Core.Options;

namespace Reports.Core.Integrations
{
    public class ReportService : IReportService
    {
        private readonly ReportsOptions _reportOptions;
        public ReportService(IOptions<ReportsOptions> reportOptions)
        {
            _reportOptions = reportOptions.Value;
        }
        public async Task<EmployeeReportDto> GenerateEmployeeReport(List<EmployeeDto> employees, ReportTypes type, CancellationToken token)
        {

            var outputPath = _reportOptions.Path;
            var report = new EmployeeReport(new ReportingServiceEmployeeReportDto
            {
                Data = employees,
                ApplicationName = "IntegrationsReports",
                OutputPath = outputPath,
                OutputType = type
            });

            var reportResult = report.GenerateReport();
            return new EmployeeReportDto
            {
                Report = reportResult
            };

        }
    }
}

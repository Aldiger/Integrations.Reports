using Integrations.Reports.Core.Dto;
using Integrations.Reports.Core.Helpers;

namespace Integrations.Reports.Core.Integrations.Report.Dto
{
    public class ReportingServiceEmployeeReportDto
    {
        public string OutputPath { get; set; }
        public string ApplicationName { get; set; }
        public ReportTypes OutputType { get; set; }
        public IList<EmployeeDto> Data { get; set; }
    }
}

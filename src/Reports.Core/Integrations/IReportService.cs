using Integrations.Reports.Core.Dto;
using Reports.Core.Integrations.Report;

namespace Reports.Core.Integrations
{
    public interface IReportService
    {
        Task<EmployeeReportDto> GenerateEmployeeReport(List<EmployeeDto>
             employees, ReportTypes type, CancellationToken token);
    }
}

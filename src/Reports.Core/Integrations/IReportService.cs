using Integrations.Reports.Core.Dto;
using Integrations.Reports.Core.Helpers;

namespace Reports.Core.Integrations
{
    public interface IReportService
    {
        Task<EmployeeReportDto> GenerateEmployeeReport(List<EmployeeDto>
             employees, ReportTypes type, CancellationToken token);
    }
}

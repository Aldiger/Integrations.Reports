using Integrations.Reports.Core.Dto;
using MediatR;
using Reports.Core.Integrations.Report;

namespace Integrations.Reports.Core.Queries
{
    public class GetEmployeeReport : IRequest<EmployeeReportDto>
    {
        public List<EmployeeDto> Employees { get; set; }
        public ReportTypes Type { get; set; }
    }
}

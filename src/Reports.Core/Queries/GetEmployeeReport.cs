using Integrations.Reports.Core.Dto;
using Integrations.Reports.Core.Helpers;
using MediatR;

namespace Integrations.Reports.Core.Queries
{
    public class GetEmployeeReport : IRequest<EmployeeReportDto>
    {
        public List<EmployeeDto> Employees { get; set; }
        public ReportTypes Type { get; set; }
    }
}

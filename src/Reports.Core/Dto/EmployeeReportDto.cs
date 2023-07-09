using Integrations.Reports.Core.Integrations.Report.Dto;
using MediatR;

namespace Integrations.Reports.Core.Dto
{
    public class EmployeeReportDto: IRequest<EmployeeReportDto>
    {
        public ReportResultDto Report { get; set; }
    }
}

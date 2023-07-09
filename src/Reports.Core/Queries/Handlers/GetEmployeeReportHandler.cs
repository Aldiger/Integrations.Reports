using Integrations.Reports.Core.Dto;
using MediatR;
using Reports.Core.Integrations;

namespace Integrations.Reports.Core.Queries.Handlers
{
    public class GetEmployeeReportHandler : IRequestHandler<GetEmployeeReport, EmployeeReportDto>
    {
        private readonly IReportService _reportService;
        public GetEmployeeReportHandler(IReportService reportService)
        {
            _reportService = reportService;
        }
        public async Task<EmployeeReportDto> Handle(GetEmployeeReport request, CancellationToken token)
        {
            //validate request

            var result = await _reportService.GenerateEmployeeReport(request.Employees, request.Type, token);

            return result;
        }
    }
}

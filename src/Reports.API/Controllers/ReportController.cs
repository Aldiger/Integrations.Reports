using Integrations.Reports.Core.Queries;
using MediatR;
using Microsoft.AspNetCore.Mvc;

namespace Reports.API.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class ReportController : ControllerBase
    {
        private readonly IMediator _mediator;
        public ReportController(IMediator mediator)
        {
            _mediator = mediator;
        }

        [HttpPost("employee")]
        public async Task<IActionResult> EmployeeReport(GetEmployeeReport request, CancellationToken token)
        {
            var result = _mediator.Send(request, token);
            return Ok(result);
        }
    }
}
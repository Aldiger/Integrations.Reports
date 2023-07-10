using Integrations.Reports.Core.Queries;
using MediatR;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.StaticFiles;

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
            var result = await _mediator.Send(request, token);
            new FileExtensionContentTypeProvider().TryGetContentType(result.Report.FileName, out var contentType);
            return File(await System.IO.File.ReadAllBytesAsync(result.Report.FilePath), contentType, result.Report.FileName);
        }
    }
}
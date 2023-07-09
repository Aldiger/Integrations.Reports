using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Options;
using Reports.Core.Options;

namespace Reports.API.Controllers
{
    [Route("")]
    public class HomeController : ControllerBase
    {
        private readonly IOptions<AppOptions> _appOptions;
        public HomeController(IOptions<AppOptions> appOptions)
        {
            _appOptions = appOptions;
        }

        [HttpGet]
        public IActionResult Get() => Ok(_appOptions.Value.Name);
    }
}
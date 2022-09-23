using Microsoft.Azure.Functions.Extensions.DependencyInjection;
using Microsoft.Extensions.DependencyInjection;

[assembly: FunctionsStartup(typeof(CalendarSync.Startup))]
namespace CalendarSync
{
    public class Startup : FunctionsStartup
    {
        public override void Configure(IFunctionsHostBuilder builder)
        {
            builder.Services.AddSingleton<IGraphClient, GraphClient>();
            builder.Services.AddSingleton<ITableService, TableService>();
        }
    }
}

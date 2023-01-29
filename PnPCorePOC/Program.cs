using Microsoft.Extensions.DependencyInjection;
using PnP.Core.QueryModel;
using PnP.Core.Services;
using PnPCorePOC.Helper;
using System;
using System.Threading.Tasks;

namespace PnPCorePOC
{
    internal class Program
    {
        private static async Task Main(string[] args)
        {
            ///As value set target site from AppSecret.config by it SiteName key:
            var siteContextName = "pnpdemo";

            var host = SPOAuthorizations.GetAuthorization(siteContextName);
            await host.StartAsync();
            using (var scope = host.Services.CreateScope())
            {
                var pnpContextFactory = scope.ServiceProvider.GetRequiredService<IPnPContextFactory>();
                using (var context = await pnpContextFactory.CreateAsync(siteContextName))
                {
                    var myWeb = await context.Web.GetAsync(p => p.Title);
                    Console.WriteLine(myWeb.Title);

                    await context.Web.LoadAsync(p => p.Lists.QueryProperties(
                                                        l => l.Id,
                                                        l => l.Hidden,
                                                        l => l.Title));

                    foreach (var list in context.Web.Lists.AsRequested().Where(p => p.Hidden == false))
                    {
                        Console.WriteLine($"Lista: {list.Id} == {list.Title}");
                    }

                    await context.Web.LoadAsync(p => p.RoleDefinitions);
                    foreach (var role in context.Web.RoleDefinitions.AsRequested())
                    {
                        Console.WriteLine($"Rola: {role.Id} == {role.Name}");
                    }
                }
                Console.ReadLine();
            }
        }
    }
}
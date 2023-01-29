using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using PnP.Core.Auth.Services.Builder.Configuration;
using PnP.Core.Services;
using PnP.Core.Services.Builder.Configuration;
using System;
using System.Security.Cryptography.X509Certificates;

namespace PnPCorePOC.Helper
{
    public static class SPOAuthorizations
    {
        public static IHost GetAuthorization(string hostName)
        {
            try
            {
                var configuration = new GetSPOConfiguration();
                if (configuration.Configuration.ContainsKey(hostName))
                {
                    var host = Host.CreateDefaultBuilder()
                        .ConfigureServices((context, services) =>
                        {
                            services.AddLogging(
                            builder =>
                            {
                                builder.AddFilter("Microsoft", LogLevel.Warning)
                                       .AddFilter("System", LogLevel.Warning)
                                       .AddFilter("PnP.Core.Auth", LogLevel.Warning)
                                       .AddConsole();
                            });
                            services.AddPnPCore(options =>
                            {
                                options.PnPContext.GraphFirst = true;
                                options.HttpRequests.UserAgent = $"ISV|{hostName}|Product";
                                options.Sites.Add(hostName, new PnPCoreSiteOptions
                                {
                                    SiteUrl = configuration.Configuration[hostName].SiteURL.ToString(),
                                });
                            });
                            services.AddPnPCoreAuthentication(
                                options =>
                                {
                                    options.Credentials.Configurations.Add("x509certificate", new PnPCoreAuthenticationCredentialConfigurationOptions
                                    {
                                        ClientId = configuration.Configuration[hostName].ApplicationID,
                                        TenantId = configuration.Configuration[hostName].TenantID,
                                        X509Certificate = new PnPCoreAuthenticationX509CertificateOptions
                                        {
                                            StoreName = StoreName.My,
                                            StoreLocation = StoreLocation.LocalMachine,
                                            Thumbprint = configuration.Configuration[hostName].CertificateThumbprint
                                        }
                                    });
                                    options.Credentials.DefaultConfiguration = "x509certificate";
                                    options.Sites.Add(hostName, new PnPCoreAuthenticationSiteOptions
                                    {
                                        AuthenticationProviderName = "x509certificate"
                                    });
                                }
                            );
                        })
                        .UseConsoleLifetime()
                        .Build();
                    return host;
                }
                else
                {
                    var cfc = Console.ForegroundColor;
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine($"Settings for {hostName} not found.");
                    Console.ForegroundColor = cfc;
                    return null;
                }
            }
            catch (Exception ex)
            {
                var cfc = Console.ForegroundColor;
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"GetAuthorization exception {ex.Message}");
                Console.ForegroundColor = cfc;
                return null;
            }
        }

        public static GraphServiceClient CreateGraphClient(PnPContext context)
        {
            try
            {
                return new GraphServiceClient(new DelegateAuthenticationProvider((requestMessage) =>
                {
                    return context.AuthenticationProvider.AuthenticateRequestAsync(new Uri("https://graph.microsoft.com"), requestMessage);
                }));
            }
            catch (Exception ex)
            {
                var cfc = Console.ForegroundColor;
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"GraphServiceClient exception {ex.Message}");
                Console.ForegroundColor = cfc;
                return null;
            }
        }
    }
}
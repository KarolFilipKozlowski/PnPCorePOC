# About Me:

🪪 See my blog: [Karol Kozłowski I CitDev](https://citdev.pl/blog/)

# PnPCorePOC by Karol Kozłowski
It's my repository helper for [Microsoft 365 & Power Platform Community](https://pnp.github.io/).

🔗PnP Core Documentation: [# PnP Core SDK](https://pnp.github.io/pnpcore/)

# Start Here:

## I. Create new solution:

Add new project (.NET Framework) with .NET Framework 4.8
From NuGet add:
- **Microsoft.Extensions.Hosting** package
- **PnP.Core** package
- **PnP.Core.Auth** package
- **Microsoft.Graph** package
- [OPTIONAL] **GitInfo** package

Update all packages.

### 0) [OPTIONAL] GIT:
Add project to a code repository.

### 1a) Add _AppSecret.config_:
Select project, from context menu select: Add -> New item -> Application Configuration File, add **AppSecret.config**.
Replace code with:
````
<appSettings>
	<add key="SPO:0:TenantID" value="8ae35f9e-b3c6-486b-bdf8-5c8da0cff7b9"/>
	<add key="SPO:0:SiteURL" value="https://contoso.sharepoint.com/sites/pnpdemo"/>
	<add key="SPO:0:ApplicationID" value="64f4c29a-b1d5-47c5-91d3-ecdfaedeb72a"/>
	<add key="SPO:0:CertificateThumbprint" value="4AA402CDC596471696C6159254DEF6B30ABBB44D"/>
</appSettings>
````
**In git add this file for ignore!**

Select **AppSecret.config**, in file properties -> copy to Output Directory set **Copy always**.
 
### 2b) Edit _App.config_:
In end of **<configuration>** add:
````
<appSettings file="AppSecret.config">
</appSettings>
````
After **<assemblyBinding xmlns="urn:schemas-microsoft-com:asm.v1">** add:
````
<probing privatePath="Bin"/>
````
### 3) [OPTIONAL] Edit _AssemblyInfo.cs_, add on end:
Edit `AssemblyCopyrightt`:
 ````
[assembly: AssemblyProduct("https://github.com/KarolFilipKozlowski/PnPCorePOC | " + ThisAssembly.Git.Branch)]
[assembly: AssemblyCopyright("CitDeV ©  2023")]
````
On end change/add:
````
[assembly: AssemblyVersion(ThisAssembly.Git.BaseVersion.Major + "." + ThisAssembly.Git.BaseVersion.Minor + "." + ThisAssembly.Git.BaseVersion.Patch)]

[assembly: AssemblyFileVersion(ThisAssembly.Git.SemVer.Major + "." + ThisAssembly.Git.SemVer.Minor + "." + ThisAssembly.Git.SemVer.Patch)]

[assembly: AssemblyInformationalVersion(
    ThisAssembly.Git.SemVer.Major + "." +
    ThisAssembly.Git.SemVer.Minor + "." +
    ThisAssembly.Git.Commits + "-" +
    ThisAssembly.Git.Commit)]
````

### 4) [OPTIONAL] Go to _properties_ of app:
- Select **Application**. In Icon and manifest add icon: **PnPCoreApp_254.ico**.
- Select **Build**. In platform target select: **x64**.
- Select **Build Events**. In pre-build event add.
````
rmdir $(SolutionDir)APP /s /q

mkdir $(SolutionDir)APP
mkdir $(SolutionDir)APP\Bin
mkdir $(SolutionDir)APP\Logs

copy $(TargetDir)*.exe  $(SolutionDir)APP
copy $(TargetDir)*.config  $(SolutionDir)APP
copy $(TargetDir)*.dll  $(SolutionDir)APP\Bin

powershell Compress-Archive -Path '$(SolutionDir)APP' -DestinationPath '$(SolutionDir)APP\$(ProjectName).zip' -Force
````

**Add folder _APP_ as git ignore!**

## II. Create new classes:

Select project, from context menu select: Add -> New folder, set name **Helper**.\
Select folder, from context menu select: Add -> New class, set name **SPOConfiguraion.cs**.\
Again select folder, from context menu select: Add -> New class, set name **SPOAuthorizations.cs**.

### 1) Edit _SPOConfiguraion.cs_:

Replace flowing class:
````
class SPOConfiguration
{
}
````

With:
````
public class GetSPOConfiguration
{
    public Dictionary<string, ConfigurationProperties> Configuration { get; set; }
    public bool IsCorrect { get; private set; }

    public GetSPOConfiguration()
    {
        Configuration = GetConfiguration();
        if (Configuration != null)
        {
            IsCorrect = true;
        }
    }

    private Dictionary<string, ConfigurationProperties> GetConfiguration()
    {
        Dictionary<string, ConfigurationProperties> _configuration = new Dictionary<string, ConfigurationProperties>();

        try
        {
            foreach (var site in ConfigurationManager.AppSettings.AllKeys.Where(w => new Regex(@"^SPO:[0-9]*:SiteName").IsMatch(w)).ToList())
            {
                string iSite = site.ToString().Split(':')[1];
                string keySite = ConfigurationManager.AppSettings[$"SPO:{iSite}:SiteName"];
                _configuration.Add(keySite, new ConfigurationProperties());
                _configuration[keySite].SiteName = keySite;

                int i = 0;
                foreach (var key in ConfigurationManager.AppSettings.AllKeys.Where(w => w.StartsWith($"SPO:{iSite}:")))
                {
                    //Get TenantID:
                    Regex rgTenantID = new Regex(@"^SPO:[0-9]*:TenantID");
                    if (rgTenantID.IsMatch(key))
                    {
                        _configuration[keySite].TenantID = ConfigurationManager.AppSettings[$"SPO:{iSite}:TenantID"];
                        i++;
                    }
                    //Get SiteURL:
                    Regex rgSiteURL = new Regex(@"^SPO:[0-9]*:SiteURL");
                    if (rgSiteURL.IsMatch(key))
                    {
                        try
                        {
                            _configuration[keySite].SiteURL = new Uri(ConfigurationManager.AppSettings[$"SPO:{iSite}:SiteURL"]);
                            i++;
                        }
                        catch (Exception)
                        {
                            var cfc = Console.ForegroundColor;
                            Console.ForegroundColor = ConsoleColor.Yellow;
                            Console.WriteLine($"Invalid URL for SiteName: {keySite}.");
                            Console.ForegroundColor = cfc;
                        }

                    }
                    //Get ApplicationID:
                    Regex rgApplicationID = new Regex(@"^SPO:[0-9]*:ApplicationID");
                    if (rgApplicationID.IsMatch(key))
                    {
                        _configuration[keySite].ApplicationID = ConfigurationManager.AppSettings[$"SPO:{iSite}:ApplicationID"];
                        i++;
                    }
                    //Get CertificateThumbprint:
                    Regex rgCertificateThumbprint = new Regex(@"^SPO:[0-9]*:CertificateThumbprint");
                    if (rgCertificateThumbprint.IsMatch(key))
                    {
                        _configuration[keySite].CertificateThumbprint = ConfigurationManager.AppSettings[$"SPO:{iSite}:CertificateThumbprint"];
                        i++;
                    }
                }
                if (i == 4)
                {
                    _configuration[keySite].IsCorrect = true;
                }
                else
                {
                    var cfc = Console.ForegroundColor;
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine($"Configuration for (SiteName) {keySite} is missing some of properties!");
                    Console.ForegroundColor = cfc;
                }
            }
        }
        catch (Exception ex)
        {
            var cfc = Console.ForegroundColor;
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"SPOAppSecretConfiguration parser exception {ex.Message}");
            Console.ForegroundColor = cfc;
        }

        return _configuration;
    }
}

public class ConfigurationProperties
{
    public bool IsCorrect { get; set; }
    public string SiteName { get; set; }
    public string TenantID { get; set; }
    public Uri SiteURL { get; set; }
    public string ApplicationID { get; set; }
    public string CertificateThumbprint { get; set; }
}
````

Add messing using (above namespace):
````
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text.RegularExpressions;
````

### 1) Edit SPOAuthorizations.cs_:

Replace flowing class:
````
class SPOAuthorizations
{
}
````

With:
````
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
````

Add messing using (above namespace):
````
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Graph;
using PnP.Core.Auth.Services.Builder.Configuration;
using PnP.Core.Services;
using PnP.Core.Services.Builder.Configuration;
using System;
using System.Security.Cryptography.X509Certificates;
````

## II. Main program:

Open class **Program.cs**, change Main void:
````
static void Main(string[] args)
````
To:
````
static async Task Main(string[] args)
````

Add using:
````
using PnPCorePOC.Helper;
````

Add example code:
````
///As value set target site from AppSecret.config by it SiteName key:
var siteContextName = "Contso";

var host = SPOAuthorizations.GetAuthorization(siteContextName);
await host.StartAsync();
using (var scope = host.Services.CreateScope())
{
    var pnpContextFactory = scope.ServiceProvider.GetRequiredService<IPnPContextFactory>();
    using (var context = await pnpContextFactory.CreateAsync(siteContextName))
    {
        var myWeb = await context.Web.GetAsync(p => p.Title);
        Console.WriteLine(myWeb.Title);
    }
}
Console.ReadLine();
````

Run application (F5), you should see in console name of SharePoint Online site.

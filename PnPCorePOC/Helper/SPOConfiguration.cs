using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text.RegularExpressions;

namespace PnPCorePOC.Helper
{
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
}
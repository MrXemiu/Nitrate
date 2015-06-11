using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Web.Administration;
using Nitrate.Plugins;

namespace Nitrate.Plugins.Iis
{

  public class IisConfig
  {
    public string ClrVersion { get; set; }
    public string Site { get; set; }
    public int Port { get; set; }
    public string Path { get; set; }
    public string Name { get; set; }

    public IisConfig()
    {
      Site = "Default Web Site";
      Port = 8080;
      Name = "Orchard";
      Path = @"orchard\src\Orchard.Web";
    }
  }

  [Export(typeof(IPlugin))]
  public class Iis : BasePlugin<IisConfig>
  {
    public override string InstallationInstructions
    {
      get { return "Please install IIS by searching the Start menu for \"Turn Windows features on or off\"."; }
    }

    public override string Description
    {
      get { return "Manages IIS 7.5."; }
    }

    public override void Execute(string configName, IisConfig config, string subCommand, Dictionary<string, string> args)
    {
      var serverManager = new ServerManager();
      var targetSite = serverManager.Sites[config.Site];

      if(targetSite == null)
      {
        throw new ServerManagerException(string.Format("The web server doesn't contain the site {0}", config.Site));
      }
      
      var newApp = CreateNewApplication(config,serverManager, targetSite);
      
      serverManager.CommitChanges();
    }

    private static Application CreateNewApplication(IisConfig config, ServerManager serverManager, Site targetSite)
    {
      var physicalPath = GetPhysicalApplicationPath(config, targetSite);

      var app = targetSite.Applications.Add("/" + config.Name, physicalPath);

      var appPool = serverManager.ApplicationPools.Add(config.Name);

      app.ApplicationPoolName = appPool.Name;

      return app;
    }

    private static string GetPhysicalApplicationPath(IisConfig config, Site targetSite)
    {
      return Path.Combine(targetSite.Applications["/"].VirtualDirectories["/"].PhysicalPath, config.Path);
    }

    protected override Dictionary<string, IisConfig> SampleConfiguration()
    {
      return new Dictionary<string, IisConfig>
      {
        {
          "Orchard", new IisConfig
          {
            ClrVersion = "v4.0", 
            Site = "Default Web Site", 
            Port = 8080, 
            Name = "Orchard", 
            Path = @"orchard\src\Orchard.Web"
          }
        }
      };
    }
  }
}

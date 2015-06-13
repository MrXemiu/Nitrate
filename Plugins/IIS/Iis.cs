using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.ServiceProcess;
using System.Text;
using Microsoft.Web.Administration;
using Nitrate.Plugins;

namespace Nitrate.Plugins.Iis
{

    public class IisConfig
    {
        public string Site { get; set; }
        public string Path { get; set; }
        public string Name { get; set; }
        public bool BrowseOnStart { get; set; }

        public IisConfig()
        {
            Site = "Default Web Site";
            Name = "Orchard";
            Path = @"orchard\src\Orchard.Web";
        }
    }

    [Export(typeof(IPlugin))]
    public class Iis : BasePlugin<IisConfig>
    {
        private const string IisServiceName = "W3SVC";
        private const int TimeOutMs = 15000;

        private static class Commands
        {
            public const string Configure = "configure";
            public const string Recycle = "recycle";
            public const string Restart = "restart";
        }

        public override string InstallationInstructions
        {
            get { return "Please install IIS by searching the Start menu for \"Turn Windows features on or off\"."; }
        }

        public override string Description
        {
            get { return "Manages IIS 7+"; }
        }

        private IDictionary<string, SubCommand> _subCommands;
        public override IDictionary<string, SubCommand> SubCommands
        {
            get
            {
                if (_subCommands == null)
                {
                    _subCommands = new Dictionary<string, SubCommand> {
                        { Commands.Configure, new SubCommand { Description = "configures a web application in IIS" } },
                        { Commands.Restart, new SubCommand {Description = "restarts the IIS service"} },
                        { Commands.Recycle, new SubCommand {Description = "recycles the configured application's app pool"} }
                    };
                }
                return _subCommands;
            }
        }

        public override void Execute(string configName, IisConfig config, string subCommand, Dictionary<string, string> args)
        {
            switch (subCommand)
            {
                case Commands.Configure:
                    ConfigureWebApplication(config);
                    break;
                case Commands.Recycle:
                    RecycleAppPool(config);
                    break;
                case Commands.Restart:
                    RestartIis();
                    break;
            }
        }

        private void RestartIis()
        {
            var service = new ServiceController(IisServiceName);

            var stopTimetMs = Environment.TickCount;
            var timeOutDurationMs = TimeSpan.FromMilliseconds(TimeOutMs);

            service.Stop();
            service.WaitForStatus(ServiceControllerStatus.Stopped, timeOutDurationMs);

            var startTimeMs = Environment.TickCount;
            timeOutDurationMs = TimeSpan.FromMilliseconds(TimeOutMs - (startTimeMs - stopTimetMs));

            service.Start();
            service.WaitForStatus(ServiceControllerStatus.Running, timeOutDurationMs);
        }

        private void RecycleAppPool(IisConfig config)
        {
            var serverManager = new ServerManager();

            var localWebSite = GetLocalWebSite(serverManager, config);

            var app = GetWebApplication(config, serverManager, localWebSite, false);

            var appPool = serverManager.ApplicationPools[app.ApplicationPoolName];

            appPool.Recycle();
        }

        private void ConfigureWebApplication(IisConfig config)
        {
            var serverManager = new ServerManager();

            var localWebSite = GetLocalWebSite(serverManager, config);

            var app = GetWebApplication(config, serverManager, localWebSite);

            if (config.BrowseOnStart) LaunchApplication(localWebSite, app);
        }

        private Site GetLocalWebSite(ServerManager serverManager, IisConfig config)
        {
            var site = serverManager.Sites[config.Site];

            if (site == null)
            {
                throw new ServerManagerException(string.Format("The web server doesn't contain the site {0}", config.Site));
            }

            return site;
        }

        private static Application GetWebApplication(IisConfig config, ServerManager serverManager, Site localWebSite, bool createIfNeeded = true)
        {
            var relativeAppPath = Path.Combine("/", config.Name);
            var app = localWebSite.Applications[relativeAppPath];
            if (app == null && createIfNeeded)
                app = CreateNewApplication(config, serverManager, localWebSite);
            return app;
        }

        private static Application CreateNewApplication(IisConfig config, ServerManager serverManager, Site localWebSite)
        {
            var appPath = GetPhysicalApplicationPath(config);

            if (!Directory.Exists(appPath)) throw new DirectoryNotFoundException(string.Format("The directory {0} doesn't exist.", config.Path));

            var app = localWebSite.Applications.Add("/" + config.Name, appPath);

            var appPool = serverManager.ApplicationPools[config.Name] ?? serverManager.ApplicationPools.Add(config.Name);

            app.ApplicationPoolName = appPool.Name;

            serverManager.CommitChanges();

            return app;
        }

        private static string GetPhysicalApplicationPath(IisConfig config)
        {
            return Path.Combine(Config.Current.Path, config.Path);
        }

        private void LaunchApplication(Site localWebSite, Application app)
        {
            var binding = localWebSite.Bindings.First();
            var uriBuilder = new UriBuilder(binding.Protocol, "localhost", binding.EndPoint.Port, app.Path);
            Process.Start(uriBuilder.ToString());
        }

        protected override Dictionary<string, IisConfig> SampleConfiguration()
        {
            return new Dictionary<string, IisConfig>
              {
                {
                  "Orchard", new IisConfig
                  {
                    Site = "Default Web Site",
                    Name = "Orchard",
                    Path = @"orchard\src\Orchard.Web",
                    BrowseOnStart = true
                  }
                }
              };
        }
    }
}

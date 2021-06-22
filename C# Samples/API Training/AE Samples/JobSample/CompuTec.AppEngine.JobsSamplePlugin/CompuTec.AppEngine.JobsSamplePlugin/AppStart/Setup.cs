using CompuTec.AppEngine.Base.Infrastructure.Plugins;
using System;

namespace CompuTec.AppEngine.JobsSamplePlugin.AppStart
{
    public class Setup : PluginSetup
    {
        public override bool CheckUpdate(Version currentVersion)
        {
            return false;
        }

        public override Version Update(string token)
        {
            return new Version(1, 0, 0, 0);
        }


        public override Version Version { get; }
    }
}

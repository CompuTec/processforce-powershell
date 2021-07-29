using System;
using System.Collections.Generic;
using CompuTec.AppEngine.Base.Infrastructure.Configuration;
using CompuTec.AppEngine.Base.Infrastructure.Plugins;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CompuTec.AppEngine.FirstPlugin.Setup.SettingsCollections
{
    class CountControllerSettings : SettingsCollection<IPluginConfiguration>
    {
        public override List<SettingDefinition> GetSettings()
        {
            //var lc = new LabelConfiguration();

            List<SettingDefinition> settings = new List<SettingDefinition>();


            settings.Add(new SettingDefinition<string>("Message", "Action Completed", false, true));

            return settings;
        }
    }
}


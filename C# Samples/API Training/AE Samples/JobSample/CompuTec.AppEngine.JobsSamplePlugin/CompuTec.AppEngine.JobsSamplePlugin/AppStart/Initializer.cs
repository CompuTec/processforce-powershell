using CompuTec.AppEngine.Base.Infrastructure.Plugins;
using CompuTec.AppEngine.Base.Infrastructure.Translations;
using System.Reflection;

namespace CompuTec.AppEngine.JobsSamplePlugin.AppStart
{
    public class Initializer : PluginInitializer
    {
        public override TranslationStreamDelegate[] TranslationStreamDelegate => new TranslationStreamDelegate[]
        {
            () => Assembly.GetAssembly(this.GetType()).GetManifestResourceStream("CompuTec.AppEngine.JobsSamplePlugin.Translations.messages.xml")
        };

        public override ODataBuilderDelegate ODataBuilderDelegate => builder =>
        {
        };
    }
}

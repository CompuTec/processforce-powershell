using CompuTec.AppEngine.Base.Infrastructure.Plugins;
using CompuTec.AppEngine.Base.Infrastructure.Translations;
using System.Reflection;
using CompuTec.AppEngine.First;

namespace CompuTec.AppEngine.FirstPlugin.AppStart
{
    public class Initializer : PluginInitializer
    {
        public override TranslationStreamDelegate[] TranslationStreamDelegate => new TranslationStreamDelegate[]
        {
            () => Assembly.GetAssembly(this.GetType()).GetManifestResourceStream("CompuTec.AppEngine.FirstPlugin.Translations.messages.xml")
        };

        public override ODataBuilderDelegate ODataBuilderDelegate => builder =>
        {
        };

        public override void BeforeInitialize()
        {
            base.BeforeInitialize();
            var myInfo = new Info();
            CompuTec.Core2.CoreManager.InitializeLibrary(myInfo);
        }
    }
}

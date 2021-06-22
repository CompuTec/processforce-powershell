using CompuTec.AppEngine.Base.Infrastructure.Plugins;
using CompuTec.AppEngine.Base.Infrastructure.Translations;
using System.Reflection;
using CompuTec.AppEngine.FirstPlugin.API;

namespace CompuTec.AppEngine.FirstPlugin.AppStart
{
    public class Initializer : PluginInitializer
    {
        public override TranslationStreamDelegate[] TranslationStreamDelegate => new TranslationStreamDelegate[]
        {
            ///Point to a location of translations 
            () => Assembly.GetAssembly(this.GetType()).GetManifestResourceStream("CompuTec.AppEngine.FirstPlugin.Translations.messages.xml")
        };

        public override ODataBuilderDelegate ODataBuilderDelegate => builder =>
        {
            ///here you can register custom Odata controlers
        };

        public override void BeforeInitialize()
        {
            base.BeforeInitialize();
             
            var myInfo = new FirstPluginInfo();
            CompuTec.Core2.CoreManager.InitializeLibrary(myInfo);
        }
    }
}

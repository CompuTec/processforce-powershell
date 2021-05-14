using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using CompuTec.AppEngine.Base.Infrastructure.Plugins;
using CompuTec.AppEngine.First.DBInstall.Tables;
using CompuTec.AppEngine.First;

namespace CompuTec.AppEngine.FirstPlugin.Plugin.AppStart
{
    public class Setup : PluginSetup
    {
        public override bool CheckUpdate(Version currentVersion)
        {
            return currentVersion < new Version(Info.NameVersion);
        }

        public override Version Update(string token)
        {

            var info = new Info();

            Console.WriteLine("Update");

            List<CompuTec.Core2.DI.Setup.UDO.Model.ICustomField> customUdoFieldList =  First.DBInstall.CustomUDOFields.getCustomFields();
            CompuTec.Core2.DI.Setup.UDO.Setup setup = new CompuTec.Core2.DI.Setup.UDO.Setup(token, customUdoFieldList, false, System.Reflection.Assembly
                .GetAssembly(typeof(ToDoTable)), "CompuTec.AppEngine.First.DBInstall.Tables", "CompuTec.AppEngine.First.DBInstall.Tables",
                "CompuTec.AppEngine.First.DBInstall.Tables", "CompuTec.AppEngine.First.DBInstall.Tables", "CompuTec.AppEngine.First.DBInstall.Tables");

            setup.BaseLibInformation = info;

            if (setup.IsUpdateRequiredNew(true))
            {
                Console.WriteLine("Instaling...");
                var updateResult = setup.Update();

                if (!updateResult.Success)
                {    
                    var message = new StringBuilder();

                    updateResult.Errors.ForEach(e =>
                    {
                        message.Append(e.Message);
                    });

                    throw new Exception(message.ToString());
                }


                Console.WriteLine(updateResult.ToString());
            }

            Console.WriteLine("Install finish");

            return Version;
        }

        public override Version Version => new Version(Info.NameVersion);
    }
}  
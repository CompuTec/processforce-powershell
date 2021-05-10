using CompuTec.AppEngine.FirstPlugin.Api.Services;
using StructureMap;

namespace CompuTec.AppEngine.FirstPlugin.AppStart
{
    public class DIRegistry : Registry
    {
        public DIRegistry()
        {
            For<DataSource>().Singleton().Use<DataSource>();
            For<ITodoService>().Use<TodoService>();
        }
    }
}

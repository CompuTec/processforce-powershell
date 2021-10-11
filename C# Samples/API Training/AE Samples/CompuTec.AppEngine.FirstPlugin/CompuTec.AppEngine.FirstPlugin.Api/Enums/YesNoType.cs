using CompuTec.Core2.DI.Attributes;

namespace CompuTec.AppEngine.FirstPlugin.API.Enums
{
	[EnumType(new int[] { 1, 2 }, new string[] { "Y", "N" }, 2)]
	public enum YesNoType { Yes = 1, No = 2 }
}

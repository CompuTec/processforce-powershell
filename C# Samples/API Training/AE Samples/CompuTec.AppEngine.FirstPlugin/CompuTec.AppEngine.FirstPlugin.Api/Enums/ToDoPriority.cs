using CompuTec.Core2.DI.Attributes;

namespace CompuTec.AppEngine.FirstPlugin.API.Enums
{
	[EnumType(new int[] { 1, 2, 3 }, new string[] { "L", "M", "H" }, 2)]
	public enum ToDoPriority
	{
		Low = 1, Medium = 2, Huge = 3
	}
}

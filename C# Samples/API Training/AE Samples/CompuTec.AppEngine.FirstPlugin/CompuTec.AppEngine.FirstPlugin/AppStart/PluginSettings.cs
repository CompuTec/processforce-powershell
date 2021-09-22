using CompuTec.AppEngine.Base.Infrastructure.Configuration;
using CompuTec.AppEngine.Base.Infrastructure.Exceptions;
using CompuTec.AppEngine.Base.Infrastructure.Plugins;
using CompuTec.AppEngine.FirstPlugin.API.Enums;
using System;
using System.Collections.Generic;

namespace CompuTec.AppEngine.FirstPlugin.AppStart
{
	public class PluginSettings : SettingsCollection<IPluginConfiguration>
	{
		public override List<SettingDefinition> GetSettings()
		{
			List<SettingDefinition> settings = new List<SettingDefinition>();

			#region  SalesOrderToApproveEventJob node
			settings.Add(new SettingDefinition<Dictionary<string, object>>("SalesOrderToApproveEventJob", new Dictionary<string, object>(), false, true));

			settings.Add(new SettingDefinition<string>($"SalesOrderToApproveEventJob:TaskPriority", ToDoPriority.Medium.ToString(), PriorityValidation, null, false, true));
			#endregion

			return settings;
		}

		public static void PriorityValidation(string key, string newStatus, IConfiguration configuration)
		{
			if (!Enum.TryParse<ToDoPriority>(newStatus, out var a))
			{
				List<string> allowedPriorities = new List<string>();
				foreach (var e in (typeof(ToDoPriority).GetEnumValues()))
				{
					allowedPriorities.Add(e.ToString());
				}

				throw new AppEngineException($"Incorrect value for Priority: {newStatus}. Allowed values: {string.Join(",", allowedPriorities)}");
			}

		}
	}
}

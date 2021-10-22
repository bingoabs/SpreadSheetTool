using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpreadSheetTool.Models
{
	public static class SpreadSheetExtension
	{
		public static SharedPointSSRecord GetValueOrDefault(this Dictionary<string, SharedPointSSRecord> container,
			string key, SharedPointSSRecord default_value)
		{
			if (container.ContainsKey(key))
				return container[key];
			return default_value;
		}
		public static bool IsNullOrEmptyOrMeaningless(this string message)
		{
			if (string.IsNullOrWhiteSpace(message))
				return true;
			return message.Trim() == "-";
		}
	}
}

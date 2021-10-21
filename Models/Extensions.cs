using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpreadSheetTool.Models
{
	public static class DictionaryExtension
	{
		public static SharedPointSSRecord GetValueOrDefault(this Dictionary<string, SharedPointSSRecord> container,
			string key, SharedPointSSRecord default_value)
		{
			if (container.ContainsKey(key))
				return container[key];
			return default_value;
		}
	}
}

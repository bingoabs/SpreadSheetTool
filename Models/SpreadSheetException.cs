using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpreadSheetTool.Models
{
	class SpreadSheetException: Exception
	{
		public SpreadSheetException(string message) : base(message) { }
	}
}

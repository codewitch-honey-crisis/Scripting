using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Scripting;
namespace ScriptingDemo
{
	class Program
	{
		static void Main(string[] args)
		{
			using (var host = new ScriptHost())
			{
				var engine = host.Create("javascript");
				Console.WriteLine(engine.Evaluate("2+2"));
				engine.AddCode("var i = 1;");
				dynamic obj = engine.Script;
				Console.WriteLine(obj.i);
			}
			
		}
	}
}

using System;
using System.Runtime.InteropServices;
using Scripting;
namespace ScriptingDemo
{
	[ComVisible(true)]
	[ClassInterface(ClassInterfaceType.AutoDispatch /*.AutoDual */)]  // automaticly implements com interface IDispatch ( and IUnknown )
	public class AppObject
	{
		/// <summary>
		/// allows scripts to write to the console
		/// </summary>
		public void writeLine(string s) { Console.WriteLine(s); }

	}
	
	class Program
	{
		static void Main(string[] args)
		{
			using (var host = new ScriptHost())
			{
				// create a javascript engine
				var engine = host.Create("javascript"); // you can use the ChakraJS const here for IE9+'s fast engine
				// evaluate an expression
				Console.WriteLine(engine.Evaluate("2+2"));
				// add some code to the current script
				engine.AddCode("var i = 1;");
				// get the object for the script
				dynamic obj = engine.Script;
				// print var i
				Console.WriteLine(obj.i);
				// add an app object to the script
				engine.Globals.Add("app", new AppObject());
				// let the script call the app object
				engine.Run("app.writeLine('Hello World!');");
			}
			
		}
	}
}

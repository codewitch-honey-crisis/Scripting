using System;
using System.Runtime.InteropServices.ComTypes;

namespace Scripting
{
	public class ScriptException : Exception
	{
		internal ScriptException(object err) : base(_FillExInfo(err))
		{
			EXCEPINFO e;
			var ase = err as IActiveScriptError;
			ase.GetExceptionInfo(out e);
			string src;
			ase.GetSourceLineText(out src);
			Code = src;
			uint ctx;
			uint lineNo;
			int pos;
			ase.GetSourcePosition(out ctx, out lineNo, out pos);
			ScriptSource = e.bstrSource;
			Line = unchecked((int)lineNo) + 1;
			Position = pos;
		}
		static string _FillExInfo(object err)
		{
			EXCEPINFO e;
			var ase = err as IActiveScriptError;
			ase.GetExceptionInfo(out e);
			return e.bstrDescription;
		}
		public string ScriptSource { get; }
		public string Code { get; }
		public int Line { get; }
		public int Position { get; }

	}
}

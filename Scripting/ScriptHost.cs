using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices.ComTypes;

namespace Scripting
{
	
	public sealed class ScriptHost : IDisposable, IActiveScriptSite
	{

		// IE9+ fast JS script engine
		public const string ChakraJS = "{16d51579-a30b-4c8b-a276-0ff4dc41e755}";

		private const int S_OK = 0;
		private const int TYPE_E_ELEMENTNOTFOUND = unchecked((int)(0x8002802B));
		private const int E_NOTIMPL = -2147467263;
		IList<ScriptEngine> _engines = new List<ScriptEngine>();
		internal bool RemoveEngine(ScriptEngine engine)
		{
			if(_engines.Contains(engine))
			{
				engine.Inner.Close();
				return _engines.Remove(engine);
			}
			return false;
		}
		public ScriptEngine Create(string language)
		{
			Type type = null;
			Guid clsid;
			if (Guid.TryParse(language, out clsid))
			{
				type = Type.GetTypeFromCLSID(clsid, true);
			}
			else
			{
				type = Type.GetTypeFromProgID(language, true);
			}
			if (null == type)
				throw new ArgumentException("The specified script language " + language + " is not installed on this system","language");
			var @as = Activator.CreateInstance(type) as IActiveScript;
			if(null==@as)
				throw new ArgumentException("The specified script language " + language + " does not refer to a scripting engine", "language");
			@as.SetScriptSite(this);
			var result = new ScriptEngine(@as);
			var p = @as as IActiveScriptParse;
			if (null == p)
				throw new NotSupportedException("The specified script language " + language + " is not supported");
			p.InitNew();
			_engines.Add(result);
			return result;
		}
		void IActiveScriptSite.GetDocVersionString(out string v)
		{
			v = string.Empty;
		}
		void IActiveScriptSite.GetItemInfo(string name, uint returnMask, out object item, IntPtr ppti)
		{
			Console.WriteLine("Fetch item " + name);
			item = null;
			if ((returnMask & (uint)SCRIPTINFOFLAGS.SCRIPTINFO_ITYPEINFO) == (uint)SCRIPTINFOFLAGS.SCRIPTINFO_ITYPEINFO)
				return;

			for (int ic=_engines.Count,i=0;i<ic;++i)
			{
				var e = _engines[i];
				if(e.Globals.TryGetValue(name,out item))
				{
					return;
				}
			}
		}

		void IActiveScriptSite.GetLCID(out uint id)
		{
			id = 2048u; // DEFAULT
		}

		void IActiveScriptSite.OnEnterScript()
		{
			
		}

		void IActiveScriptSite.OnLeaveScript()
		{
			
		}

		void IActiveScriptSite.OnScriptError(object err)
		{
			throw new ScriptException(err);
		}

		void IActiveScriptSite.OnScriptTerminate(ref object result, ref EXCEPINFO info)
		{
			
		}

		void IActiveScriptSite.OnStateChange(SCRIPTSTATE state)
		{
			
		}

		#region IDisposable Support
		void _Dispose(bool disposing)
		{
			if (null!=_engines)
			{
				while(0<_engines.Count)
					RemoveEngine(_engines[_engines.Count - 1]);
				_engines = null;
			}
		}

		~ScriptHost()
		{
		  _Dispose(false);
		}

		// This code added to correctly implement the disposable pattern.
		void IDisposable.Dispose()
		{
			Close();
		}
		public void Close()
		{
			_Dispose(true);
			GC.SuppressFinalize(this);
		}
		#endregion

	}
}

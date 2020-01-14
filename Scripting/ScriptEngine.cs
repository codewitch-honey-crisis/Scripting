using System;
using System.Runtime.InteropServices.ComTypes;

namespace Scripting
{
	public sealed class ScriptEngine : IDisposable
	{
		static Guid _IID_IUnknown = Guid.Parse(@"00000000-0000-0000-C000-000000000046");
		IActiveScript _inner;
		internal ScriptEngine(IActiveScript inner)
		{
			_inner = inner;
		}
		internal IActiveScript Inner {
			get { return _inner; }
		}
		public ScriptHost Host { 
			get {
				object site;
				_inner.GetScriptSite(ref _IID_IUnknown, out site);
				return site as ScriptHost;
			} 
		}
		public object Evaluate(string expression)
		{
			var p = _inner as IActiveScriptParse;
			object result;
			EXCEPINFO excepinfo;
			p.ParseScriptText(expression, "", IntPtr.Zero, "", 0, 0, 32 /*Expression*/, out result, out excepinfo);
			_inner.SetScriptState(SCRIPTSTATE.SCRIPTSTATE_CONNECTED);

			return result;
		}
		public object Run(string statement)
		{
			var p = _inner as IActiveScriptParse;
			object result;
			EXCEPINFO excepinfo;
			p.ParseScriptText(statement, "", IntPtr.Zero, "", 0, 0, 0, out result, out excepinfo);
			_inner.SetScriptState(SCRIPTSTATE.SCRIPTSTATE_CONNECTED);

			return result;
		}
		public object AddCode(string code)
		{
			var p = _inner as IActiveScriptParse;
			object result;
			EXCEPINFO excepinfo;
			p.ParseScriptText(code, "", IntPtr.Zero, "", 0, 0, 66 /*IsVisible | Persistent*/, out result, out excepinfo);
			_inner.SetScriptState(SCRIPTSTATE.SCRIPTSTATE_CONNECTED);

			return result;
		}
		public object Script {
			get {
				IntPtr dispatch;
				_inner.GetScriptDispatch(null, out dispatch);
				return System.Runtime.InteropServices.Marshal.GetObjectForIUnknown(dispatch);
			}
		}
		#region IDisposable Support
		void _Dispose(bool disposing)
		{
			if (null!=_inner)
			{
				var host = Host;
				if(null!=host)
				{
					if (!host.RemoveEngine(this))
						_inner.Close();
				} else
				{
					_inner.Close();
				}
				_inner = null;
			}
		}

		~ScriptEngine()
		{
		   // Do not change this code. Put cleanup code in Dispose(bool disposing) above.
		   _Dispose(false);
		}

		// This code added to correctly implement the disposable pattern.
		public void Close()
		{
			_Dispose(true);
			GC.SuppressFinalize(this);
		}
		void IDisposable.Dispose()
		{
			Close();
		}
		#endregion
	}
}

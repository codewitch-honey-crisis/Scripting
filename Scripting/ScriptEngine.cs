using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices.ComTypes;

namespace Scripting
{
	public sealed class ScriptEngine : IDisposable
	{
		static Guid _IID_IUnknown = Guid.Parse(@"00000000-0000-0000-C000-000000000046");
		IActiveScript _inner;
		_GlobalDictionary _globals;
		internal ScriptEngine(IActiveScript inner)
		{
			_inner = inner;
			_globals = new _GlobalDictionary(this);
		}
		public IDictionary<string,object> Globals {
			get {
				return _globals;
			}
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

		private sealed class _GlobalDictionary: IDictionary<string, object>
		{
			Dictionary<string, object> _inner;
			WeakReference<ScriptEngine> _outer;
			internal _GlobalDictionary(ScriptEngine outer)
			{
				_outer = new WeakReference<ScriptEngine>(outer);
				_inner = new Dictionary<string, object>(StringComparer.InvariantCulture);
			}
			public object this[string key] { get => _inner[key]; set => _inner[key] = value; }

			public ICollection<string> Keys => _inner.Keys;

			public ICollection<object> Values => _inner.Values;

			public int Count => _inner.Count;

			public bool IsReadOnly => false;

			public void Add(string key, object value)
			{
				ScriptEngine outer;
				if(_outer.TryGetTarget(out outer))
				{
					outer.Inner.AddNamedItem(key, (uint)(SCRIPTITEMFLAGS.SCRIPTITEM_NOCODE | SCRIPTITEMFLAGS.SCRIPTITEM_ISVISIBLE));
				}
				_inner.Add(key, value);
			}

			void ICollection<KeyValuePair<string,object>>.Add(KeyValuePair<string, object> item)
			{
				Add(item.Key, item.Value);
			}

			public void Clear()
			{
				_inner.Clear();
			}

			bool ICollection<KeyValuePair<string, object>>.Contains(KeyValuePair<string, object> item)
			{
				return (_inner as ICollection<KeyValuePair<string, object>>).Contains(item);
			}

			public bool ContainsKey(string key)
			{
				return _inner.ContainsKey(key);
			}

			public void CopyTo(KeyValuePair<string, object>[] array, int arrayIndex)
			{
				(_inner as ICollection < KeyValuePair<string, object> >).CopyTo(array, arrayIndex);
			}

			public IEnumerator<KeyValuePair<string, object>> GetEnumerator()
			{
				return _inner.GetEnumerator();
			}

			public bool Remove(string key)
			{
				return _inner.Remove(key);
			}

			bool ICollection<KeyValuePair<string, object>>.Remove(KeyValuePair<string, object> item)
			{
				return (_inner as ICollection<KeyValuePair<string, object>>).Remove(item);
			}

			public bool TryGetValue(string key, out object value)
			{
				return _inner.TryGetValue(key, out value);
			}

			IEnumerator IEnumerable.GetEnumerator()
			{
				return _inner.GetEnumerator();
			}
		}
	}
}

// simple .NET Script Host
// De Plancke Ronny

// credits
// C# Interfaces for the Windows Scripting Host
// by Uwe Keim , http://www.codeproject.com/KB/cs/ZetaScriptingHost.aspx
// based on C++ Script Host
// by Ladislav Nevery , http://www.codeproject.com/KB/cpp/ScriptYourApps.aspx

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.Runtime.InteropServices;
using ComTypes = System.Runtime.InteropServices.ComTypes;

namespace Scripting
{
    /// <summary>
    /// instance of this class will be created to represent the named item declared with the SCRIPTITEMFLAGS.SCRIPTITEM_NOCODE flag.
    /// this instance will be visible to the script.
    /// </summary>
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch /*.AutoDual */)]  // automaticly implements com interface IDispatch ( and IUnknown )
    public class ScriptInterface
    {
        /// <summary>
        /// allows scripts to write to the console
        /// </summary>
        public void WriteLine(string s) { Console.WriteLine(s); }

    }

    /// <summary>
    /// declare JavaScript as a COM coclass  
    /// </summary>
    [ComImport, Guid("f414c260-6ac0-11cf-b6d1-00aa00bbbb58")]
    class JavaScript
    {
    }

    /// <summary>
    /// declare VbScript as a COM coclass  
    /// </summary>
    [ComImport, Guid("b54f3741-5b07-11cf-a4b0-00aa004a55e8")]
    class VbScript
    {
    }

    public static class _ScriptHost
    {
        public static void run()
        {

            IActiveScript script;
            IActiveScriptParse parse;

            script = Activator.CreateInstance(Type.GetTypeFromProgID("Javascript")) as IActiveScript;
            parse = (IActiveScriptParse)script;

            ComTypes.EXCEPINFO excepinfo;

            ScriptSite host = new ScriptSite();

            script.SetScriptSite(host);
            script.AddNamedItem(@"app", (uint)(SCRIPTITEMFLAGS.SCRIPTITEM_ISVISIBLE | SCRIPTITEMFLAGS.SCRIPTITEM_NOCODE));

            // Hello world java script
            string source = @"app.WriteLine('hello world.'); ";

            try
            {
                parse.InitNew();
                object result;
                parse.ParseScriptText(source, "", IntPtr.Zero, "",0, 0, 0,out result, out excepinfo);
                script.SetScriptState(SCRIPTSTATE.SCRIPTSTATE_CONNECTED);
                script.Close();
            }
            catch (COMException e)
            {
                Console.WriteLine("Exception thrown by the ScriptEngine : {0}", e.Message);
            }
        }
    }

    /// <summary>
    /// Site for the Windows Script engine.
    /// </summary>
    /// <see cref="http://msdn.microsoft.com/library/en-us/script56/html/4d604a11-5365-46cf-ab71-39b3dbbe9f22.asp"/>
    class ScriptSite : IActiveScriptSite
    {

        ScriptInterface s_interface = new ScriptInterface();

        #region IActiveScriptSite Members

        /// <summary>
        /// Retrieves the locale identifier associated with the host\'s user interface.
        /// </summary>
        public void GetLCID(out uint id)
        {
            id = 2048; // use default locale
        }

        /// <summary>
        /// Allows the scripting engine to obtain information about an item added with the IActiveScript::AddNamedItem method.
        /// </summary>
        public void GetItemInfo(string name, uint returnMask, out object item, IntPtr ppti)
        {
            item = null;
            if ((returnMask & (uint)SCRIPTINFOFLAGS.SCRIPTINFO_IUNKNOWN) != 0)
                item = s_interface;
            else
                item = null;

            // we dont\'t provide type information
            // the item object cannot source events and name binding must be realized with the IDispatch::GetIDsOfNames method
            ppti = IntPtr.Zero; // no events on our item class ,

        }

        /// <summary>
        /// Informs the host that an execution error occurred while the engine was running the script.
        /// </summary>
        public void OnScriptError(object err)
        {
            ComTypes.EXCEPINFO e;
            var ase = err as IActiveScriptError; 
            ase.GetExceptionInfo(out e);
            string txt;
            ase.GetSourceLineText(out txt);
            uint ctx;
            uint lineNo;
            int pos;
            ase.GetSourcePosition(out ctx, out lineNo,out pos);
            
            Console.WriteLine("Exception = {0} , source = {1} at line {2}, position {3}, text = {4}", e.bstrDescription, e.bstrSource,lineNo,pos,txt);
        }


        /// <summary>
        /// Retrieves a host-defined string that uniquely identifies the current document version
        /// </summary>
        public void GetDocVersionString(out string v)
        { v = string.Empty; }


        /// <summary>
        /// Informs the host that the script has completed execution.
        /// </summary>
        /// <param name="result">script results</param>
        /// <param name="info">EXCEPINFO structure that contains exception information generated when the script terminated, or NULL if no exception was generated. </param>
        public void OnScriptTerminate(ref object result, ref ComTypes.EXCEPINFO info)
        { }

        /// <summary>
        /// Informs the host that the scripting engine has changed states.
        /// </summary>
        public void OnStateChange(SCRIPTSTATE state)
        { }

        /// <summary>
        /// Informs the host that the scripting engine has begun executing the script code.
        /// </summary>
        public void OnEnterScript()
        { }

        /// <summary>
        /// Informs the host that the scripting engine has returned from executing script code.
        /// </summary>
        public void OnLeaveScript()
        { }

        #endregion
    }
}
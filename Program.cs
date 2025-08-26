using Microsoft.ClearScript;
using Microsoft.ClearScript.V8;
using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Xml;


namespace Wysh {
	internal class Program {
        protected Assembly _appAssembly;
        internal Assembly Assembly {
            get { return _appAssembly; }
        }

        protected string _appNamespace;
        internal string Namespace {
            get { return _appNamespace; }
        }

        protected Internal.ResourceManager rsrc;
        public Internal.ResourceManager Rsrc
        {
            get { return rsrc; }
        }

        static void Main(string[] args) {
			Program p = new Program(ParseCommand(args));
		}

		internal Program(Dictionary<string,string> scriptSetup) {
            _appAssembly = Assembly.GetEntryAssembly();
            Type[] types = _appAssembly.GetTypes();
			_appNamespace = typeof(Program).Namespace;

            this.rsrc = new Internal.ResourceManager(this);

            V8ScriptEngine engine = AcquireEngine(scriptSetup);
            AddWysh(engine);

            // execute based on file extension
            if (scriptSetup["_ScriptName"].EndsWith(".wysh")) ExecuteWysh(engine, scriptSetup["_ScriptName"], scriptSetup);
            else if (scriptSetup["_ScriptName"].EndsWith(".js")) ExecuteJS(engine, scriptSetup["_ScriptName"], scriptSetup);
            else Console.WriteLine("Invalid file type.");
        }

        private static void ExecuteWysh(V8ScriptEngine engine, string scriptFile, Dictionary<string, string> scriptSetup) {
            // parse as XML
            XmlDocument wsh = new XmlDocument();
            wsh.Load(scriptSetup["_ScriptName"]);

            // instance resources
            Internal.ScriptResourceProxy resproxy = new Internal.ScriptResourceProxy(wsh);
            engine.AddHostObject("Resources", resproxy);

            // find jobs
            XmlNodeList jobs = wsh.SelectNodes("/package/job");
            if (jobs == null || jobs.Count == 0) return;

            // using first job for now...
            XmlNodeList scripts = jobs[0].SelectNodes("script");
            if (scripts == null || scripts.Count == 0) return;

            // execute scripts in order
            // if 'src', load based on src
            // not checking lang. All must be js
            for (int i = 0; i < scripts.Count; i++) {
                if (scripts[i].Attributes["src"] != null)
                    engine.ExecuteDocument(scripts[i].Attributes["src"].Value);
                else engine.Execute(scripts[i].InnerText);
            }
        }

        private static void ExecuteJS(V8ScriptEngine engine, string scriptFile, Dictionary<string,string> scriptSetup) {
            engine.ExecuteDocument(scriptFile);
        }

		private static void AddComTypes(V8ScriptEngine engine) {
            engine.AddCOMType("XMLHttpRequest", "MSXML2.XMLHTTP");

            engine.AddCOMType("FileSystemObject", "Scripting.FileSystemObject");
            engine.AddCOMType("Shell", "WScript.Shell");
            engine.AddCOMType("Network", "WScript.Network");
            engine.AddCOMType("Dictionary", new Guid("{ee09b103-97e0-11cf-978f-00a02463e06f}"));
			engine.AddCOMType("StreamBuffer", "ADODB.Stream");
            engine.AddCOMType("CDOMessage", "CDO.Message");
            engine.AddCOMType("CDOConfiguration", "CDO.Configuration");

            engine.AddCOMType("ExcelApp", HostItemFlags.DirectAccess, "Excel.Application");
            engine.AddCOMType("DOMDocument", "Msxml2.DOMDocument.6.0");

            // add sleep, scriptname, etc...
        }

		private static void SetSearchPaths(V8ScriptEngine engine) {
			// add search paths from internal modules (none yet)

			// and enable file loading
			engine.DocumentSettings.AccessFlags = Microsoft.ClearScript.DocumentAccessFlags.EnableFileLoading | Microsoft.ClearScript.DocumentAccessFlags.EnforceRelativePrefix;
		}

        private static V8ScriptEngine AcquireEngine(Dictionary<string,string> scriptSetup) {
            V8ScriptEngine engine = new V8ScriptEngine(V8ScriptEngineFlags.EnableArrayConversion | V8ScriptEngineFlags.EnableDateTimeConversion | V8ScriptEngineFlags.EnableStringifyEnhancements);
            
            engine.AddHostType("Console", typeof(Console));
            engine.AddHostType("Environment", typeof(Environment));

            AddComTypes(engine);
            SetSearchPaths(engine);

            return engine;
        }

        private void AddWysh(V8ScriptEngine engine) {
            engine.Execute(rsrc.GetString("Lib.Base.js"));
            engine.Execute(rsrc.GetString("Lib.Excel.js"));
			engine.Execute(rsrc.GetString("Lib.Email.js"));
			engine.Execute(rsrc.GetString("Lib.Template.js"));
		}

		private static Dictionary<string, string> ParseCommand(string[] args) {
            Dictionary<string, string> scriptSetup = new Dictionary<string, string>();
			// parse all flags up to the script file name
			// all flags after into own array
			for(int i = 0; i < args.Length; i++) {
				if(args[i].StartsWith("-") || args[i].StartsWith("/")) {
					// it's a flag or setting
					// break on : or =
					// set to true or value
					// if key starts with _ then skip
				} else {
					// scriptname
					scriptSetup["_ScriptName"] = args[i];
				}

			}

			return scriptSetup;
		}

	}
}

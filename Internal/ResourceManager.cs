using Newtonsoft.Json;
using System;
using System.IO;
using System.IO.Compression;
using System.Reflection;
using System.Text;

namespace Wysh.Internal
{
	
	/* Adopted from udxLib */
	internal class ResourceManager {
		private Program parent;
		public ResourceManager(Program prog) {
			parent = prog;
		}

        public string GetString(string name) {
            Stream stream = GetStream(name);
            if (stream == null) return "";
            using (StreamReader reader = new StreamReader(stream, Encoding.UTF8)) {
                return reader.ReadToEnd();
            }
        }

        public string GetGzString(string name) {
            Stream str = GetStream(name);
            if (str == null) return null;

            GZipStream gzStream = new GZipStream(str, CompressionMode.Decompress);
            using (StreamReader reader = new StreamReader(gzStream, Encoding.UTF8)) {
                return reader.ReadToEnd();
            }
        }
        
        public object GetJSON(string name) {
            string source = GetGzString(name);
            if (source == null) return null;
            return JsonConvert.DeserializeObject(source);
        }
        public Stream GetStream(string name) {
            try {
                Stream str = parent.Assembly.GetManifestResourceStream(parent.Namespace + "." + name);
                //if (str == null)
                    //EventLog.Log(EventType.Warning, "Could not locate resource " + parent.Namespace + "." + name + ".");
                return str;
            } catch (Exception e) {
                //EventLog.Log(EventType.Error, e.Message);
                throw new Exception(parent.Namespace + ": " + e.Message);
            }
        }

        public static Stream GetStream(Assembly a, string name) {
            try {
                string res = "";
                string[] resList = a.GetManifestResourceNames();
                foreach (string resName in resList) {
                    if (resName.EndsWith(name)) res = resName;
                }
                Stream str = a.GetManifestResourceStream(res);
                //if (str == null)
                    //EventLog.Log(EventType.Warning, "Could not locate resource '" + name + "'.");
                return str;
            } catch (Exception e) {
                //EventLog.Log(EventType.Error, e.Message);
                throw new Exception(a.GetName().Name + ": " + e.Message);
            }
        }

        public static Stream GetStream(string assembly, string name) {
            try {
                Assembly a = Assembly.Load(assembly);
                return GetStream(a, name);
            } catch (Exception e) {
                //EventLog.Log(EventType.Error, e.Message);
                throw new Exception(assembly + ": " + e.Message);
            }
        }

        public static Stream GetStream(Type type, string name) {
            string assembly = type.Assembly.GetName().Name;
            return GetStream(assembly, name);
        }
    }
}

using System.Runtime.CompilerServices;
using System.Text.Json;
using System.Xml;

namespace Wysh.Internal {
	public class ScriptResourceProxy {
		private XmlDocument resDoc;

		public ScriptResourceProxy(XmlDocument xml) {
			this.resDoc = xml;
		}

		public string getText(string id, string t = "text") {
			XmlNode node = resDoc.SelectSingleNode($"//{t}[@id='{id}']");
			if (node == null) return null;
			return node.InnerText.Trim();
		}

		public string getAttrib(string id, string t, string attrib) {
			XmlNode node = resDoc.SelectSingleNode($"//{t}[@id='{id}']");
			if (node == null) return null;
			return node.Attributes[attrib].ToString();
		}

		public string getQuery(string id) {
			return getText(id, "query");
		}

		public string getTemplate(string id) {
			// this should call the engine and proxy the creation of a template instance...
			return getText(id, "template");
		}

		public object getData(string id) {
			string str = getText(id, "json");
			return JsonSerializer.Deserialize(str, typeof(object));
        }
	}
}

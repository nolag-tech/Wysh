class Template {
	constructor(tmpl, defaults, scope) {
		this.tmpl = tmpl;
		this.defaults = defaults;
		this.scope = scope;
		this.blanks = true;
	}

	exec(obj) {
		var output = '';

		var rxFIELD = /%\(([A-Za-z0-9,_|.]*)\)/g;
		var rxIF = /@IF\(([A-Za-z0-9,_|.]*)\)((.|\n)*?)@END/g;
		var rxEACH = /@EACH\(([A-Za-z0-9,_|.]*)\)((.|\n)*?)@END(EACH)?/g;

		var fnEACH = (function (w, g, txt) {
			if (!obj[g]) return "";

			var out = "";
			for (var i = 0; i < obj[g].length; i++) {
				out += Template.execStatic(txt, obj[g][i]);
			}
			return out;
		}).bind(this);

		var fnFIELD = (function (w, g) {
			g = g.split('|');
			f = (g[0]).split(',');
			var cnt = null;
			var fields = [];
			if (f.length == 1) cnt = (typeof obj[g[0]] != 'undefined' && obj[g[0]] != null) ? obj[g[0]] : this.defaults[g[0]];
			else for (var i = 0; i < f.length; i++) fields[i] = (obj[f[i]]) ? obj[f[i]] : this.defaults[f[i]];
			for (var i = 1; i < g.length; i++) {
				if (f.length == 1) cnt = this.scope[g[i]](cnt);
				else cnt = this.scope[g[i]](fields);
			}
			if (cnt === 0 || cnt === -1 || cnt === false) cnt += '';
			return cnt ? cnt : (this.defaults[g[0]] ? this.defaults[g[0]] : (this.blanks ? '' : w));
		}).bind(this);

		var fnIF = (function (w, g, txt) {
			g = g.split('|');
			var fields = (g[0]).split(',');
			var result = true;
			var data;

			if (fields.length == 1) {
				data = (typeof obj[g[0]] != 'undefined' && obj[g[0]] != null) ? obj[g[0]] : this.defaults[g[0]];
				if (typeof data == 'undefined' || data == null) result = false;
			} else {
				data = [];
				for (var i = 0; i < fields.length; i++) data[i] = (obj[fields[i]]) ? obj[fields[i]] : this.defaults[fields[i]];
			}

			for (var i = 1; i < g.length; i++) result = this.scope[g[i]](data) && result;

			return result ? txt : '';
		}).bind(this);


		output = this.tmpl.replace(rxEACH, fnEACH);
		output = output.replace(rxFIELD, fnFIELD);
		output = output.replace(rxIF, fnIF);
		return output;
	}
}

Template.execStatic = function (str, obj) {
	var fn = (function (w, g) {
		return (typeof obj[g] != 'undefined' && obj[g] != null) ? obj[g] : w;
	});
	return str.replace(/%\(([A-Za-z0-9_]*)\)/g, fn);
};

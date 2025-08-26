Object.defineProperty(Object.prototype, "expand", {
    value: function (t) {
        var e = this.prototype || this;
        for (const n in t) Object.prototype.hasOwnProperty.call(t, n) && Object.defineProperty(e, n, {
            value: "function" == typeof t[n] && e == this ? t[n].bind(e) : t[n]
        });
        return e;
    }
});

String.expand({
    trim: function () {
        return this.replace(/^\s+|\s+$/g, "");
    }
});

Array.expand({
    toDictionary: function () {
        var dict = new Dictionary();
        for (let i = 0; i < this.length; i++)
            dict.add(i, this[i]);
        return dict;
    },

    toVBArray: function () {
        var dict = this.toDictionary();
        return dict.Items();
    }
});

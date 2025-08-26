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

Date.expand({
    clone: function () {
        return new Date(this.getTime());
    },

    daysUntil: function (d) {
        return Math.ceil((d.getTime() - this.getTime()) / (1000 * 60 * 60 * 24));
    },

    addDays: function (n) {
        this.setDate(this.getDate() + n);
    },

    toJulian: function () {
        var DAY = 1000 * 60 * 60 * 24;

        // sets start to 12/31 previous year so 1/1 will be day 1 instead of 0
        var start = new Date(this.getFullYear(), 0, 0);
        var delta = d - start;

        return (d.getYear() - 1900) * 1000 + Math.floor(delta / DAY);
    },

    toDateCode: function (withTime) {
        var dateCode = this.getYear();
        if (this.getMonth() + 1 < 10) dateCode = dateCode + '0' + (this.getMonth() + 1);
        else dateCode = dateCode + '' + (this.getMonth() + 1);
        if (this.getDate() < 10) dateCode = dateCode + '0' + this.getDate();
        else dateCode = dateCode + '' + this.getDate();
        if (withTime) {
            if (this.getHours() < 10) dateCode = dateCode + '0' + this.getHours();
            else dateCode = dateCode + '' + this.getHours();
            if (this.getMinutes() < 10) dateCode = dateCode + '0' + this.getMinutes();
            else dateCode = dateCode + '' + this.getMinutes();
        }
        return dateCode;
    }
});

Object.assign(Date, {
    getDateCode: function (withTime) {
        return (new Date()).toDateCode(withTime);
    }
});
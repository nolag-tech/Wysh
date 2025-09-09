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
    left: function (len) {
        if (len <= 0) return '';
        if (len >= this.length) return this;
        return this.substring(0, len);
    },

    right: function (len) {
        if (len <= 0) return '';
        if (len >= this.length) return this;
        return this.substring(this.length - len, this.length);
    },

    repeat: function (times) {
        var s = '';
        for (var i = 0; i < times; i++) s += this;
        return s;
    },

    startsWith: function (str) {
        return (this.left(str.length) == str);
    },

    endsWith: function (str) {
        return (this.right(str.length) == str);
    },

    leftPad: function (char, width) {
        if (!char) char = ' ';
        return (char.repeat(width - this.length) + this).right(width);
    },

    rightPad: function (char, width) {
        if (!char) char = ' ';
        return (this + char.repeat(width - this.length)).left(width);
    },

    trim: function () {
        return this.replace(/^\s+|\s+$/g, "");
    },

    clean: function () {
        return this.replace(/\s{2,}/g, ' ').trim();
    },

    reverse: function () {
        return this.split('').reverse().join('');
    },

    toInt: function () {
        return parseInt(this);
    },

    toFloat: function () {
        return parseFloat(this);
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
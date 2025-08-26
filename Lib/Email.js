class Email {
    constructor(params, emailConf) {
        this.buffer = new StreamBuffer();
        this.message = new CDOMessage();
        this.config = new CDOConfiguration();

        this.emailConf = emailConf;

        for (key in emailConf) {
            config.Fields.Item(`http://schemas.microsoft.com/cdo/configuration/${key}`) = decode(emailConf[key]);
        }

        config.Fields.Update();
        message.Configuration = config;
        message.Fields.Update();

        this.options = {
            app: "Reporting",
            subject: "Untitled",
            html: false
        };

        Object.assign(this.options, params);

        this.buffer.Open();
    }

    WriteLine(str) {
        this.buffer.WriteText(str, 1);
    }

    Send() {
        this.buffer.Position = 0;

        if (!this.options.html) this.message.TextBody = this.buffer.ReadText(-1);
        else this.message.HTMLBody = this.buffer.ReadText(-1);

        this.message.Subject = `[${this.options.app.toUpperCase()}] ${this.options.subject}`;
        this.message.From = this.emailConf.smtpemailaddress;
        if (this.options.hasOwnProperty("replyto")) this.message.ReplyTo = this.options.replyto;
        this.message.To = this.options.mailto;

        if (this.options.bcc) this.message.Bcc = this.options.bcc;
        if (this.options.cc) this.message.Cc = this.options.cc;

        this.message.Send();
        this.buffer.Close();
    }

    AddAttachment(fname) {
        this.message.AddAttachment(fname);
    }

    decode(x) {
        return (('' + x).substr(0, 5) == 'data:') ? atob(('' + x).replace('data:', '')) : x;
    }
}
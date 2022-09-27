class MainConfig {
    constructor() {
        this.email = false;
        this.headered = [];

        this.i_spreadsheet = 0;
        this.i_tab = 0;
        this.k_column = 0;

        this.advanced = false;
        this.value_maps = [];
        this.custom_variables = [];
    }
}

class RecipientsConfig {
    constructor() {
        this.to = {
            enabled: false,
            is_template: false,
            col: 0,
            template_string: '',
            delim: ''
        }

        this.cc = {
            enabled: false,
            is_template: false,
            col: 0,
            template_string: '',
            delim: ''
        },

        this.bcc = {
            enabled: false,
            is_template: false,
            col: 0,
            template_string: '',
            delim: ''
        }
    }
}

class AttachmentsConfig {
    constructor() {
        this.archive = null;
        this.archive_name = '';
        this.conditional = false;
        this.filename_template = '';
        this.filename_expression = '';
    }
}

class EmailGenerateConfig {
    constructor() {
        this.filename_template = '';
        this.subject_template = '';
    }
}

class HtmlGenerateConfig {
    constructor() {
        this.output_type = 0;
        this.filename_template = '';
    }
}

class Config {
    constructor() {
        this.main = new MainConfig();
        this.recipients = new RecipientsConfig();
        this.attachments = new AttachmentsConfig();
        this.email_generate = new EmailGenerateConfig();
        this.html_generate = new HtmlGenerateConfig();
    }
}

export { MainConfig, RecipientsConfig, AttachmentsConfig, EmailGenerateConfig, HtmlGenerateConfig, Config }
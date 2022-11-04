<script>
import JSZip from 'jszip';
import OxMsg from './../../oxmsg.js';
import { Config } from './../../Configs.js';
let { Email, Attachment } = OxMsg;

export default {
    props: {
        options: Config,
        spreadsheets: Array
    },

    data: () => {
        return {
            progress_label: '',
            progress_percent: 0,
            generate_index: -1,
            errors: []
        }
    },

    methods: {
        flag_error(error, source) {
            if (error) {
                let lines = error.stack.split('\n');
                let exception = lines[0];
                let line = parseInt(lines[1].split('>')[lines[1].split('>').length - 1].split(')')[0].slice(1).split(':')[0]) - 2;
                if (!line) { line = null }
                this.errors.push({ source, line, exception });
            }

            else { this.errors.push({ source, line: null, exception: null }) }
        },

        async errored_sequence(fns) {
            let args = undefined;
            for(let f = 0; f < fns.length; f++) {
                let out = fns[f].constructor.name === 'AsyncFunction' ? await fns[f](args) : fns[f](args);
                if (this.errors.length > 0 && this.options.email_generate.break_on_error) { return false; }
                args = out;
            }

            return true;
        },

        errored_sequence_promise(fn) {
            return new Promise(async (resolve, reject) => {
                let out = fn.constructor.name === 'AsyncFunction' ? await fn() : fn();
                if (this.errors.length > 0 && this.options.email_generate.break_on_error) { reject() }
                else { resolve(out) }
            });
        },

        get_context(index) {
            let headered = this.options.main.headered;
            let spreadsheets = this.spreadsheets;
            let i_spreadsheet = this.options.main.i_spreadsheet;
            let i_tab = this.options.main.i_tab;
            let k_column = this.options.main.k_column;
            let row = spreadsheets[i_spreadsheet].sheets[i_tab].aoa[index];

            let ctx = {
                key: row[k_column],

                // shallow clone (shaking off object references)
                col: [...row],
                cols: [...row],
                spreadsheets: spreadsheets.map(s => s.sheets.map(t => t.aoa))
            }

            if (this.options.main.advanced) {
                ctx.mapped = {};

                let v_maps = this.options.main.value_maps;
                for (let v = 0; v < v_maps.length; v++) {
                    let m_spreadsheet = v_maps[v][0];
                    let m_tab = v_maps[v][1];
                    let k_col = v_maps[v][2];
                    let v_col = v_maps[v][3];
                    let overwrite = v_maps[v][4];
                    let s_col = v_maps[v][5];

                    let source = row[s_col];
                    let m_aoa = spreadsheets[m_spreadsheet].sheets[m_tab].aoa;

                    for (let m = headered[m_spreadsheet][m_tab] ? 1 : 0; m < m_aoa.length; m++) {
                        let key = m_aoa[m][k_col];
                        let value = m_aoa[m][v_col];

                        if (source == key) {
                            ctx.mapped[s_col] = value;
                            if (overwrite) { ctx.col[s_col] = value }
                            break;
                        }
                    }

                    if (ctx.mapped[s_col] == undefined) { ctx.mapped[s_col] = '' }
                }

                // static context; ref for custom vars
                let s_ctx = JSON.parse(JSON.stringify(ctx));

                let arr = this.options.main.custom_variables;
                ctx.vars = {};

                for (let a = 0; a < arr.length; a++) {
                    let name = arr[a][0];
                    let expression = arr[a][1];
                    try {
                        let return_val = elevated_exec(s_ctx, expression);
                        ctx.vars[name] = return_val;
                    }

                    catch(error) {
                        this.flag_error(error, `custom variable: ${name}`);
                        return;
                    }
                }
            }

            return ctx;
        },

        async get_email_buffer(index) {
            let key = this.spreadsheets[this.options.main.i_spreadsheet].sheets[this.options.main.i_tab].aoa[index][this.options.main.k_column];

            let ctx = this.get_context(index);
            let subject = `${madlib(ctx, this.options.email_generate.subject_template)}`;

            let email = new Email(true);
            email.subject(subject);

            let options = this.options.recipients;
            let props = [ 'to', 'cc', 'bcc' ];

            for (let p = 0; p < props.length; p++) {
                let k = props[p];
                let obj = options[k];

                if (obj.enabled) {
                    try {
                        let delim = obj.delim;
                        let value;
                        if (obj.is_template) { value = madlib(ctx, obj.template_string) }
                        else { value = ctx.col[obj.col] }
                    
                        let split = delim && value.includes(delim) ? value.split(delim) : [value];
                        for (let s = 0; s < split.length; s++) { email[k](split[s]) }
                    }

                    catch(error) { this.flag_error(error, `${key} ${props[p]} recipient`) }
                }
            }

            if (this.options.attachments.archive) {

                let data = this.options.attachments.archive;
                let { conditional, filename_template, filename_expression, error_on_blank } = this.options.attachments;

                try {
                    let val = '';
                    if (conditional) { val = elevated_exec(ctx, filename_expression) }
                    else { val = madlib(ctx, filename_template) }

                    let file = data.file(val);
                    if (file) { email.attach(new Attachment(await file.async('uint8array'), val)) }
                    else if (error_on_blank) { this.flag_error(null, `${key} attachment "${val}" not found`) }
                }

                catch(error) { this.flag_error(error, `${key} attachment`) }      
            }

            let html_element = document.createElement('html');
            html_element.innerHTML = madlib(ctx, tinymce.activeEditor.getContent());
            let images = html_element.getElementsByTagName('img');
            for (let i = 0; i < images.length; i++) {
                let src = images[i].getAttribute('src');
                let type = src.split('/')[1].split(';')[0];
                let base64 = src.split(',')[1];
                let uint8a = Uint8Array.from(atob(base64), (c) => c.charCodeAt(0));
                let attachment = new Attachment(uint8a, `img_${i}.${type}`, `img_${i}`);
                images[i].setAttribute('src', `cid:img_${i}`);
                email.attach(attachment);
            }

            email.bodyHtml(`
                <style>
                    p, ul li, ol li {
                        font-family: Calibri;
                        font-size: 14.5px;
                    }
                </style>
                ${html_element.innerHTML}`
            );

            return email.msg();
        },

        async start_generate() {
            this.progress_percent = 0;
            this.progress_label = '';
            this.errors = [];
            if (this.generate_index > -1) {
                await this.errored_sequence([
                    () => { return this.get_email_buffer(this.generate_index) },
                    (buffer) => { return { buffer, filename: `${madlib(this.get_context(this.generate_index), this.options.email_generate.filename_template).replace(/\\[?%*:|"<>]/g, '_')}.msg` } },
                    ({ filename, buffer }) => { trigger_download(buffer, filename) }
                
                ]);
            }

            else {
                let zip = new JSZip();

                let headered = this.options.main.headered;
                let spreadsheets = this.spreadsheets;
                let i_spreadsheet = this.options.main.i_spreadsheet;
                let i_tab = this.options.main.i_tab;
                let aoa = spreadsheets[i_spreadsheet].sheets[i_tab].aoa;

                let i = (headered[i_spreadsheet][i_tab] ? 1 : 0) - 1;

                for (i; i < aoa.length; i++) {
                    let current_errors = this.errors.length;

                    let success = await this.errored_sequence([
                        () => {
                            let filename = `${madlib(this.get_context(i), this.options.email_generate.filename_template).replace(/[\\?%*:|"<>]/g, '_')}.msg`;
                            
                            this.progress_percent = Math.ceil((i / aoa.length) * 100);
                            this.progress_label = `${filename}...`;
                        
                            return filename;
                        },

                        async (filename) => { return { filename, buffer: await this.get_email_buffer(i) } },
                        async ({ filename, buffer }) => {
                            if (!(this.options.email_generate.skip_on_error && this.errors.length > current_errors)) { zip.file(filename, buffer) }
                        }
                    ]);

                    if (!success && this.options.email_generate.break_on_error) {
                        this.progress_label = 'failed...';
                        break;
                    }
                }

                if ((this.errors.length > 0 && !this.options.email_generate.break_on_error) || this.errors.length == 0) {
                    this.progress_label = 'zipping...';
                    zip.generateAsync({ type: 'uint8array' }).then((u8) => {
                        this.progress_label = 'done.';
                        trigger_download(u8, `export_${Date.now()}.zip`);
                    })
                }
            }
        }
    },

    computed: {
        keys() {
            if (this.spreadsheets.length == 0) { return [] }
            let { i_spreadsheet, i_tab, k_column } = this.options.main;

            let aoa = this.spreadsheets[i_spreadsheet].sheets[i_tab].aoa;
            let arr = [];
            for(let i = 0; i < aoa.length; i++) { arr.push([i, aoa[i][k_column]]) }

            return arr;
        }
    }
}
</script>

<script setup>
    import Widget from '../Widget.vue';
</script>

<template>
    <widget title="Generate">
        <table class="table table-borderless" style="margin-top: 10px;">
            <thead>
                <tr>
                    <th scope="col">Filename Template</th>
                    <th scope="col">Subject Template</th>
                    <th scope="col">Break on Error</th>
                    <th scope="col">Skip on Error</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td>
                        <div class="input-group">
                            <input type="text" class="form-control" placeholder="<filename template>" v-model="options.email_generate.filename_template">
                            <div class="input-group-append"><span class="input-group-text">.msg</span></div>
                        </div>
                    </td>
                    <td>
                        <div class="input-group">
                            <input type="text" class="form-control" placeholder="<subject template>" v-model="options.email_generate.subject_template">
                        </div>
                    </td>
                    <td>
                        <div class="form-check form-switch">
                            <input type="checkbox" class="form-check-input" v-model="options.email_generate.break_on_error">
                        </div>
                    </td>
                    <td>
                        <div class="form-check form-switch">
                            <input type="checkbox" class="form-check-input" v-model="options.email_generate.skip_on_error">
                        </div>
                    </td>
                </tr>
            </tbody>
        </table>
        <div style="margin: 0 5px;">
            <div class="input-group" style="margin-top: 5px;">
                <select class="form-control" :disabled="spreadsheets.length == 0" v-model="generate_index">
                    <template v-if="spreadsheets.length > 0">
                        <option :value="-1">-- ALL --</option>
                        <option v-for="v in keys" :key="v[0]" :value="v[0]">{{v[1]}}</option>
                    </template>
                </select>
                <button type="button" class="btn btn-secondary" @click="start_generate()" :disabled="spreadsheets.length == 0 || !options.email_generate.filename_template">Generate</button>
            </div>
            <div class="progress-wrapper">
                <div class="progress-bg" :style="`width: ${progress_percent}%;`">
                    <div class="progress-label">{{progress_label}}</div>
                </div>
            </div>
            <div v-if="errors.length > 0">
                <div v-for="(error, i) in errors" :key="i" class="error">
                    <p class="error-heading">{{error.source}}</p>
                    <p class="error-body" v-if="error.exception">{{error.exception}}{{error.line ? ` at line ${error.line}` : ''}}</p>
                </div>
            </div>
        </div>
    </widget>
</template>

<style scoped>
    .progress-wrapper {
        display: relative;
        width: 100%;
        height: 20px;
        background-color: #dee2e6;
        margin: 10px 0 5px 0;
    }

    .progress-bg {
        background-color: #202020;
        display: relative;
        height: 100%;
        transition: all 0.1s;

    }

    .progress-label {
        position: absolute;
        color: #FFFFFF;
        margin: 0 auto;
        line-height: 20px;
        width: 100%;
        left: 0;
        text-align: center;
        mix-blend-mode: difference;
    }

    .error {
        font-family: monospace;
        font-weight: bold;
        margin: 10px 0 5px 0;
        background-color: #FFBBBB;
        color: #BB0000;
        padding: 5px 10px;
        border-radius: 5px;
        border: solid 2px #FF7777;
        font-size: 14px;
    }

    .error-heading, .error-body {
        margin: 0;
        padding: 0;
    }

    .error-heading {

    }

    .error-body {
    }
</style>
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
        }
    },

    methods: {
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
                    let return_val = elevated_exec(s_ctx, expression);
                    ctx.vars[name] = return_val;
                }
            }

            return ctx;
        },

        async get_email_buffer(index) {
            let ctx = this.get_context(index);
            let subject = `${madlib(ctx, this.options.email_generate.subject_template)}`;

            let { to, cc } = this.options.recipients;

            let email = new Email(true);
            email.subject(subject);

            if (to.enabled) {
                let split = to.delim ? ctx.col[to.col].split(to.delim) : [ctx.col[to.col]];
                for (let s = 0; s < split.length; s++) { email.to(split[s]) }
            }

            if (cc.enabled) {
                let split = cc.delim ? ctx.col[cc.col].split(cc.delim) : [ctx.col[cc.col]];
                for (let s = 0; s < split.length; s++) { email.cc(split[s]) }
            }

            if (this.options.attachments.archive) {
                let data = this.options.attachments.archive;
                let { conditional, filename_template, filename_expression } = this.options.attachments;

                try {
                    let val = '';
                    if (conditional) { val = elevated_exec(ctx, filename_expression) }
                    else { val = madlib(ctx, filename_template) }
                    let file = data.file(val);
                    if (file) { email.attach(new Attachment(await file.async('uint8array'), val)) }
                }

                catch(e) { /* abort? */ }
                
                
            }

            let html_element = document.createElement('html');
            html_element.innerHTML = tinymce.activeEditor.getContent();
            let images = html_element.getElementsByTagName('img');
            for (let i = 0; i < images.length; i++) {
                let src = images[i].getAttribute('src');
                let type = src.split('/')[1].split(';')[0];
                let base64 = src.split(',')[1];
                let uint8a = Uint8Array.from(atob(base64), (c) => c.charCodeAt(0));
                let attachment = new Attachment(uint8a, `img_${i}.${type}`, `img_${i}`);
                images[i].src = `cid:img_${i}`;
                email.attach(attachment);
            }

            email.bodyHtml(`
                <style>
                    p, ul li, ol li {
                        font-family: Calibri;
                        font-size: 14.5px;
                    }
                </style>
                ${madlib(ctx, html_element.innerHTML)}`
            );

            return email.msg();
        },

        async start_generate() {
            this.progress_percent = 0;
            this.progress_label = '';
            if (this.generate_index > -1) {
                let buffer = await this.get_email_buffer(this.generate_index);
                let filename = `${madlib(this.get_context(this.generate_index), this.options.email_generate.filename_template).replace(/[/\\?%*:|"<>]/g, '_')}.msg`;
                trigger_download(buffer, filename);
            }

            else {
                let zip = new JSZip();

                let headered = this.options.main.headered;
                let spreadsheets = this.spreadsheets;
                let i_spreadsheet = this.options.main.i_spreadsheet;
                let i_tab = this.options.main.i_tab;
                let aoa = spreadsheets[i_spreadsheet].sheets[i_tab].aoa;
            
                for (let i = headered[i_spreadsheet][i_tab] ? 1 : 0; i < aoa.length; i++) {
                    let filename = `${madlib(this.get_context(i), this.options.email_generate.filename_template).replace(/[/\\?%*:|"<>]/g, '_')}.msg`;

                    this.progress_percent = (i / aoa.length) * 100;
                    this.progress_label = `${filename}...`;

                    let buffer = await this.get_email_buffer(i);

                    zip.file(filename, buffer);
                }

                this.progress_label = 'zipping...';

                zip.generateAsync({ type: 'uint8array' }).then((u8) => {
                    this.progress_label = 'done.';
                    trigger_download(u8, `export_${Date.now()}.zip`);
                });
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
            <label style="margin-top: 20px;">{{progress_label}}</label>
            <div class="progress" style="margin: 5px 0; border-radius: 0px;">
                <div class="progress-bar progress-bar-striped progress-bar-animated bg-secondary" role="progressbar" :style="`width: ${progress_percent}%;`"></div>
            </div>
        </div>
    </widget>
</template>

<style scoped>
th {
    font-size: 1rem;
    font-weight: bold;
    padding: 5px;
}
</style>
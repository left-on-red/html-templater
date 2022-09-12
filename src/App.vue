<script setup>
    import SpreadsheetWidget from './components/widgets/SpreadsheetWidget.vue';
    import TemplateWidget from './components/widgets/TemplateWidget.vue';
    import AttachmentsWidget from './components/widgets/AttachmentsWidget.vue';
    import RecipientsWidget from './components/widgets/RecipientsWidget.vue';
    import GenerateEmailsWidget from './components/widgets/GenerateEmailsWidget.vue';
    import GenerateHtmlWidget from './components/widgets/GenerateHtmlWidget.vue';
    import Docs from './Docs.vue';
    import config from './../package.json';
</script>

<script>
    import JSZip from 'jszip';

    export default {
        data: () => {
            return {
                template_extension: 'template',
                docs: false,
                spreadsheets: [],
                main_options: {
                    email_mode: false,
                    headered: [],
                    i_spreadsheet: 0,
                    i_tab: 0,
                    k_column: 0,
                    value_maps: [],
                    custom_variables: [],
                    advanced: false
                },

                recipient_options: {
                    enable_to: false,
                    to_col: 0,
                    to_delim: '',
                    enable_cc: false,
                    cc_col: 0,
                    cc_delim: ''
                },

                attachments_options: {
                    file: null,
                    conditional: false,
                    flag: false,
                    filename_template: '',
                    filename_expression: ''
                },

                email_generate_options: {
                    filename_template: '',
                    subject_template: ''
                },

                html_generate_options: {
                    output_type: 0,
                    filename_template: ''
                }
            }
        },

        methods: {
            async export_config() {
                let zip = new JSZip();

                // kind of annoyed that I can't destructure back into an object
                let { main_options, recipient_options, email_generate_options, html_generate_options } = this;

                let attachments_options = {
                    conditional: this.attachments_options.conditional,
                    flag: this.attachments_options.flag,
                    filename_template: this.attachments_options.filename_template,
                    filename_expression: this.attachments_options.filename_expression,
                    archive_filename: this.attachments_options.file ? this.attachments_options.file.name : null
                }

                let obj = { main_options, recipient_options, email_generate_options, html_generate_options, attachments_options };
                for (let o in obj) { zip.file(`${o}.json`, JSON.stringify(obj[o], null, 4)) }
        
                if (this.attachments_options.file) { zip.file(this.attachments_options.file.name, await this.attachments_options.file.data.generateAsync({ type: 'uint8array' })) }

                let sheet_index = [];
                for (let s = 0; s < this.spreadsheets.length; s++) {
                    zip.folder('spreadsheets').file(this.spreadsheets[s].name, this.spreadsheets[s].buffer);
                    sheet_index.push(this.spreadsheets[s].name);
                }

                zip.file('spreadsheets.json', JSON.stringify(sheet_index, null, 4));
                zip.file('template.html', tinymce.activeEditor.getContent());

                let blob = new Blob([await zip.generateAsync({ type: 'uint8array' })]);
                let anchor = document.createElement('a');
                anchor.href = window.URL.createObjectURL(blob);
                anchor.download = `config_${Date.now()}.${this.template_extension}`;
                anchor.click();
            },

            async import_config() {
                let upload = document.createElement('input');
                upload.type = 'file';
                upload.accept = `.${this.template_extension}`;
                upload.addEventListener('change', async (event) => {
                    let file = event.target.files[0];
                    let buffer = await file.arrayBuffer();

                    let zip = await JSZip.loadAsync(buffer);

                    let recur = (base, src) => {
                        for (let s in src) {
                            if (typeof base[s] == 'object' && base[s] != null) {
                                if (base[s] instanceof Array) {
                                    while (base[s].length > 0) { base[s].pop() }
                                    for (let i = 0; i < src[s].length; i++) { base[s].push(src[s][i]) }
                                }

                                else { recur(base[s], src[s]) }
                            }

                            else { base[s] = src[s] }
                        }
                    }

                    let json_files = [ 'main_options', 'recipient_options', 'attachments_options', 'email_generate_options', 'html_generate_options' ];
                    for (let j = 0; j < json_files.length; j++) {
                        try {
                            let obj = JSON.parse(await zip.file(`${json_files[j]}.json`).async('string'));
                            if (json_files[j] == 'attachments_options') {
                                let filename = obj.archive_filename;
                                delete obj.archive_filename;
                                obj.file = {
                                    name: filename,
                                    data: await JSZip.loadAsync(await zip.file(filename).async('uint8array'))
                                }

                                this.attachments_options.file = {
                                    name: '',
                                    data: null
                                }
                            }

                            recur(this[json_files[j]], obj);
                        }
                        
                        catch(e) {}
                    }

                    //tinymce.activeEditor.getContent()
                    tinymce.activeEditor.setContent(await zip.file('template.html').async('string'));

                    while (this.spreadsheets.length > 0) { this.spreadsheets.pop() }
                    let spreadsheet_index = JSON.parse(await zip.file('spreadsheets.json').async('string'));
                
                    for (let s = 0; s < spreadsheet_index.length; s++) {
                        let spreadsheet_buffer = await zip.file(`spreadsheets/${spreadsheet_index[s]}`).async('uint8array');
                        let spreadsheet = XLSX.read(spreadsheet_buffer);
                        let names = spreadsheet.SheetNames;

                        let obj = {
                            name: spreadsheet_index[s],
                            buffer: spreadsheet_buffer,
                            sheets: []
                        }

                        for (let n = 0; n < names.length; n++) {
                            let aoa = XLSX.utils.sheet_to_csv(spreadsheet.Sheets[names[n]], { FS: '\t' }).split('\n').map(v => v.split('\t')).filter(v => v.join('') != '');
                            obj.sheets.push({ name: names[n], aoa });
                        }

                        this.spreadsheets.push(obj);
                    }
                });

                upload.click();
            }
        }
    }
</script>

<template>
    <div class="container" style="margin-bottom: 100px; margin-top: 60px;">
        <h1 style="padding: 0; margin: 20px 0 0 0;">
            <span :class="main_options.email_mode ? 'text-danger' : 'text-primary'">
                {{main_options.email_mode ? 'EMAIL' : 'HTML'}}
            </span>
            TEMPLATER
            <span class="version">v{{config.version}}</span>
            <div class="form-check form-switch d-inline-block" style="vertical-align: top;">
                <input type="checkbox" class="form-check-input" v-model="main_options.email_mode" :disabled="docs">
            </div>
            <div class="d-inline-block float-end" style="margin-left: 5px;">
                <button type="button" class="btn btn-outline-secondary" @click="docs = !docs" :style="docs ? 'background-color: #6c757d !important; color: #fff !important;' : ''">Docs</button>
            </div>
            <div class="d-inline-block float-end">
                <button type="button" class="btn btn-outline-secondary" @click="export_config()" :disabled="docs">Export Config</button>
            </div>
            <div class="d-inline-block float-end" style="margin-right: 5px;">
                <button type="button" class="btn btn-outline-secondary" @click="import_config()" :disabled="docs">Import Config</button>
            </div>
        </h1>
        <Docs v-if="docs" style="margin-top: 20px;" />
        <template v-else>
            <SpreadsheetWidget :spreadsheets="spreadsheets" :options="main_options" />
            <TemplateWidget />
            <AttachmentsWidget v-if="main_options.email_mode" :options="attachments_options" />
            <RecipientsWidget v-if="main_options.email_mode" :options="recipient_options" :aoa="spreadsheets[main_options.i_spreadsheet] ? spreadsheets[main_options.i_spreadsheet].sheets[main_options.i_tab].aoa : []" :headered="main_options.headered[main_options.i_spreadsheet] ? main_options.headered[main_options.i_spreadsheet][main_options.i_tab] : false" />
            <GenerateEmailsWidget v-if="main_options.email_mode" :options="{ spreadsheets, main_options, recipient_options, attachments_options, html_generate_options, email_generate_options }" />
            <GenerateHtmlWidget v-else :options="{ spreadsheets, main_options, recipient_options, attachments_options, html_generate_options, email_generate_options }" />
        </template>
    </div>
</template>

<style scoped>
    input[type="checkbox"]:focus {
        box-shadow: none !important;
        background-image: url("data:image/svg+xml,%3csvg xmlns='http://www.w3.org/2000/svg' viewBox='-4 -4 8 8'%3e%3ccircle r='3' fill='%23fff'/%3e%3c/svg%3e") !important;
    }

    input[type="checkbox"]:checked {
        background-color: #dc3545 !important;
        border-color: #dc3545 !important;
    }

    input[type="checkbox"] {
        background-image: url("data:image/svg+xml,%3csvg xmlns='http://www.w3.org/2000/svg' viewBox='-4 -4 8 8'%3e%3ccircle r='3' fill='%23fff'/%3e%3c/svg%3e") !important;
        background-color: #007bff !important;
        border-color: #007bff !important;
        transition: all 0.2s;
    }

    span {
        display: inline-block;
        width: 150px;
        transition: all 0.2s;
        text-align: center;
    }

    span.version {
        font-size: 24px;
        vertical-align: top;
        width: initial;
        padding: 0 10px;
        letter-spacing: 3px;
        background-color: #404040;
        margin: 10px 15px 10px -10px;
        color: #F1F1F1;
        border-radius: 50px;
    }
</style>

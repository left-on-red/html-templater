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

    import { Config } from './Configs.js';

    export default {
        data: () => {
            return {
                template_extension: 'template',
                docs: false,
                spreadsheets: [],
                // this property CANNOT be named "config"
                // seems to be a reserved Vue prop
                options: new Config()
            }
        },

        computed: {
            main_aoa() {
                //spreadsheets[config.main.i_spreadsheet] ? spreadsheets[config.main.i_spreadsheet].sheets[config.main.i_tab].aoa : []
                let { i_spreadsheet, i_tab } = this.options.main;
                return this.spreadsheets[i_spreadsheet] ? this.spreadsheets[i_spreadsheet].sheets[i_tab].aoa : [];
            },

            main_headered() {
                //main_options.headered[main_options.i_spreadsheet] ? main_options.headered[main_options.i_spreadsheet][main_options.i_tab] : false
                let { i_spreadsheet, i_tab, headered } = this.options.main;
                return headered[i_spreadsheet] ? headered[i_spreadsheet][i_tab] : false;
            }
        },

        methods: {
            async export_config() {
                let zip = new JSZip();

                let configs = Object.keys(this.options);
                for (let c = 0; c < configs.length; c++) {
                    let obj = this.options[configs[c]];
                    // don't think I can serialize the file data w/ json. not that it would be a good idea anyway. still annoying tho
                    if (configs[c] == 'attachments') {
                        let clone = Object.assign({}, obj);
                        if (clone.archive) { zip.file(clone.archive_name, await clone.archive.generateAsync({ type: 'uint8array' })) }

                        delete clone.archive;
                        obj = clone;
                    }

                    zip.file(`${configs[c]}.json`, JSON.stringify(obj, null, 4))
                }

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

                    let configs = Object.keys(this.options);
                    for (let c = 0; c < configs.length; c++) {
                        try {
                            let obj = JSON.parse(await zip.file(`${configs[c]}.json`).async('string'));
                            recur(this.options[configs[c]], obj);
                        }

                        catch {}

                        if (this.options.attachments.archive_name) {
                            let archive_name = this.options.attachments.archive_name;
                            try { this.options.attachments.archive = await JSZip.loadAsync(await zip.file(archive_name).async('uint8array')) }
                            catch {
                                this.options.attachments.archive = null;
                                this.options.attachments.archive_name = '';
                            }
                        }
                    }
                    
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
        <div class="heading">
            <h1>TEMPLATER</h1>
            <span>v{{config.version}}</span>
            <div class="reactives">
                <div :class="`mode-btn${options.main.email ? ' active' : ''}`" @click="options.main.email = !options.main.email">{{options.main.email ? 'EMAIL' : 'HTML'}}</div>
                <div :class="`docs-btn${docs ? ' active' : ''}`" @click="docs = !docs">DOCS</div>
            </div>
        </div>
        <hr />
        <Docs v-if="docs" style="margin-top: 20px;" />
        <template v-else>
            <div class="import-export">
                <div @click="import_config()">IMPORT CONFIG</div>
                <div @click="export_config()">EXPORT CONFIG</div>
            </div>
            <SpreadsheetWidget :spreadsheets="spreadsheets" :options="options.main" />
            <TemplateWidget />
            <AttachmentsWidget v-if="options.main.email" :options="options.attachments" />
            <RecipientsWidget v-if="options.main.email" :options="options.recipients" :aoa="main_aoa" :headered="main_headered" />
            <GenerateEmailsWidget v-if="options.main.email" :options="options" :spreadsheets="spreadsheets" />
            <GenerateHtmlWidget v-else :options="options" :spreadsheets="spreadsheets" />
        </template>
    </div>
</template>

<style scoped>
    .heading {
        margin: 0;
        padding: 0;
        display: block;
        display: flex;
    }
    .heading h1 {
        display: inline-block;
        font-size: 24px;
        font-weight: bold;
        letter-spacing: 4px;
        font-family: monospace;
        padding: 0;
        margin: 0 !important;
        line-height: 10px;
        padding: 10px 0;
    }
    .heading span {
        display: inline-block;
        background-color: #303030;
        color: #F1F1F1;
        font-size: 12px;
        font-family: monospace;
        font-weight: bold;
        border-radius: 25px;
        line-height: 10px;
        margin: 6px 0;
        padding: 4px 7px;
    }
    .reactives {
        margin-left: auto;
    }

    .reactives div {
        font-size: 20px;
        letter-spacing: 2px;
        font-weight: bold;
        font-family: monospace !important;
        transition: all 0.2s !important;
        width: 100px;
        height: 100%;
        text-align: center;
        cursor: pointer;
        display: inline-block;
    }

    .reactives .mode-btn {
        background-color: #007bff;
        border: solid 1px #007bff;
        color: #FFF;
    }

    .reactives .mode-btn.active {
        background-color: #dc3545;
        border-color: #dc3545;
    }

    .reactives .docs-btn {
        font-weight: normal;
        color: #6c757d;
        border: solid 1px #6c757d;
    }

    .reactives .docs-btn.active {
        background-color: #6c757d;
        color: #FFF;
    }

    .import-export div {
        font-size: 14px;
        font-family: monospace !important;
        letter-spacing: 1px;
        transition: all 0.2s;
        height: 100%;
        width: 150px;
        text-align: center;
        cursor: pointer;
        display: inline-block;
        margin: 0 5px 0 0;
        padding: 5px 0;
        background-color: #6c757d;
        color: #FFF;
    }

    .import-export div:hover {
        background-color: #5c636a;
    }
</style>

<script>
import JSZip from 'jszip';

export default {
    props: {
        options: Object
    },

    data: () => {
        return {
            progress_label: '',
            progress_percent: 0,
        }
    },

    methods: {
        get_context(index) {
            let headered = this.options.main_options.headered;
            let spreadsheets = this.options.spreadsheets;
            let i_spreadsheet = this.options.main_options.i_spreadsheet;
            let i_tab = this.options.main_options.i_tab;
            let k_column = this.options.main_options.k_column;
            let row = spreadsheets[i_spreadsheet].sheets[i_tab].aoa[index];

            let ctx = {
                key: row[k_column],

                // shallow clone (shaking off object references)
                col: [...row],
                spreadsheets: spreadsheets.map(s => s.sheets.map(t => t.aoa))
            }

            if (this.options.main_options.advanced) {
                ctx.mapped = {};

                let v_maps = this.options.main_options.value_maps;
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

                let arr = this.options.main_options.custom_variables;
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

        get_html_text(index) {
            return new Promise((resolve) => {
                setTimeout(() => {
                    let ctx = this.get_context(index);
                    resolve(madlib(ctx, tinymce.activeEditor.getContent()));
                });
            });
        },

        async start_generate() {
            this.progress_percent = 0;
            this.progress_label = '';

            let type = this.options.html_generate_options.output_type;

            let headered = this.options.main_options.headered;
            let spreadsheets = this.options.spreadsheets;
            let i_spreadsheet = this.options.main_options.i_spreadsheet;
            let i_tab = this.options.main_options.i_tab;
            let k_column = this.options.main_options.k_column;
            let aoa = spreadsheets[i_spreadsheet].sheets[i_tab].aoa;
            
            // spreadsheet
            if (type == 4) {
                let computed = [];
                for (let i = headered[i_spreadsheet][i_tab] ? 1 : 0; i < aoa.length; i++) {
                    let key = aoa[i][k_column];

                    this.progress_percent = (i / aoa.length) * 100;
                    this.progress_label = `${key}...`;

                    let value = madlib(this.get_context(i), await this.get_html_text(i));
                    computed.push([key, value]);
                }

                let workbook = XLSX.utils.book_new();
                let worksheet = XLSX.utils.aoa_to_sheet(computed);
                XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
                XLSX.writeFile(workbook, `export_${Date.now()}.xlsx`, { type: 'array' });

                this.progress_label = `done.`;
            }

            else {
                let many = [0, 1].includes(type) ? true : false;
                let extension = [1, 3].includes(type) ? 'html.txt' : 'html';

                if (many) {
                    let zip = new JSZip();
                    
                    for (let i = headered[i_spreadsheet][i_tab] ? 1 : 0; i < aoa.length; i++) {
                        let filename = `${madlib(this.get_context(i), this.options.html_generate_options.filename_template).replace(/[/\\?%*:|"<>]/g, '_')}.${extension}`;
                        
                        this.progress_percent = (i / aoa.length) * 100;
                        this.progress_label = `${filename}...`;
                        
                        zip.file(filename, await this.get_html_text(i));
                    }

                    this.progress_label = 'zipping...';

                    zip.generateAsync({ type: 'uint8array' }).then((u8) => {
                        this.progress_label = 'done.';
                        trigger_download(u8, `export_${Date.now()}.zip`);
                    });
                }

                else {
                    let files = [];

                    for (let i = headered[i_spreadsheet][i_tab] ? 1 : 0; i < aoa.length; i++) {
                        this.progress_percent = (i / aoa.length) * 100;
                        this.progress_label = `${aoa[i][k_column]}...`;

                        files.push(await this.get_html_text(i));
                    }

                    trigger_download((new TextEncoder()).encode(files.join('\n\n')), `export_${Date.now()}.${extension}`);
                    this.progress_label = 'done.';
                }
            }
        }
    },

    computed: {
        keys() {
            if (this.options.spreadsheets.length == 0) { return [] }
            let { i_spreadsheet, i_tab, k_column } = this.options.main_options;

            let aoa = this.options.spreadsheets[i_spreadsheet].sheets[i_tab].aoa;
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
                    <th scope="col">Output</th>
                    <th scope="col">Filename Template</th>
                    <th scope="col"></th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td>
                        <div class="input-group">
                            <select class="form-control" :disabled="options.spreadsheets.length == 0" v-model="options.html_generate_options.output_type">
                                <option :value="0">Many Files (.html)</option>
                                <option :value="1">Many Files (.html.txt)</option>
                                <option :value="2">One File (.html)</option>
                                <option :value="3">One File (.html.txt)</option>
                                <option :value="4">Spreadsheet</option>
                            </select>
                        </div>
                    </td>
                    <td>
                        <div class="input-group">
                            <input type="text" class="form-control" placeholder="<filename template>" :disabled="[2, 3, 4].includes(options.html_generate_options.output_type)" v-model="options.html_generate_options.filename_template">
                            <div class="input-group-append"><span class="input-group-text">{{ options.html_generate_options.output_type == 1 ? '.html.txt' : '.html' }}</span></div>
                        </div>
                    </td>
                    <td>
                        <button type="button" class="btn btn-secondary" @click="start_generate()" :disabled="options.spreadsheets.length == 0 || ([0, 1].includes(options.html_generate_options.output_type) && options.html_generate_options.filename_template == '')">Generate</button>
                    </td>
                </tr>
            </tbody>
        </table>
        <label style="margin-top: 20px;">{{progress_label}}</label>
        <div class="progress" style="margin: 5px 0; border-radius: 0px;">
            <div class="progress-bar progress-bar-striped progress-bar-animated bg-secondary" role="progressbar" :style="`width: ${progress_percent}%;`"></div>
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
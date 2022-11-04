<script>
import JSZip from 'jszip';
import { AttachmentsConfig } from './../../Configs.js';

export default {
    props: {
        spreadsheets: Array,
        options: AttachmentsConfig
    },
   
    methods: {
        async upload() {
            let upload = document.createElement('input');
            upload.type = 'file';
            upload.accept = '.zip,*';
            upload.addEventListener('change', async (event) => {
                let file = event.target.files[0];
                if (file.name.endsWith('.zip')) {
                    JSZip.loadAsync(await file.arrayBuffer()).then((data) => {
                        this.options.archive = data;
                        this.options.archive_name = file.name;
                    });
                }
                
                else {
                    this.options.archive = await file.arrayBuffer();
                    this.options.archive_name = file.name;
                }
            });

            upload.click();
        }
    }
}
</script>

<script setup>
    import Widget from '../Widget.vue';
    import CodeEditor from './../../components/CodeEditor.vue';
</script>

<template>
    <widget title="Attachments">
        <table class="table table-borderless" style="margin-top: 10px;">
            <thead>
                <tr>
                    <th scope="col">Zip Archive</th>
                    <th scope="col">Conditional</th>
                    <th scope="col">Error on None</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td>
                        <div class="input-group">
                            <button type="button" class="btn btn-outline-secondary" @click="upload()">Upload</button>
                            <input type="text" class="form-control" placeholder="<no file specified>" readonly :value="options.archive ? options.archive_name : ''">
                        </div>
                    </td>
                    <td>
                        <div class="form-check form-switch">
                            <input type="checkbox" class="form-check-input" v-model="options.conditional" :disabled="options.archive == null" />
                        </div>
                    </td>
                    <td>
                        <div class="form-check form-switch">
                            <input type="checkbox" class="form-check-input" v-model="options.error_on_blank" :disabled="options.archive == null" />
                        </div>
                    </td>
                </tr>
            </tbody>
        </table>
        <div v-if="options.archive && options.archive_name.endsWith('.zip')" style="padding: 5px;">
            <div v-if="options.conditional" style="width: 100%; height: 200px;">
                <CodeEditor v-model="options.filename_expression" />
            </div>
            <div v-else class="input-group">
                <div class="input-group-prepend"><span class="input-group-text">Filename Template</span></div>
                <input type="text" class="form-control" placeholder="<template>" v-model="options.filename_template">
            </div>
        </div>
    </widget>
</template>

<style scoped>
</style>
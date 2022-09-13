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
    import AttachmentsArchive from '../controls/attachments/AttachmentsArchive.vue';
</script>

<template>
    <widget title="Attachments">
        <div class="input-group">
            <button type="button" class="btn btn-outline-secondary" @click="upload()">Upload</button>
            <input type="text" class="form-control" placeholder="<no file specified>" readonly :value="options.archive ? options.archive_name : ''">
        </div>
        <AttachmentsArchive v-if="options.archive && options.archive_name.endsWith('.zip')" :options="options" />
    </widget>
</template>

<style scoped>

</style>
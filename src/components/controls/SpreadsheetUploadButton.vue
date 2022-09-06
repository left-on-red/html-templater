<script>
export default {
    methods: {
        async upload() {
            let upload = document.createElement('input');
            upload.type = 'file';
            upload.accept = '.xlsx,.xls,.wks,.txt,.csv';
            upload.addEventListener('change', async (event) => {
                let file = event.target.files[0];
                let buffer = await file.arrayBuffer();
                let spreadsheet = XLSX.read(buffer);
                let names = spreadsheet.SheetNames;

                let obj = {
                    name: file.name,
                    buffer,
                    sheets: []
                }

                for (let n = 0; n < names.length; n++) {
                    let aoa = XLSX.utils.sheet_to_csv(spreadsheet.Sheets[names[n]], { FS: '\t' }).split('\n').map(v => v.split('\t')).filter(v => v.join('') != '');
                    obj.sheets.push({ name: names[n], aoa });
                }

                this.$emit('upload', obj);
            });

            upload.click();
        },
    }
}
</script>

<template>
    <button type="button" class="btn btn-outline-secondary" @click="upload()">Upload</button>
</template>
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
                    let workbook = spreadsheet.Sheets[names[n]];
                    let aoa = [];
                    let range = XLSX.utils.decode_range(workbook['!ref']);

                    let row_count = range.e.r + 1;
                    let col_count = range.e.c + 1;

                    for (let r = 0; r < row_count; r++) {
                        let row = [];
                        for (let c = 0; c < col_count; c++) {
                            let index = XLSX.utils.encode_cell({ r, c });

                            let content = '';
                            try { content = workbook[index].v }
                            catch {}

                            row.push(content);
                        }

                        aoa.push(row);
                    }

                    //let aoa = XLSX.utils.sheet_to_csv(spreadsheet.Sheets[names[n]], { FS: '\t' }).split('\n').map(v => v.split('\t')).filter(v => v.join('') != '');
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
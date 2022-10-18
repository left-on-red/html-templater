<script>
    import { MainConfig } from './../../Configs.js';
export default {
    props: {
        spreadsheets: Array,
        options: MainConfig
    },

    watch: {
        spreadsheets: {
            deep: true,
            handler(arr) { 
                if (arr.length == this.options.headered.length) { return }
                if (arr.length == 0) {
                    while (this.options.headered.length) { this.options.headered.pop() }
                    return;
                }
                
                this.options.headered.push(arr[arr.length - 1].sheets.map(v => false)) }
        }
    }
}
</script>

<script setup>
    import Widget from '../Widget.vue';
    import TabPicker from '../controls/TabPicker.vue';
    import ColumnPicker from '../controls/ColumnPicker.vue';
    import SpreadsheetUploadPicker from '../controls/SpreadsheetUploadPicker.vue';
    import ValueMapsTable from '../controls/ValueMapsTable.vue';
    import VariablesTable from '../controls/VariablesTable.vue';
    import CodeEditor from '../CodeEditor.vue';
</script>

<template>
    <widget title="Data Source">
        <table class="table table-borderless" style="margin-top: 10px;">
            <thead>
                <tr>
                    <th scope="col">Iterative Spreadsheet</th>
                    <th scope="col">Unique Key</th>
                    <th scope="col">Headered</th>
                    <th scope="col">Advanced Options</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td>
                        <div class="input-group">
                            <SpreadsheetUploadPicker :spreadsheets="spreadsheets" v-model="options.i_spreadsheet" />
                            <TabPicker :spreadsheet="spreadsheets[options.i_spreadsheet]" v-model="options.i_tab" />
                        </div>
                    </td>
                    <td>
                        <div class="input-group">
                            <ColumnPicker :sheet="spreadsheets[options.i_spreadsheet] ? spreadsheets[options.i_spreadsheet].sheets[options.i_tab] ? spreadsheets[options.i_spreadsheet].sheets[options.i_tab].aoa : [] : []" v-model="options.k_column" :headered="options.headered[options.i_spreadsheet] ? options.headered[options.i_spreadsheet][options.i_tab] : false" />
                        </div>
                    </td>
                    <td>
                        <div class="input-group">
                            <div class="form-check form-switch">
                                <input v-if="options.headered[options.i_spreadsheet]" type="checkbox" class="form-check-input" v-model="options.headered[options.i_spreadsheet][options.i_tab]">
                                <input v-else type="checkbox" class="form-check-input" disabled>
                            </div>
                        </div>
                    </td>
                    <td>
                        <div class="input-group">
                            <div class="form-check form-switch">
                                <input type="checkbox" class="form-check-input" :disabled="!options.headered[options.i_spreadsheet]" v-model="options.advanced">
                            </div>
                        </div>
                    </td>
                </tr>
            </tbody>
        </table>
        <ValueMapsTable v-if="options.advanced" :spreadsheets="spreadsheets" :headered="options.headered" :table="options.value_maps" :main_spreadsheet="options.i_spreadsheet" :main_tab="options.i_tab" />
        <VariablesTable v-if="options.advanced" :table="options.custom_variables" />
    </widget>
</template>

<style scoped>
th {
    font-size: 1rem;
    font-weight: bold;
    padding: 5px;
}
</style>
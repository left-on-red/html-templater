<script>
export default {
    props: {
        spreadsheets: Array,
        values: Array,
        headered: Array,
        main_spreadsheet: Number,
        main_tab: Number
    },

    watch: {
        values: {
            deep: true,
            handler() { this.$emit('update', this.values) }
        }
    }
}
</script>
<script setup>
import SpreadsheetUploadPicker from './SpreadsheetUploadPicker.vue';
import TabPicker from './TabPicker.vue';
import ColumnPicker from './ColumnPicker.vue';
</script>

<template>
    <tr>
        <td>
            <div class="input-group">
                <SpreadsheetUploadPicker :spreadsheets="spreadsheets" v-model="values[0]" />
                <TabPicker :spreadsheet="spreadsheets[values[0]]" v-model="values[1]" />
            </div>
        </td>
        <td>
            <div class="input-group">
                <ColumnPicker :sheet="spreadsheets[values[0]] ? spreadsheets[values[0]].sheets[values[1]].aoa: []" :headered="headered[values[0]][values[1]]" v-model="values[2]" />
            </div>
        </td>
        <td>
            <div class="input-group">
                <ColumnPicker :sheet="spreadsheets[values[0]] ? spreadsheets[values[0]].sheets[values[1]].aoa: []" :headered="headered[values[0]][values[1]]" v-model="values[3]" />
            </div>
        </td>
        <td>
            <div class="input-group">
                <div class="form-check form-switch">
                    <input v-if="headered[values[0]]" type="checkbox" class="form-check-input" v-model="headered[values[0]][values[1]]">
                    <input v-else type="checkbox" class="form-check-input" disabled>
                </div>
            </div>
        </td>
        <td>
            <div class="input-group">
                <div class="form-check form-switch">
                    <input type="checkbox" class="form-check-input" v-model="values[4]">
                </div>
            </div>
        </td>
        <td>
            <div class="input-group">
                <ColumnPicker :sheet="spreadsheets[main_spreadsheet] ? spreadsheets[main_spreadsheet].sheets[main_tab].aoa : []" :headered="headered[main_spreadsheet][main_tab]" v-model="values[5]" />
            </div>
        </td>
        <td>
            <button type="button" class="btn-close" @click="$emit('destroy')"></button>
        </td>
    </tr>
</template>

<style scoped>
    tr td {
        padding: 0 !important;
    }
    tr .input-group .form-switch {
        margin-left: 10px;
        margin-top: 3px;
    }

    tr .btn-close {
        padding: 10px;
    }
</style>
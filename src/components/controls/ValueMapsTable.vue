<script>
export default {
    props: {
        spreadsheets: Array,
        headered: Array,
        table: Array,
        main_spreadsheet: Number,
        main_tab: Number
    }
}
</script>
<script setup>
import ValueMapsRow from './ValueMapsRow.vue';
</script>

<template>
    <h6>Key Value Maps</h6>
    <table class="table table-bordered">
        <thead>
            <tr>
                <th scope="col">Spreadsheet</th>
                <th scope="col">Key</th>
                <th scope="col">Value</th>
                <th scope="col">Headered</th>
                <th scope="col">Overwrite</th>
                <th scope="col" colspan="2">Source</th>
            </tr>
        </thead>
        <tbody>
            <ValueMapsRow v-for="(v, i) in table" :key="i" :values="table[i]" :spreadsheets="spreadsheets" :headered="headered" :main_spreadsheet="main_spreadsheet" :main_tab="main_tab" @destroy="table.splice(i, 1)" />
            <tr>
                <td colspan="7" style="padding: 0 !important;"><button type="button" class="btn btn-outline-secondary" @click="table.push([0, 0, 0, 0, false, 0]);" :disabled="spreadsheets.length == 0">Add New Mapping</button></td>
            </tr>
        </tbody>
    </table>
</template>

<style scoped>
    button {
        width: 100%;
    }
</style>
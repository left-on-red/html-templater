<script>
export default {
    props: {
        aoa: Array,
        options: Object,
        headered: Boolean
    }
}
</script>

<script setup>
    import Widget from '../Widget.vue';
    import ColumnPicker from '../controls/ColumnPicker.vue';
</script>

<template>
    <widget title="Recipients">
        <table class="table table-borderless" style="margin-top: 10px;">
            <thead>
                <tr>
                    <th scope="col"></th>
                    <th scope="col">Enable</th>
                    <th scope="col">Delimiter</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    
                    <td>
                        <div class="input-group">
                            <span class="input-group-text">To</span>
                            <ColumnPicker :sheet="aoa" :headered="headered" :disabled="!options.enable_to" v-model="options.to_col" />
                        </div>
                    </td>

                    <td>
                        <div class="form-check form-switch mx-2">
                            <input type="checkbox" class="form-check-input" v-model="options.enable_to" :disabled="aoa.length == 0">
                        </div>
                    </td>

                    <td>
                        <input class="form-control" style="width: 100px;" v-model="options.to_delim">
                    </td>
                </tr>

                <tr>
                    <td>
                        <div class="input-group">
                            <span class="input-group-text">CC</span>
                            <ColumnPicker :sheet="aoa" :headered="headered" :disabled="!options.enable_cc" v-model="options.cc_col" />
                        </div>
                    </td>

                    <td>
                        <div class="form-check form-switch mx-2">
                            <input type="checkbox" class="form-check-input" v-model="options.enable_cc" :disabled="aoa.length == 0">
                        </div>
                    </td>

                    <td>
                        <input class="form-control" style="width: 100px;" v-model="options.cc_delim">
                    </td>
                </tr>
            </tbody>
        </table>
    </widget>
</template>

<style scoped>
th {
    font-size: 1rem;
    font-weight: bold;
    padding: 5px;
}
</style>
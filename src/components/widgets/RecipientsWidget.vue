<script>
import { RecipientsConfig } from './../../Configs.js';
export default {
    props: {
        aoa: Array,
        options: RecipientsConfig,
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
                            <ColumnPicker :sheet="aoa" :headered="headered" :disabled="!options.to.enabled" v-model="options.to.col" />
                        </div>
                    </td>

                    <td>
                        <div class="form-check form-switch mx-2">
                            <input type="checkbox" class="form-check-input" v-model="options.to.enabled" :disabled="aoa.length == 0">
                        </div>
                    </td>

                    <td>
                        <input class="form-control" style="width: 100px;" v-model="options.to.delim">
                    </td>
                </tr>

                <tr>
                    <td>
                        <div class="input-group">
                            <span class="input-group-text">CC</span>
                            <ColumnPicker :sheet="aoa" :headered="headered" :disabled="!options.cc.enabled" v-model="options.cc.col" />
                        </div>
                    </td>

                    <td>
                        <div class="form-check form-switch mx-2">
                            <input type="checkbox" class="form-check-input" v-model="options.cc.enabled" :disabled="aoa.length == 0">
                        </div>
                    </td>

                    <td>
                        <input class="form-control" style="width: 100px;" v-model="options.cc.delim">
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
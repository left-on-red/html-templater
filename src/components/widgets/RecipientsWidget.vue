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
                    <th scope="col">Templated</th>
                    <th scope="col">Error on Empty</th>
                    <th scope="col">Delimiter</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td>
                        <div class="input-group">
                            <span class="input-group-text">To</span>
                            <input v-if="options.to.is_template" class="form-control" v-model="options.to.template_string" placeholder="template">
                            <ColumnPicker v-else :sheet="aoa" :headered="headered" :disabled="!options.to.enabled" v-model="options.to.col" />
                        </div>
                    </td>

                    <td>
                        <div class="form-check form-switch mx-2">
                            <input type="checkbox" class="form-check-input" v-model="options.to.enabled" :disabled="aoa.length == 0">
                        </div>
                    </td>

                    <td>
                        <div class="form-check form-switch mx-2">
                            <input type="checkbox" class="form-check-input" v-model="options.to.is_template">
                        </div>
                    </td>

                    <td>
                        <div class="form-check form-switch mx-2">
                            <input type="checkbox" class="form-check-input" v-model="options.to.error_on_blank">
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
                            <input v-if="options.cc.is_template" class="form-control" v-model="options.cc.template_string" placeholder="template">
                            <ColumnPicker v-else :sheet="aoa" :headered="headered" :disabled="!options.cc.enabled" v-model="options.cc.col" />
                        </div>
                    </td>

                    <td>
                        <div class="form-check form-switch mx-2">
                            <input type="checkbox" class="form-check-input" v-model="options.cc.enabled" :disabled="aoa.length == 0">
                        </div>
                    </td>

                    <td>
                        <div class="form-check form-switch mx-2">
                            <input type="checkbox" class="form-check-input" v-model="options.cc.is_template">
                        </div>
                    </td>

                    <td>
                        <div class="form-check form-switch mx-2">
                            <input type="checkbox" class="form-check-input" v-model="options.cc.error_on_blank">
                        </div>
                    </td>

                    <td>
                        <input class="form-control" style="width: 100px;" v-model="options.cc.delim">
                    </td>
                </tr>

                <tr>
                    <td>
                        <div class="input-group">
                            <span class="input-group-text">BCC</span>
                            <input v-if="options.bcc.is_template" class="form-control" v-model="options.bcc.template_string" placeholder="template">
                            <ColumnPicker v-else :sheet="aoa" :headered="headered" :disabled="!options.bcc.enabled" v-model="options.bcc.col" />
                        </div>
                    </td>

                    <td>
                        <div class="form-check form-switch mx-2">
                            <input type="checkbox" class="form-check-input" v-model="options.bcc.enabled" :disabled="aoa.length == 0">
                        </div>
                    </td>

                    <td>
                        <div class="form-check form-switch mx-2">
                            <input type="checkbox" class="form-check-input" v-model="options.bcc.is_template">
                        </div>
                    </td>

                    <td>
                        <div class="form-check form-switch mx-2">
                            <input type="checkbox" class="form-check-input" v-model="options.bcc.error_on_blank">
                        </div>
                    </td>

                    <td>
                        <input class="form-control" style="width: 100px;" v-model="options.bcc.delim">
                    </td>
                </tr>
            </tbody>
        </table>
    </widget>
</template>

<style scoped>

.input-group input, .input-group select {
    width: 300px !important;
}
</style>
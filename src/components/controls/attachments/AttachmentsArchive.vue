<script>
export default {
    props: {
        spreadsheets: Array,
        options: Object
    }
}
</script>

<script setup>

</script>

<template>
    <div class="attachments-archive">
        <h6>ZIP Archive</h6>
        <div class="container">
            <div class="row mb-2">
                <div class="col">
                    <div class="input-group">
                        <div class="form-check form-switch">
                            <input type="checkbox" class="form-check-input" id="zip_expressions_checkbox" v-model="options.conditional" />
                            <label for="zip_expressions_checkbox" class="form-check-label">Conditional Expression</label>
                        </div>
                        <!--<div class="form-check form-switch mx-5">
                            <input type="checkbox" class="form-check-input" id="zip_expressions_flag" v-model="options.flag" />
                            <label for="zip_expressions_flag" class="form-check-label">Flag if not found?</label>
                        </div>-->
                    </div>
                </div>
            </div>
            <div class="row">
                <div class="col">
                    <textarea v-if="options.conditional" class="form-control" rows="3" placeholder="Javascript Expression..." v-model="options.filename_expression"></textarea>
                    <div v-else class="input-group">
                        <div class="input-group-prepend"><span class="input-group-text">Filename Template</span></div>
                        <input type="text" class="form-control" placeholder="<template>" v-model="options.filename_template">
                    </div>
                </div>
            </div>
        </div>
    </div>
</template>

<style scoped>
    .attachments-archive {
        margin-top: 20px;
    }

    label {
        margin: 4px 0 0 5px;
    }
</style>
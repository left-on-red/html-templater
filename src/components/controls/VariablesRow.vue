<script>
export default {
    props: {
        values: Array
    },

    watch: {
        values: {
            deep: true,
            handler() { this.$emit('update', this.values) }
        }
    }
}
</script>

<template>
    <div class="row">
        <div class="col-12" style="padding: 0 0 5px 0;">
            <div class="input-group">
                <div class="input-group-prepend"><span class="input-group-text">Variable Name</span></div>
                <input type="text" class="form-control" v-model="values[0]">
                <button type="button" class="btn-close" @click="$emit('destroy')"></button>
            </div>
        </div>
        <div class="col-12" style="padding: 0 0 10px 0;">
            <textarea class="form-control" rows="10" placeholder="Javascript Expression..." data-gramm="false" v-model="values[1]"></textarea>
        </div>
    </div>
    </template>

<style>
    tr .input-group .form-switch {
        margin-left: 10px;
        margin-top: 3px;
    }

    .btn-close {
        padding: 10px;
        margin-left: 5px !important;
    }
</style>
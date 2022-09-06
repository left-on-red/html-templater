<script>
export default {
    props: {
        spreadsheet: Object,
        modelValue: Number
    },

    computed: {
        sheets() { return this.spreadsheet ? this.spreadsheet.sheets.map(v => v.name) : [] }
    },

    watch: {
        spreadsheet() {
            //this.$emit('update:modelValue', 0);
        }
    }
}
</script>
<script setup>
</script>

<template>
    <select class="form-select" v-model="modelValue" :disabled="sheets.length == 0" @change="$emit('update:modelValue', parseInt($event.target.value))">
        <option v-for="(s, i) in sheets" :key="i" :value="i">{{s}}</option>
    </select>
</template>

<style scoped>
</style>
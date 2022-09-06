<script>
export default {
    props: {
        spreadsheets: Array,
        modelValue: Number
    }
}
</script>

<script setup>
    import SpreadsheetPicker from './SpreadsheetPicker.vue';
    import SpreadsheetUploadButton from './SpreadsheetUploadButton.vue';
</script>

<template>
    <SpreadsheetUploadButton @upload="spreadsheets.push($event); $emit('update:modelValue', spreadsheets.length - 1);" />
    <SpreadsheetPicker :spreadsheets="spreadsheets" v-model="modelValue" @update:modelValue="$emit('update:modelValue', $event)" />
</template>

<style scoped>
</style>
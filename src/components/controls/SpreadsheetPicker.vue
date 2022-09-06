<script>
export default {
    props: {
        spreadsheets: Array,
        disabled: Boolean,
        modelValue: Number
    }
}
</script>
<script setup>
</script>

<template>
    <select v-model="modelValue" class="form-select" :disabled="disabled || spreadsheets.length == 0" @change="$emit('update:modelValue', parseInt($event.target.value))">
        <option v-for="(s, i) in spreadsheets" :key="i" :value="i">{{s.name}}</option>
    </select>
</template>

<style scoped>
</style>
<script>
export default {
    props: {
        sheet: Array,
        headered: Boolean,
        modelValue: Number
    },

    watch: {
        sheet() {
            if (!this.sheet[0] || !this.sheet[0][this.modelValue]) { this.$emit('update:modelValue', 0) }
        }
    },

    methods: {
        // MIT - https://github.com/avilaton/excel-column-name
        int_to_col(number) {
            let colName = '';
            let dividend = Math.floor(Math.abs(number));
            let rest;
        
            while (dividend > 0) {
                rest = (dividend - 1) % 26;
                colName = String.fromCharCode(65 + rest) + colName;
                dividend = parseInt((dividend - rest)/26);
            }

            return colName;
        }
    }
}
</script>
<script setup>
</script>

<template>
    <select class="form-select" v-model="modelValue" :disabled="sheet.length == 0" @change="$emit('update:modelValue', parseInt($event.target.value))">
        <option v-if="sheet[0]" v-for="(s, i) in sheet[0]" :key="i" :value="i">
            {{headered ? s : `Column ${int_to_col(i+1)}`}}
        </option>
    </select>
</template>

<style scoped>
</style>
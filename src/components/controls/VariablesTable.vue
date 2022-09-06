<script>
export default {
    props: {
        table: Array
    }
}
</script>
<script setup>
import VariablesRow from './VariablesRow.vue';
</script>

<template>
    <h6>Custom Variables</h6>
    <div class="container">
        <VariablesRow v-for="(v, i) in table" :key="i" :values="table[i]" @destroy="table.splice(i, 1)" />
        <div class="row">
            <div class="col-12" style="padding: 0;">
                <button type="button" class="btn btn-outline-secondary" @click="table.push([`variable${table.length + 1}`, ''])">Add New Variable</button>
            </div>
        </div>
    </div>
</template>

<style scoped>
    button {
        width: 100%;
    }
</style>
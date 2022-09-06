<script>
    export default {
        props: {
            title: String
        }
    }
</script>

<template>
  <div class="widget">
    <h2 v-if="title">{{title}}</h2>
    <div class="widget-content">
        <slot></slot>
    </div>
  </div>
</template>

<style scoped>
.widget {
  /* background-color: #00000005;
  border: solid 2px #0000001F; */
  padding: 0;
  margin: 0 0 20px 0;
  color: #202020;
  border: solid 1px #dee2e6;
}

.widget .widget-content {
    padding: 5px;
    margin: 0;
}

.widget h2 {
    padding: 5px;
    margin: 0;
    background-color: #dee2e6;
}
</style>

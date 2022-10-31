<script>
    import loader from '@monaco-editor/loader';

    export default {
        props: {
            modelValue: String
        },

        mounted() {
            loader.init().then((monaco) => {
                let options = {
                    language: 'javascript',
                    minimap: { enabled: false },
                    automaticLayout: true
                }

                monaco.languages.typescript.javascriptDefaults.addExtraLib(`
                    /** the key value of the current iterative item */
                    declare let key : string;

                    /** an array representing the column values within the current iterative item */
                    declare let cols : string[];

                    /** a key-value mapped object where the key is the column number (zero-indexed) and the value is the mapped value */
                    declare let mapped : Object;

                    /** a 4D array representing all of the \`spreadsheets -> sheets -> rows -> columns\` that have currently been uploaded */
                    declare let spreadsheets : string[][][][];

                `, 'htmltemplater.d.ts');

                let editor = monaco.editor.create(this.$refs.code_editor, options);

                editor.setValue(this.modelValue);
                editor.onDidChangeModelContent(() => { this.$emit('update:modelValue', editor.getValue()) })
            });
        }
    }
</script>

<template>
    <div ref="code_editor" style="width: 100%; height: 100%; border: solid 1px #dee2e6;"></div>
</template>

<style scoped>
</style>

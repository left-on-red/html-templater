<script>
export default {
    props: {},

    mounted() {
        tinymce.init({
            //mode: mode,
            //editor_selector: selector,
            selector: '#editor',
            //elements: elements,
            //skin: 'lightgray',
            //width: 'auto',
            plugins: [
                "advlist autolink link image lists charmap preview hr anchor colorpicker importcss",
                "searchreplace visualblocks visualchars code codemirror fullscreen insertdatetime media nonbreaking",
                "table contextmenu directionality emoticons paste textcolor",
                "noneditable media contentblock"
            ],
            paste_data_images: true,
            forced_root_block: 'p',
            codemirror: {					// From http://www.avoid.org/codemirror-for-tinymce4/
                indentOnInit: true, 		// Whether or not to indent code on init.
                config: {					// CodeMirror config object
                    indentWithTabs: true,	// Turn on use of tabs instead of spaces
                    lineWrapping: false		// Enable line wrapping of long lines
                },
                path: 'CodeMirror'			// Path to CodeMirror distribution
            },
            style_formats: [
                {
                    title: 'Inline', items: [
                        { title: 'Bold', icon: 'bold', format: 'bold' },
                        { title: 'Italic', icon: 'italic', format: 'italic' },
                        { title: 'Underline', icon: 'underline', format: 'underline' },
                        { title: 'Strikethrough', icon: 'strikethrough', format: 'strikethrough' },
                        //{ title: 'Superscript', icon: 'superscript', format: 'superscript' },
                        //{ title: 'Subscript', icon: 'subscript', format: 'subscript' },
                        //{ title: 'Code', icon: 'code', format: 'code' }
                    ]
                },
                /*{
                    title: 'Blocks', items: [
                        { title: 'Paragraph', format: 'p' },
                        { title: 'Div', format: 'div' },
                        { title: 'Blockquote', format: 'blockquote' },
                        { title: 'Pre', format: 'pre' }
                    ]
                },
                {
                    title: 'Headings', items: [
                        { title: 'Heading 1', format: 'h1' },
                        { title: 'Heading 2', format: 'h2' },
                        { title: 'Heading 3', format: 'h3' },
                        { title: 'Heading 4', format: 'h4' },
                        { title: 'Heading 5', format: 'h5' },
                        { title: 'Heading 6', format: 'h6' }
                    ]
                },*/
                {
                    title: 'Alignment', items: [
                        { title: 'Left', icon: 'alignleft', format: 'alignleft' },
                        { title: 'Center', icon: 'aligncenter', format: 'aligncenter' },
                        { title: 'Right', icon: 'alignright', format: 'alignright' },
                        { title: 'Justify', icon: 'alignjustify', format: 'alignjustify' },
                    ]
                },
                {
                    title: 'Table Alignment', items: [
                        { title: 'Align Top', icon: 'valigntop', format: 'valigntop' },
                        { title: 'Align Middle', icon: 'valignmiddle', format: 'valignmiddle' },
                        { title: 'Align Bottom', icon: 'valignbottom', format: 'valignbottom' }
                    ]
                },
            ],
            //content_css: "/Content/admin/css/TinyMCE4.css",
            //body_id: "main",
            importcss_append: true,
            content_style: "",
            media_live_embeds: true,
            extended_valid_elements: "img[id|style|class|src|border=0|alt|title|hspace|vspace|width|height|align|onmouseover|onmouseout|name|data-caption|ismap],contentblock[name|style|class|title],span[*],@[xmlns],svg[*],use[*]",
            //link_list: "/find/AdminJSON/LinkList4",
            //contentblock_list: "/find/AdminJSON/ContentBlockList4",
            target_list: [
                { title: 'None', value: '' },
                { title: 'New tab/window', value: '_blank' },
            ],
            valid_children: "+a[div|data|table|ul|h1|h2|h3|h4|h5|h6|p|#text]",
            menu: {
                edit: { title: 'Edit', items: 'undo redo | cut copy paste pastetext | selectall | searchreplace' },
                view: { title: 'View', items: 'visualaid visualchars visualblocks code' },
                //insert: { title: 'Insert', items: 'link image media anchor | charmap emoticons | template hr | contentblock' },
                insert: { title: 'Insert', items: 'link image variable | charmap' },
                //format: { title: 'Format', items: 'bold italic underline strikethrough superscript subscript | formats fontsizeselect | removeformat' },
                format: { title: 'Format', items: 'bold italic underline strikethrough | removeformat' },
                //table: { title: 'Table', items: 'inserttable tableprops deletetable | cell row column' },
                //tools: { title: 'Tools', items: 'code fullscreen ' }
            },
            //toolbar1: "undo redo  restoredraft | bold italic styleselect removeformat | alignleft aligncenter alignright alignjustify | bullist numlist outdent indent | link unlink image media fullpage | forecolor backcolor | code preview",
            toolbar1: 'undo redo restoredraft | bold italic underline strikethrough removeformat | bullist numlist outdent indent | link unlink image | forecolor backcolor | code preview',
            toolbar2: "",
            toolbar3: "",
            //contextmenu: "cut copy paste pastetext selectall searchreplace | bold italic removeformat | link anchor ",
            contextmenu: 'cut copy paste pastetext | bold italic underline strikethrough removeformat | link unlink',
            convert_fonts_to_spans: true,
            image_advtab: true,
            relative_urls: false,
            remove_script_host: true,
            document_base_url: "",
            browser_spellcheck: true,
            gecko_spellcheck: true,
            setup: function (ed) {
                ed.on('init', function (evt) {
                    ed.getBody().setAttribute('spellcheck', true);
                });
            }
        });
    }
}
</script>

<script setup>
    import Widget from '../Widget.vue';
</script>

<template>
    <widget title="Template">
        <div id="editor"></div>
    </widget>
</template>

<style scoped>

</style>
<script setup>
</script>

<script>

    export default {
    }
</script>

<template>
    <div class="row">
        <div class="col d-lg-none"></div>
        <div class="col d-none d-lg-block">
            <ul>
                <li>
                    <a href="#about">About</a>
                    <ul>
                        <li><a href="#what_can_it_do">What Can It Do?</a></li>
                        <li><a href="#outlook_email_draft_files">Outlook Email Draft Files</a></li>
                    </ul>
                </li>
                <li>
                    <a href="#data_source">Data Source</a>
                    <ul>
                        <li><a href="#iterative_spreadsheet">Iterative Spreadsheet</a></li>
                        <li><a href="#unique_key">Unique Key</a></li>
                        <li><a href="#headered">Headered</a></li>
                        <li><a href="#advanced_options">Advanced Options</a></li>
                    </ul>
                </li>
                <li>
                    <a href="#key_value_maps">Key Value Maps</a>
                    <ul>
                        <li><a href="#key_value_maps_spreadsheet">Spreadsheet</a></li>
                        <li><a href="#key_value_maps_key">Key</a></li>
                        <li><a href="#key_value_maps_value">Value</a></li>
                        <li><a href="#key_value_maps_headered">Headered</a></li>
                        <li><a href="#key_value_maps_overwrite">Overwrite</a></li>
                        <li><a href="#key_value_maps_source">Source</a></li>
                    </ul>
                </li>
                <li>
                    <a href="#custom_variables">Custom Variables</a>
                    <ul>
                        <li><a href="custom_variables_variable_name">Variable Name</a></li>
                        <li><a href="custom_variables_expression">Expression</a></li>
                    </ul>
                </li>
                <li><a href="#template">Template</a></li>
                <li>
                    <a href="#attachments">Attachments</a>
                    <ul>
                        <li><a href="#attachments_uploaded_file">Uploaded File</a></li>
                        <li><a href="#attachments_conditional_expression">Conditional Expression</a></li>
                        <li><a href="#attachments_filename_template">Filename Template</a></li>
                    </ul>
                </li>
                <li>
                    <a href="#recipients">Recipients</a>
                    <ul>
                        <li><a href="#recipients_to">To</a></li>
                        <li><a href="#recipients_cc">CC</a></li>
                        <li><a href="#recipients_enable">Enable</a></li>
                        <li><a href="#recipients_delimiter">Delimiter</a></li>
                    </ul>
                </li>
                <li>
                    <a href="#generate_email_templater">Generate Email</a>
                    <ul>
                        <li><a href="#generate_email_filename_template">Filename Template</a></li>
                        <li><a href="#generate_email_subject_template">Subject Template</a></li>
                        <li><a href="#generate_email_generate">Generate</a></li>
                    </ul>
                </li>
                <li>
                    <a href="#generate_html_templater">Generate HTML</a>
                    <ul>
                        <li><a href="#generate_html_output_mode">Output Mode</a></li>
                        <li><a href="#generate_html_filename_template">Filename Template</a></li>
                    </ul>
                </li>
                <li>
                    <a href="#custom_code">Custom Code</a>
                    <ul>
                        <li><a href="#javascript_expressions">JavaScript Expressions</a></li>
                        <li><a href="#templated_text_inputs">Templated Text Inputs</a></li>
                        <li><a href="#javascript_context">JavaScript Context</a></li>
                    </ul>
                </li>
            </ul>
        </div>
        <div class="col-9">
            <h2 id="about">About</h2>
            <h3 id="what_can_it_do">What Can It Do?</h3>
            <p>This Templater Tool can generate both <a href="#outlook_email_draft_files">Outlook Email Draft Files</a> and HTML content using data-oriented spreadsheets, user-defined templates, and magic.</p>
            <h3 id="outlook_email_draft_files">Outlook Email Draft Files</h3>
            <p>One of the key features of this tool is the ability to generate these <a href="#outlook_email_draft_files">Outlook Email Draft Files</a>. These files have .msg file extensions and can be opened directly in Outlook. The nature of these files allows us to programmatically populate the various fields of an email and export that all into a single file that can be immediately ready to send to the recipient.</p>
            
            <h2 id="data_source">Data Source</h2>
            <h3 id="iterative_spreadsheet">Iterative Spreadsheet</h3>
            <p>This determines the overall dataset that will be used to generate the defined HTML content. Every item is represented by 1 row, with a consistent column scheme across the entire spreadsheet. This spreadsheet input allows you to upload new spreadsheets and pick an individual sheet within to select as your iterative data</p>
            <h3 id="unique_key">Unique Key</h3>
            <p>This is a specific column that’s unique among all of the rows. i.e. a client name or similar. This is used elsewhere in the application when a unique key is required for unique identification</p>
            <h3 id="headered">Headered</h3>
            <p>A true/false toggle switch that determines whether or not the currently selected spreadsheet is headered (the first row is reserved for column headers.) Every single individual sheet within all uploaded spreadsheets can be individually headered/non-headered</p>
            <h3 id="advanced_options">Advanced Options</h3>
            <p>A true/false toggle switch that enables some of the more advanced templating features. This includes <a href="#key_value_maps">Key Value Maps</a> and <a href="#custom_variables">Custom Variables</a></p>
            <h2 id="key_value_maps">Key Value Maps (Advanced Options)</h2>
            <p>A series of “maps” tells the templater to interpret certain columns differently.</p>
            <h3 id="key_value_maps_spreadsheet">Spreadsheet</h3>
            <p>The spreadsheet that the current value map will use as a lookup table</p>
            <h3 id="key_value_maps_key">Key</h3>
            <p>The column that the current value map will use to fetch a specific value from the spreadsheet lookup table</p>
            <h3 id="key_value_maps_value">Value</h3>
            <p>The column that the current value map will use to map the key from the spreadsheet lookup table to a specific value</p>
            <h3 id="key_value_maps_headered">Headered</h3>
            <p>A true/false toggle switch that determines whether or not the currently selected spreadsheet is headered</p>
            <h3 id="key_value_maps_overwrite">Overwrite</h3>
            <p>A true/false toggle switch that tells the templater whether or not to actually overwrite the key value source with the fetched value. This behavior distance may have specific use cases if you use custom variables</p>
            <h3 id="key_value_maps_source">Source</h3>
            <p>The column from the <a href="#iterative_spreadsheet">Iterative Spreadsheet</a> that the current value map will use to coerce that value into what was found using the spreadsheet lookup table</p>
            
            <h2 id="custom_variables">Custom Variables (Advanced Options)</h2>
            <p>A series of custom variables that can be defined using <a href="#javascript_expressions">JavaScript Expressions</a>. These can be used within any <a href="#templated_text_inputs">Templated Text Input</a></p>
            <h3 id="custom_variables_variable_name">Variable Name</h3>
            <p>Used to identify the custom variable when it’s used in any other <a href="#javascript_expressions">JavaScript Expressions</a> or <a href="#templated_text_inputs">Templated Text Inputs</a></p>
            <h3 id="custom_variables_expression">Expression</h3>
            <p>A <a href="javascript_expression">JavaScript Expression</a> that defines a Custom Variable. The return value determines what your Custom Variable equals (calculated per item within the <a href="#iterative_spreadsheet">Iterative Spreadsheet</a>)</p>

            <h2 id="template">Template</h2>
            <p>This represents the HTML template that the templater is going to use for every item within the iterative spreadsheet. The content editor was ripped straight out of Genesis along with its config. Slight modifications have been made to make it work in this context. If you know how to use the Genesis content editor, you know how to use this one. This entire field is a <a href="#templated_text_inputs">Templated Text Input</a>.</p>
            
            <h2 id="attachments">Attachments (Email Templater)</h2>

            <h3 id="attachments_uploaded_file">Uploaded File</h3>
            <p>A zip archive that the templater can use to attach files to the various <a href="#outlook_email_draft_files">Outlook Email Draft Files</a>.</p>
            <h3 id="attachments_conditional_expression">Conditional Expression</h3>
            <p>A true/false toggle switch that determines whether or not the <a href="#attachments_filename_template">Filename Template</a> is a <a href="#javascript_expressions">JavaScript Expression</a> or a <a href="#templated_text_inputs">Templated Text Input</a>.</p>
            <h3 id="attachments_filename_template">Filename Template</h3>
            <p>If <a href="#attachments_conditional_expression">Conditional Expression</a> is set to true, this is <a href="#javascript_expressions">JavaScript Expression</a> where the return value would be used as the filename to look up the attach file from the zip archive. Otherwise, it’s a <a href="#templated_text_inputs">Templated Text Input</a> that represents the attach filename.</p>
            
            <h2 id="recipients">Recipients (Email Templater)</h2>
            <h3 id="recipients_to">To</h3>
            <p>The specified recipients that the generated <a href="#outlook_email_draft_files">Outlook Email Draft Files</a> will be configured to address the emails to. Represented as a specified column within the <a href="iterative_spreadsheet">Iterative Spreadsheet</a>.</p>
            <h3 id="recipients_cc">CC</h3>
            <p>The specified recipients that the generated <a href="#outlook_email_draft_files">Outlook Email Draft Files</a> will be configured to CC the emails to. Represented as a specified column within the <a href="iterative_spreadsheet">Iterative Spreadsheet</a>.</p>
            <h3 id="recipients_enable">Enable</h3>
            <p>A true/false toggle switch that determines whether or not that to enable this specific functionality</p>
            <h3 id="recipients_delimiter">Delimiter</h3>
            <p>A short text input that determines how to split the cell’s contents if more than 1 recipient is expected.</p>
            
            <h2 id="generate_email">Generate (Email Templater)</h2>
            <h3 id="generate_email_filename_template">Filename Template</h3>
            <p>A <a href="#templated_text_input">Templated Text Input</a> that determines the filenames of the <a href="#outlook_email_draft_files">Outlook Email Draft Files</a> that the templater will generate</p>
            <h3 id="generate_email_subject_template">Subject Template</h3>
            <p>A <a href="#templated_text_input">Templated Text Input</a> that determines the subjects of the <a href="#outlook_email_draft_files">Outlook Email Draft Files</a> that the templater will generate</p>
            <h3 id="generate_email_generate">Generate</h3>
            <p>A dropdown that allows you to pick a specific item in the <a href="#iterative_spreadsheet">Iterative Spreadsheet</a> that you want to generate. Select <b>"-- ALL --"</b> if you want to generate all items as a zip archive. Clicking the <b>"Generate"</b> button will initiate the Email generation process</p>
            
            <h2 id="generate_html">Generate (HTML Templater)</h2>
            <h3 id="generate_html_output_mode">Output Mode</h3>
            <p>A dropdown that allows you to specify how you want the generated HTML data to be exported.</p>
            <ul>
                <li>Specifying one of the <b>Single File</b> options will result in a single text/html file that includes all of the generated HTML in sequential order.</li>
                <li>Specifying one of the <b>Many Files</b> options will result in a zip archive that includes all of the generated HTML files using the <a href="#generate_html_filename_template">Filename Template</a></li>
                <li>Specifying the <b>Spreadsheet</b> option will result in a spreadsheet where every row will contain the Unique Key and the generated HTML content.</li>
            </ul>
            <h3 id="generate_html_filename_template">Filename Template</h3>
            <p>A <a href="#templated_text_inputs">Templated Text Input</a> that determines the filenames of the HTML files that the templater will generate. Only applicable if <a href="#output_mode">Output Mode</a> is set to one of the <b>Many Files</b> options.</p>
            
            <h2 id="custom_code">Custom Code</h2>
            <h3 id="javascript_expressions">JavaScript Expressions</h3>
            <p>Blocks of functional JavaScript code that can be used for templating purposes to greatly extend functionality. For a basic JavaScript cheatsheet that’s available online, you can refer to:</p>
            <ul>
                <li><a href="https://htmlcheatsheet.com/js/" target="_blank">https://htmlcheatsheet.com/js/</a></li>
                <li><a href="https://github.com/iLoveCodingOrg/javascript-cheatsheet/blob/master/js-cheatsheet.pdf" target="_blank">https://github.com/iLoveCodingOrg/javascript-cheatsheet/blob/master/js-cheatsheet.pdf</a></li>
            </ul>
            <h3 id="templated_text_inputs">Templated Text inputs</h3>
            <p>There are many <a href="#templated_text_inputs">Templated Text Inputs</a> where you can embed custom inline JavaScript code into text content. Any content within <b>{curly braces}</b> in <a href="#templated_text_inputs">Templated Text Inputs</a> will be interpreted as inline JavaScript.</p>
            <h3 id="javascript_context">JavaScript Context</h3>
            <p>When you’re executing JavaScript code within the Templater Tool, you can use these variables to reference external data:</p>
            <ul>
                <li>
                    <b>key</b>
                    <p>The key of the current iterative item. Uses the <a href="#unique_key">Unique Key</a> to calculate</p>
                </li>
                <li>
                    <b>col[]</b>
                    <p>An array that represents all of the column values for the current iterative item. Items may not be exactly as they are in the spreadsheet if any of the <a href="#key_value_maps">Key Value Maps</a> have been configured to <a href="#key_value_maps_overwrite">Overwrite</a></p>
                </li>
                <li>
                    <b>spreadsheets[][][][]</b>
                    <p>A 4D array that represents all of the spreadsheets that have been uploaded to the Templater Tool. The array axises are as follows:</p>
                    <ul>
                        <li>Spreadsheet file (in the order they were uploaded)</li>
                        <li>Individual sheets (additional tabs that exist within a spreadsheet file)</li>
                        <li>Rows (rows within any given sheet)</li>
                        <li>Columns (values within the columns of any given row)</li>
                    </ul>
                </li>
                <li>
                    <b>vars.*</b>
                    <p>The object where all of the custom variables are defined. To reference a specific variable that you define, you would do <b>vars.&lt;variable_name&gt;</b></p>
                </li>
            </ul>
            <a href="#about" style="margin: 20px 0 40px 0; display: inline-block; width: 100%; text-align: center;">Back to the top</a>
        </div>
        <div class="col d-lg-none"></div>
    </div>
</template>

<style scoped>
    h1, h2, h3, h4, h5, h6 {
        font-family: Calibri, sans-serif;
        letter-spacing: 0px;
        margin: 0;
        padding: 0;
    }

    h3 { font-size: 20px; }
    h2 { font-size: 30px; }
    p, ul li { font-size: 1.1em; }

    a {
        color: #4171c4 !important;
        text-decoration: none;
        font-weight: bold;
    }
</style>

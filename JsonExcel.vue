<template>
	<a
			:id="id_name"
            :title="title"
            :class="classes"
			@click="generate_excel">
		<slot>
			Download Excel
		</slot>
	</a>
</template>



<script>
import {get, isObject, has} from 'lodash';
    export default {
        data: function(){
            return {
                animate   : true,
                animation : '',
            }
        },
        props: {
            'data':{
                type: Array,
                required: true
            },
            'type':{
                type: String,
                default: 'button'
            },
            'title':{
                type: String
            },
            'classes': {
                type: Object
            },
            'vfields':{
                type: Object,
                required: true
            },
            'name':{
                type: String,
                default: "data.xls"
            }
        },
        created: function () {
        },
        computed:{
            id_name : function(){
                var now = new Date().getTime();
                return 'export_' + now;
            }
        },
        methods: {
            emitXmlHeader: function () {
                var headerRow =  '<ss:Row>\n';
                for (var colName in this.vfields) {

                    var name = colName;
                    if(isObject(this.vfields[colName]) && has(this.vfields[colName], 'title')) {
                        name = this.vfields[colName]['title'];
                    }

                    headerRow += '<ss:Cell>\n';
                    headerRow += '<ss:Data ss:Type="String">' + this.upperFirst(name) + '</ss:Data>\n';
                    headerRow += '</ss:Cell>\n';
                }
                headerRow += '</ss:Row>\n';
                return '<?xml version="1.0"?>\n' +
                    '<ss:Workbook xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">\n' +
                    '<ss:Worksheet ss:Name="Sheet1">\n' +
                    '<ss:Table>\n\n' + headerRow;
            },

            emitXmlFooter: function() {
                return '\n</ss:Table>\n' +
                    '</ss:Worksheet>\n' +
                    '</ss:Workbook>\n';
            },

            upperFirst: function(str) {
                return str.charAt(0).toUpperCase() + str.slice(1);
            },

            jsonToSsXml: function (jsonObject) {
                var row;
                var col;
                var xml;
                var data = typeof jsonObject != "object"
                    ? JSON.parse(jsonObject)
                    : jsonObject;

                xml = this.emitXmlHeader();

                for (row = 0; row < data.length; row++) {
                    xml += '<ss:Row>\n';

                    for (col in data[row]) {
                        if( this.vfields[col] !== undefined) {
                            var type = 'String';
                            var value = '';

                            if(isObject(this.vfields[col]) && has(this.vfields[col], 'key')) {
                                value = get(data[row][col], this.vfields[col]['key'], 'N/A');
                            } else {
                                value = data[row][col];
                            }

                            if(value === null) {
                                value = '';
                            }

                            xml += '<ss:Cell>\n';
                            xml += '<ss:Data ss:Type="' + type + '">';
                            xml += String(value).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;') + '</ss:Data>\n';
                            xml += '</ss:Cell>\n';
                        }
                    }

                    xml += '</ss:Row>\n';
                }

                xml += this.emitXmlFooter();
                return xml;
            },
            generate_excel: function (content, filename, contentType) {
                var blob = new Blob([this.jsonToSsXml(this.data)], {
                    'type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                });

                var a = document.getElementById(this.id_name);
                a.href = window.URL.createObjectURL(blob);
                a.download = this.name;
            }
        }
    }
</script>

Vue.component('set-structure', {
    template: '<Row>\
            <i-col span="11">\
                <Card>\
                    <p slot="title">文档结构</p>\
                    <a href="#" slot="extra" @click.prevent="restStructure">\
                        <Icon type="ios-loop-strong"></Icon>\
                        重置\
                    </a>\
                    <Scroll style="height: 322px">\
                       <ul id="treeDemo" class="ztree"></ul>\
                    </Scroll>\
                </Card>\
            </i-col>\
            <i-col span="12" offset="1">\
                <Card id="editor">\
                    <p slot="title">节点设置</p>\
                    <div style="padding: 0">\
                        <textarea id="articleEditor"></textarea>\
                        <i-input  type="textarea" v-model="nodeContent" :rows="11" placeholder="节点内容..." style="padding-top: 25px"></i-input>\
                    </div>\
                </Card>\
            </i-col>\
        </Row>',
    props: ['treedata', 'docids', 'showModel'],
    data: function() {
        return {
            data5: [],
            page: {
                page: 1,
                pageTotal: 0,
                size: 5
            },
            columns: [],
            docData: [],
            docPageData: [],
            docIdArr: "",
            activeNode: {},
            nodeContent: "",
            setting: {
                edit: {
                    enable: true,
                    showRemoveBtn: false,
                    showRenameBtn: false
                },
                view: {
                    showIcon: false,
                    fontCss: {
                        "display": "inline-block",
                        "width": "80%",
                        "overflow": "hidden",
                        "text-overflow": "ellipsis",
                        "white-space": "nowrap",
                        "height": "25px"
                    }
                },
                data: {
                    simpleData: {
                        enable: true
                    }
                },
                callback: {
                    onClick: this.selectNode
                }
            }
        };
    },
    mounted: function() {
        var $__5 = this;
        this.columns = [{
            title: '文档名称',
            align: 'center',
            key: 'docFileName',
            render: function(h, params) {
                return h('div', [h('span', {
                    style: {
                        display: 'inline-block',
                        width: '100%',
                        overflow: 'hidden',
                        textOverflow: 'ellipsis',
                        whiteSpace: 'nowrap'
                    },
                    domProps: {
                        title: params.row.docFileName
                    }
                }, params.row.docFileName)]);
            }
        }, {
            title: '节点标题',
            align: 'center',
            key: 'structures[0].name',
            render: function(h, params) {
                return h('div', [h('span', {
                    style: {
                        display: 'inline-block',
                        width: '100%',
                        overflow: 'hidden',
                        textOverflow: 'ellipsis',
                        whiteSpace: 'nowrap'
                    },
                    domProps: {
                        title: params.row.structures[0].content
                    }
                }, params.row.structures[0].content)]);
            }
        }, {
            title: '操作',
            key: 'action',
            width: 70,
            align: 'center',
            render: function(h, params) {
                return h('div', [h('Button', {
                    props: {
                        type: 'primary',
                        size: 'small'
                    },
                    on: {
                        click: function() {
                            $__5.choice(params.index);
                        }
                    }
                }, '选择')]);
            }
        }];
        tinymce.init({
            selector: '#articleEditor',
            branding: false,
            elementpath: false,
            height: '100px',
            width: '100%',
            language: 'zh_CN.GB2312',
            toolbar_items_size: 'middle',
            footer: false,
            menubar: false,
            plugins: ['advlist autolink lists link image charmap print preview anchor textcolor save', 'searchreplace visualblocks code fullscreen', 'insertdatetime media table contextmenu paste code help wordcount'],
            toolbar1: 'formatselect | bullist numlist removeformat  save',
            autosave_interval: '20s',
            save_onsavecallback: function() {
                $__5.onsavecallback();
            }
        });
        this.getDate5();
    },
    methods: {
        onsavecallback: function() {
            var activeEditor = tinyMCE.activeEditor;
            var editBody = activeEditor.getBody();
            activeEditor.selection.select(editBody);
            var newText = activeEditor.selection.getContent({
                'format': 'text'
            });
            this.activeNode.title = newText;
            this.activeNode.name = newText;
            var treeObj = $.fn.zTree.getZTreeObj("treeDemo");
            treeObj.updateNode(this.activeNode);
            this.$emit('change-data', this.data5);
        },
        nodeDetail: function(data, node) {
            var _this = this;
            if (undefined != data.searchCode) {
                var formData = new FormData();
                formData.append('docIds', _this.docIdArr);
                formData.append('nodeSearchCode', data.searchCode);
                $.ajax({
                    url : URL + '/template/node',
                    data: formData,
                    contentType: false,
                    processData: false,
                    method : "post",
                    async : false,
                    success : function (response){
                        _this.docData = eval(response);
                        _this.docPageData = _this.docData.slice(0, _this.page.size);
                        _this.page.pageTotal = _this.docData.length;
                        _this.$Modal.info({
                            title: '节点来源',
                            scrollable: true,
                            render: function(h, param) {
                                return h('Card', [h('p', {
                                    props: {
                                        slot: 'title'
                                    }
                                }), h('Row', [h('i-table', {
                                    props: {
                                        height: "300",
                                        stripe: true,
                                        border: true,
                                        columns: _this.columns,
                                        data: _this.docPageData
                                    }
                                }), h('div', {
                                    style: {
                                        margin: '10xp',
                                        overflow: 'hidden'
                                    }
                                }, [h('div', {
                                    style: {
                                        float: 'right'
                                    }
                                }, [h('Page', {
                                    props: {
                                        total: _this.page.pageTotal,
                                        current: _this.page.page,
                                        pageSize: _this.page.size,
                                        showTotal: true,
                                        placement: 'top'
                                    },
                                    on: {
                                        'on-change': _this.handlePage
                                    }
                                })])])])]);
                            }
                        });
                    }
                });
            } else {
                tinymce.activeEditor.setContent(data.title);
            }
            _this.activeNode = node;
        },
        selectNode: function(root, node, data) {
            if (this.showModel) {
                this.nodeDetail(data, node);
            } else {
                tinymce.activeEditor.setContent(data.title);
                this.nodeContent = data.content;
            }
        },
        choice: function(index) {
            var _this = this;
            var choiceData = _this.docPageData[index];
            tinymce.activeEditor.setContent(choiceData.structures[0].content);
            _this.nodeContent = choiceData.structures[0].content;
            _this.page.page = 1;
            _this.$Modal.remove();
        },
        handlePage: function(value) {
            if (undefined != value) {
                var _this = this;
                _this.page.page = value;
                _this.docPageData = _this.docData.slice((value - 1) * _this.page.size, value * _this.page.size);
            }
        },
        onselectNode: function(nodes) {
            tinymce.activeEditor.setContent(nodes[0].title);
        },
        append: function(data) {
            var children = data.children || [];
            children.push({
                title: '新节点',
                expand: true
            });
            this.$set(data, 'children', children);
        },
        remove: function(root, node, data) {
            var parentKey = root.find(function(el) {
                return el === node;
            }).parent;
            var parent = root.find(function(el) {
                return el.nodeKey === parentKey;
            }).node;
            var index = parent.children.indexOf(data);
            parent.children.splice(index, 1);
        },
        restStructure: function() {
            this.getDate5();
        },
        getDate5: function() {
            this.data5 = JSON.parse(JSON.stringify(this.treedata));
            $.fn.zTree.init($("#treeDemo"), this.setting, this.data5);
        },
        getSearchCodes: function() {
            this.docIdArr = this.docids;
        }
    },
    watch: {
        treedata: function() {
            this.getDate5();
        },
        docids: function() {
            this.getSearchCodes();
        }
    },
    destroyed: function() {
        tinymce.get('articleEditor').destroy();
    }
});
Vue.component('down-tree', {
    template: "<i-input v-model=\"category\" id=\"cateoryTree\"  placeholder=\"请输入公文分类\" disabled>\n            <Dropdown slot=\"append\"  trigger=\"click\" placement=\"bottom-end\" >\n                <a href=\"javascript:void(0)\">\n                    <Icon type=\"arrow-down-b\"></Icon>\n                </a>\n                    <DropdownMenu slot=\"list\" style=\"text-align: left;height:150px\">\n                        <Scroll :height=\"150\">\n                            <Tree :data=\"data2\" show-checkbox @on-check-change=\"selectCat\"></Tree><!-- :render=\"renderContent\" -->\n                         </Scroll>\n                    </DropdownMenu>\n            </Dropdown>\n        </i-input>",
    props: ['category'],
    data: function() {
        return {
            data2: []
        };
    },
    methods: {
        disabledBortherNode: function(treeData, treeNode) {
            var _this = this;
            treeData.forEach(function(ele, index, array) {
                if (ele.id == treeNode.pId) {
                    var bortherNode = ele.children;
                    bortherNode.forEach(function(ele, index, array) {
                        if (treeNode.id != ele.id) {
                            ele.disableCheckbox = true;
                        }
                    });
                } else {
                    if (undefined != ele.children && ele.children.length > 0) {
                        _this.disabledBortherNode(ele.children, treeNode);
                    } else {}
                }
            });
        },
        selectAllNode: function(treeData) {
            var _this = this;
            treeData.forEach(function(ele, index, array) {
                if (undefined != ele.children && ele.children.length > 0) {
                    _this.selectAllNode(ele.children);
                } else {
                    ele.disableCheckbox = false;
                }
            });
        },
        selectCat: function(nodes) {
            var _this = this;
            _this.selectAllNode(_this.data2);
            nodes.forEach(function(ele, index, array) {
                _this.disabledBortherNode(_this.data2, ele);
            });
            var selectNodes = [];
            var searchCodes = [];
            nodes.filter(function(item) {
                if (!item.children) {
                    selectNodes.push(item.title);
                    searchCodes.push(item.searchCode);
                }
            });
            _this.category = selectNodes.join(',');
            var cateData = {
                title: selectNodes.join(','),
                searchCode: searchCodes.join(',')
            };
            _this.$emit('change-category', cateData);
        }
    },
    mounted: function() {
        var _this = this;
        $.ajax({
            url : URL + '/template/category/tree',
            async : false,
            method : "post",
            dataType : 'json',
            success : function (response){
                _this.data2 = response.children;
            }
        });
    }
});
Vue.component('upload-file', {
    template: "<div>\n        <Upload type=\"drag\" v-if=\"this.file == null\"\n                :before-upload=\"handleUpload\"\n                action=\"//jsonplaceholder.typicode.com/posts/\">\n            <div style=\"margin-top: 100px\">\n                <Icon type=\"ios-cloud-upload\" size=\"200\" style=\"color: #3399ff\"></Icon>\n                <h3>点击或者拖拽文件</h3>\n            </div>\n       </Upload>\n       <div v-if =\"this.file !=null\" style=\"text-align: center\">\n            <h3 style=\"padding-top: 10px\">上传的文件: <b>{{ file.name }}</b></h3>\n            <div class=\"documentIcon\">\n                <Icon type=\"document-text\" size=\"300\" style=\"color: #3399ff;\"></Icon>\n            </div>\n            <div>\n                <Tooltip content=\"删除文档\" placement=\"top\">\n                    <Icon type=\"ios-trash-outline\" size=\"40\" style=\"color: #ff3f48; cursor: pointer;\"\n                      @click.native=\"removeFile\"></Icon>\n                </Tooltip>\n                <Tooltip content=\"上传并拆分文档\" placement=\"top\">\n                    <Icon type=\"ios-cloud-upload-outline\" size=\"40\" @click.native=\"upload\"\n                          style=\"color: #36ffe2;padding-left: 25px; cursor: pointer;\"></Icon>\n                </Tooltip>         \n            </div>\n        </div>\n   </div>",
    props: ['current'],
    data: function() {
        return {
            file: null
        };
    },
    methods: {
        handleUpload: function(file) {
            this.file = file;
            return false;
        },
        removeFile: function() {
            this.file = null;
        },
        upload: function() {
            this.$Message.success('Success');
            this.$emit('change-current', this.current += 1, this.file.name);
        }
    }
});
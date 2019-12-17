Vue.component("ztree", {
    template: '<ul id="treeDemo" class="ztree"></ul>',
    props: ['nodeData'],
    data: function() {
        return {
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
                    onClick: this.onClick
                }
            }
        }
    },
    mounted() {
        var nodeData = this.nodeData;
        $.fn.zTree.init($("#treeDemo"), this.setting, nodeData);
    },
    methods: {
        onClick: function(event, treeId, treeNode, clickFlag) {
            this.$emit('click-node', treeNode)
        }
    }
})
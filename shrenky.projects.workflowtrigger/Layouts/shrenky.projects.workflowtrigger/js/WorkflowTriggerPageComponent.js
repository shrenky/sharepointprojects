Type.registerNamespace('WorkflowTrigger.Ribbon');

WorkflowTrigger.Ribbon.RibbonComponent = function () {
    WorkflowTrigger.Ribbon.RibbonComponent.initializeBase(this);
}

WorkflowTrigger.Ribbon.RibbonComponent.get_instance = function () {
    if (!WorkflowTrigger.Ribbon.RibbonComponent.s_instance) {
        WorkflowTrigger.Ribbon.RibbonComponent.s_instance = new WorkflowTrigger.Ribbon.RibbonComponent();
    }
    return WorkflowTrigger.Ribbon.RibbonComponent.s_instance;
}

WorkflowTrigger.Ribbon.RibbonComponent.prototype = {
    focusedCommands: null,
    globalCommands: null,
    registerWithPageManager: function () {
        SP.Ribbon.PageManager.get_instance().addPageComponent(this);
        SP.Ribbon.PageManager.get_instance().get_focusManager().requestFocusForComponent(this);
    },

    unregisterWithPageManager: function () {
        SP.Ribbon.PageManager.get_instance().removePageComponent(this);
    },

    init: function () { },

    getFocusedCommands: function () {
        return ['shrenky.projects.workflowtrigger.PopulateMenus'];
    },

    getGlobalCommands: function () {
        return ['shrenky.projects.workflowtrigger.PopulateMenus'];
    },

    canHandleCommand: function (commandId) {
        if (commandId === 'shrenky.projects.workflowtrigger.PopulateMenus') {
            return true;
        }
        else { return false; }
    },

    handleCommand: function (commandId, properties, sequence) {
        if (commandId === 'shrenky.projects.workflowtrigger.PopulateMenus') {
            properties.PopulationXML = this.GetDynamicMenuXml();
        }
        else {
            return handleCommand(commandId, properties, sequence);
        }
    },

    isFocusable: function () { return true; },

    receiveFocus: function () { return true; },

    yieldFocus: function () { return true; },

    GetDynamicMenuXml: function () {
        var counter = 0;
        var data = workflowtrigger.data.WorkflowData;
        var xml = '<Menu Id = "shrenky.projects.workflowtrigger.Anchor.Menu">'
        + '<MenuSection Id="shrenky.projects.workflowtrigger.Anchor.Menu.MenuSection1" >'
        + '<Controls Id="shrenky.projects.workflowtrigger.Anchor.Menu.MenuSection1.Controls">';
        var len = data.length;
        for (var i = 0;  i < len; i++){
            counter = counter + 1;
            var current = data[i];
            var workflowAssociationId = current.WorkflowAssociationId;
            var workflowName = current.WorkflowTitle;
            var workflowDesc = current.WorkflowDescription;
            var buttonXml = String.format(
                    '<Button Id= "shrenky.projects.workflowtrigger.Anchor.Menu.MenuSection1.Menu{0}" '
                    + 'Command="shrenky.projects.workflowtrigger.TriggerMenuClick" '
                    + 'MenuItemId="{1}" '
                    + 'LabelText="{2}" '
                    + 'ToolTipTitle="{2}" '
                    + 'ToolTipDescription="{3}" TemplateAlias="o1"/>', counter, workflowAssociationId, workflowName, workflowDesc);
            xml += buttonXml;
        }
        if (counter === 0) {
            var msgXml = String.format(
                    '<Button Id= "shrenky.projects.workflowtrigger.Anchor.Menu.MenuSection1.Menu{0}" '
                    + 'Command="shrenky.projects.workflowtrigger.MessageMenuClick" '
                    + 'MenuItemId="1" '
                    + 'LabelText="{0}" '
                    + 'ToolTipTitle="{0}" '
                    + 'ToolTipDescription="{0}" TemplateAlias="o1"/>', 'No workflow associated on current list');
            xml += buttonXml;
        }
        xml += '</Controls>' + '</MenuSection>' + '</Menu>';
        return xml;
    },

    selectedItems: [],
    workflowAssociationId: null,

    TriggerWorkflow : function(workflowAssoId, param)
    {
        this.workflowAssociationId = workflowAssoId;
        var ctx = SP.ClientContext.get_current();
        var web = ctx.get_web();
        var lists = web.get_lists();
        var listId = SP.ListOperation.Selection.getSelectedList();
        var list = lists.getById(listId);
        var items = SP.ListOperation.Selection.getSelectedItems();
        this.selectedItems = [];
        if (items.length > 0) {
            for (var i in items) {
                var id = items[i].id;
                var item = list.getItemById(id);
                this.selectedItems.push(item);
                ctx.load(item);
            }
            ctx.executeQueryAsync(Function.createDelegate(this, this.TriggerWorkflowExec), Function.createDelegate(this, this.TriggerWorkflowFailed));
        }
        else {
            alert("Please select item");
        }
    },

    TriggerWorkflowExec: function ()
    {
        for (var i in this.selectedItems) {
            var item = this.selectedItems[i];
            var itemId = item.get_item("ID");
            var itemTitle = item.get_item("Title");
            var itemFileRef = item.get_item("FileRef");
            var webUrl = _spPageContextInfo.webAbsoluteUrl;
            var url = webUrl + itemFileRef;
            var id = this.workflowAssociationId;
            !function outer(id, url, itemTitle) {
                $().SPServices({
                    debug: true,
                    operation: "StartWorkflow",
                    async: true,
                    item: url,
                    templateId: id,
                    workflowParameters: "<Data/>",
                    completefunc: function () { SP.UI.Notify.addNotification("Start workflow on " + itemTitle + " Successfully", true); }
                });
            }(id, url, itemTitle);
        }
    },

    TriggerWorkflowFailed: function ()
    {
        alert("Failed to trigger");
    }
}

WorkflowTrigger.Ribbon.RibbonComponent.registerClass('WorkflowTrigger.Ribbon.RibbonComponent', CUI.Page.PageComponent);
WorkflowTrigger.Ribbon.RibbonComponent.get_instance().registerWithPageManager();
NotifyScriptLoadedAndExecuteWaitingJobs("WorkflowTriggerPageComponent.js");

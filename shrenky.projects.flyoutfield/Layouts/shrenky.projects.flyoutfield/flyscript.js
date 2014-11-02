(function () {
    var fieldJsLinkOverride = {};
    fieldJsLinkOverride.Templates = {};
    fieldJsLinkOverride.Templates.Fields = {
        'FlyoutField': { 'View': ShrenkyProjectsFlyoutField_GetFlyingCallout }
    };

    SPClientTemplates.TemplateManager.RegisterTemplateOverrides(fieldJsLinkOverride);
})();

function ShrenkyProjectsFlyoutField_GetFlyingCallout(ctx) {
    var flyValue = ctx.CurrentItem.Flyout; //STSHTMLEncode needed
    var divId = "CalloutDiv" + ctx.CurrentItem.ID;
    return "<div id='" + divId + "' style=\"cursor: pointer;\" onmouseenter=\"ShrenkyProjectsFlyoutField_InitCallout('" + divId + "','" + flyValue + "');\">  <span id=\"ms-pageDescriptionImage\"> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" + flyValue + "</span></div>";
}

var content = '';

function ShrenkyProjectsFlyoutField_InitCallout(divId, param) {
    SP.SOD.executeFunc('callout.js', 'LoadCallOut', function () { console.log('load callout.js') });
    SP.SOD.executeOrDelayUntilScriptLoaded(function () { ShrenkyProjectsFlyoutField_AddCallout(divId, param); }, 'callout.js');
}

function ShrenkyProjectsFlyoutField_AddCallout(divId, param) {
    var calloutElement = document.getElementById(divId);

    var calloutOptions = new CalloutOptions();
    calloutOptions.ID = 'Callout_' + divId;
    calloutOptions.launchPoint = calloutElement;
    calloutOptions.beakOrientation = 'leftRight';
    calloutOptions.content = '<h2>' + param + '</h2>';
    calloutOptions.title = 'Title';
    CalloutManager.createNewIfNecessary(calloutOptions);
}
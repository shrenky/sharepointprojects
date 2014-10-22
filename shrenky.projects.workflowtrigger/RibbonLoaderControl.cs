using Microsoft.SharePoint;
using Microsoft.SharePoint.WebControls;
using Microsoft.SharePoint.Workflow;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Web.UI;

namespace shrenky.projects.workflowtrigger
{
    public class RibbonLoaderControl : Control
    {
        private SPWorkflowAssociationCollection workflowAssociations = null;

        public SPWorkflowAssociationCollection WorkflowAssociations
        {
            get
            {
                if (workflowAssociations == null)
                { 
                    if(SPContext.Current.List != null)
                    {
                        workflowAssociations = SPContext.Current.List.WorkflowAssociations;
                    }
                }
                return workflowAssociations;
            }
        }

        protected override void OnPreRender(EventArgs e)
        {
            SPRibbon ribbon = SPRibbon.GetCurrent(this.Page);
            if (ribbon != null)
            {
                RegisterWorkflowTriggerRibbon(ribbon);
                RegisterWorkflowAssociationIds();
            }

            base.OnPreRender(e);
        }

        private void RegisterWorkflowTriggerRibbon(SPRibbon ribbon)
        {
            ScriptLink.RegisterScriptAfterUI(this.Page, "CUI.js", false, true);
            ScriptLink.RegisterScriptAfterUI(this.Page, "SP.Ribbon.js", false, true);
            ScriptLink.RegisterScriptAfterUI(this.Page, "/_layouts/15/shrenky.projects.workflowtrigger/js/WorkflowTriggerPageComponent.js", false, true);
            ScriptLink.RegisterScriptAfterUI(this.Page, "/_layouts/15/shrenky.projects.workflowtrigger/js/jquery-1.11.1.min.js", false, true);
            ScriptLink.RegisterScriptAfterUI(this.Page, "/_layouts/15/shrenky.projects.workflowtrigger/js/jquery.SPServices-2014.01.min.js", false, true);
        }

        private void RegisterWorkflowAssociationIds()
        {
            if (WorkflowAssociations != null)
            {
                StringBuilder builder = new StringBuilder();
                builder.Append("window.workflowtrigger = window.workflowtrigger || {};");
                builder.Append("workflowtrigger.__namespace = true;");
                builder.Append("workflowtrigger.data = workflowtrigger.data || {};");
                List<DataObject> dataObjects = new List<DataObject>();
                foreach (SPWorkflowAssociation item in WorkflowAssociations)
                {
                    DataObject data = new DataObject { WorkflowAssociationId = item.Id.ToString("B"), WorkflowTitle = item.Name, WorkflowDescription = item.Description };
                    dataObjects.Add(data);
                }
                JavaScriptSerializer serializer = new JavaScriptSerializer();
                builder.AppendFormat("workflowtrigger.data = {0};", serializer.Serialize(new {WorkflowData = dataObjects}));
                string key = "WorkflowTriggerScriptKey";
                string script = builder.ToString();
                if (!this.Page.ClientScript.IsStartupScriptRegistered(key))
                {
                    this.Page.ClientScript.RegisterStartupScript(this.GetType(), key, script, true);
                }
            }
        }

    }
}

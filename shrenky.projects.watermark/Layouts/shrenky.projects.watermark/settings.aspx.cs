using System;
using Microsoft.SharePoint;
using Microsoft.SharePoint.WebControls;
using Microsoft.SharePoint.Workflow;
using System.Web.UI.WebControls;
using System.IO;
using System.Xml.Serialization;
using System.Text;
using Microsoft.SharePoint.Utilities;
using System.Collections;
using System.Xml;

namespace shrenky.projects.watermark.Layouts.shrenky.projects.watermark
{
    public partial class settings : LayoutsPageBase
    {
        protected string workflowName = string.Empty;
        protected bool allowStartManually;
        protected bool startWhenAddNewItem;
        protected bool startWhenChangeItem;
        protected SPWorkflowTemplate baseTemplate;
        protected SPWorkflowAssociation assocTemplate;
        protected string waterMarkText = string.Empty;
        protected HyperLink returnLink;

        protected string queryParams;
        protected SPList List;
        protected SPContentType contentType;
        protected bool isContentTypeTemplate;

        protected Guid TaskList;
        protected Guid HistoryList;

        #region life cycle
        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            workflowName = this.Request.Params["WorkflowName"];

            FetchAssociationInfo();
            GetTaskAndHistoryList();

            this.EnsureRequestParamsParsed();
        }

        protected void OKButton_Click(object sender, EventArgs e)
        {
            string taskListName = string.Empty;
            string historyListName = string.Empty;
            SPList taskList = null;
            SPList historyList = null;
            if (!IsValid)
                return;
            if (!isContentTypeTemplate)
            {
                // If the user requested a new task list, create it.
                if (TaskList == Guid.Empty)
                {
                    taskListName = string.Format("{0} Tasks", workflowName);
                    string description = string.Format("Task list for the {0} workflow.", workflowName);
                    TaskList = Web.Lists.Add(taskListName, description, SPListTemplateType.Tasks);
                }

                // If the user requested a new history list, create it.
                if (HistoryList == Guid.Empty)
                {
                    historyListName = string.Format("{0} History", workflowName);
                    string description = string.Format("History list for the {0} workflow.", workflowName);
                    HistoryList = Web.Lists.Add(historyListName, description, SPListTemplateType.WorkflowHistory);
                }
                taskList = Web.Lists[TaskList];
                historyList = Web.Lists[HistoryList];
            }
            // Perform association (if it does not already exist).
            bool isNewAssociation = true;
            if (assocTemplate == null)
            {
                isNewAssociation = true;
                if (!isContentTypeTemplate)
                    assocTemplate = SPWorkflowAssociation.CreateListAssociation(baseTemplate,
                                                                  workflowName,
                                                                  taskList,
                                                                  historyList);
                else
                {
                    assocTemplate = SPWorkflowAssociation.CreateSiteContentTypeAssociation(baseTemplate,
                                                                    workflowName,
                                                                    taskListName,
                                                                    historyListName);
                }
            }
            else // Modify existing template.
            {
                isNewAssociation = false;
                assocTemplate.Name = workflowName;
                assocTemplate.SetTaskList(taskList);
                assocTemplate.SetHistoryList(historyList);
            }

            assocTemplate.Name = workflowName;
            assocTemplate.AllowManual = allowStartManually;
            assocTemplate.AutoStartCreate = startWhenAddNewItem;
            assocTemplate.AutoStartChange = startWhenChangeItem;

            if (assocTemplate.AllowManual)
            {
                SPBasePermissions newPerms = SPBasePermissions.EmptyMask;

                if (Request.Params["ManualPermEditItemRequired"] == "ON")
                    newPerms |= SPBasePermissions.EditListItems;
                if (Request.Params["ManualPermManageListRequired"] == "ON")
                    newPerms |= SPBasePermissions.ManageLists;

                assocTemplate.PermissionsManual = newPerms;
            }

            //TODO Content type
            // Place data from form into the association template.
            assocTemplate.AssociationData = SerializeFormToString();
            SPList list = SPContext.Current.List;
            if (isNewAssociation)
                list.AddWorkflowAssociation(assocTemplate);
            else
                list.UpdateWorkflowAssociation(assocTemplate);

            //if (assocTemplate.AllowManual && SPContext.Current.List.EnableMinorVersions)
            //{
            //    // If this WF was selected to be the content approval WF 
            //    // (m_setDefault = true, see association page) then enable content
            //    // Approval for the list.
            //    if (list.DefaultContentApprovalWorkflowId != assocTemplate.Id && !m_setDefault)
            //    {
            //        if (!list.EnableModeration)
            //            list.EnableModeration = true;
            //        list.DefaultContentApprovalWorkflowId = assocTemplate.Id;
            //        list.Update();
            //    }
            //    else if (list.DefaultContentApprovalWorkflowId == assocTemplate.Id && !m_setDefault)
            //    {
            //        // Reset the DefaultContentApprovalWorkflowId
            //        list.DefaultContentApprovalWorkflowId = Guid.Empty;
            //        list.Update();
            //    }
            //}

            string strUrl = string.Format("{0}/_layouts/15/WrkSetng.aspx{1}", Web.Url, queryParams);//StrGetRelativeUrl(this, "WrkSetng.aspx", null) + queryParams;
            Response.Redirect(strUrl);
        }

        protected override void OnPreRender(EventArgs e)
        {
            base.OnPreRender(e);
            if (assocTemplate != null)
            {
                PopulatePageFromXml((string)assocTemplate.AssociationData);
            }
        }

        protected void CancelButton_Click(object sender, EventArgs e)
        {
            string strUrl = string.Format("{0}/_layouts/15/WrkSetng.aspx{1}", Web.Url, queryParams);//StrGetRelativeUrl(this, "WrkSetng.aspx", null) + queryParams;
            Response.Redirect(strUrl);
        }
        #endregion

        #region private
        // Deserializes the association xml string and populates the
        // fields in the form.
        internal void PopulatePageFromXml(string associationXml)
        {
            XmlSerializer serializer = new XmlSerializer(typeof(FormData));
            XmlTextReader reader = new XmlTextReader(new System.IO.StringReader(associationXml));
            FormData formdata = (FormData)serializer.Deserialize(reader);
            WaterMartTextBox.Text = formdata.WaterMarkText;
        }

        private void FetchAssociationInfo()
        {
            SPWorkflowAssociationCollection collection;
            baseTemplate = Web.WorkflowTemplates[new Guid(Request.Params["WorkflowDefinition"])];
            assocTemplate = null;

            if (contentType != null)
            {
                // Associating with a content type.
                collection = contentType.WorkflowAssociations;
                //returnLink.Text = contentType.Name;
                //returnLink.NavigateUrl = "ManageContentType.aspx" + queryParams;
            }
            else
            {
                //Associated with list:
                SPList list = Web.Lists[new Guid(Request.Params["List"])];
                collection = list.WorkflowAssociations;
                //returnLink.Text = list.Title;
                //returnLink.NavigateUrl = list.DefaultViewUrl;
            }

            if (collection == null || collection.Count < 0)
            {
                throw new SPException("No workflow association found");
            }

            startWhenAddNewItem = Request.Params["AutoStartCreate"] == "ON";
            startWhenChangeItem = Request.Params["AutoStartChange"] == "ON";
            allowStartManually = Request.Params["AllowManual"] == "ON";

            // Check if workflow association already exists.
            string strGuidAssoc = Request.Params["GuidAssoc"];
            if (strGuidAssoc != string.Empty)
            {
                assocTemplate = collection[new Guid(strGuidAssoc)];
            }

            //Check duplicate name
            SPWorkflowAssociation duplicate = collection.GetAssociationByName(workflowName, Web.Locale);
            if(duplicate != null && (assocTemplate == null || assocTemplate.Id != duplicate.Id))
            {
                throw new SPException(string.Format("Workflow association {0} exists.", workflowName));
            }

        }

        private void GetTaskAndHistoryList()
        {
            string taskListName;
            string historyListName;
            if (isContentTypeTemplate)
            {
                TaskList = new Guid(Request.Params["TaskList"]);
                HistoryList = new Guid(Request.Params["HistoryList"]);
            }
            else
            {

                // If the user has requested that a new task or history list be created, check
                // that the name does not duplicate the name of an existing list. If it does, show
                // the user an appropriate error page.

                string taskListGuid = Request.Params["TaskList"];
                if (taskListGuid[0] != 'z') // already existing list
                {
                    TaskList = new Guid(taskListGuid);
                }
                else  // new list
                {
                    SPList list = null;
                    taskListName = taskListGuid.Substring(1);
                    try
                    {
                        list = Web.Lists[taskListName];
                    }
                    catch (ArgumentException)
                    {
                    }

                    if (list != null)
                        throw new SPException("A list already exists with the same name as that proposed for the new task list. Use your browser's Back button and either change the name of the workflow or select an existing task list.&lt;br&gt;");
                }

                // Do the same for the history list
                string strHistoryListGuid = Request.Params["HistoryList"];
                if (strHistoryListGuid[0] != 'z') // user selected already existing list
                {
                    HistoryList = new Guid(strHistoryListGuid);
                }
                else // User wanted a new list
                {
                    SPList list = null;

                    historyListName = strHistoryListGuid.Substring(1);

                    try
                    {
                        list = Web.Lists[historyListName];
                    }
                    catch (ArgumentException)
                    {
                    }
                    if (list != null)
                        throw new SPException("A list already exists with the same name as that proposed for the new history list. Use your browser's Back button and either change the name of the workflow or select an existing history list.&lt;br&gt;");
                }
            }
        }

        internal string SerializeFormToString()
        {
            FormData data = new FormData();

            data.WaterMarkText = this.WaterMartTextBox.Text;
            
            using (MemoryStream stream = new MemoryStream())
            {
                XmlSerializer serializer =
                    new XmlSerializer(typeof(FormData));
                serializer.Serialize(stream, data);
                stream.Position = 0;
                byte[] bytes = new byte[stream.Length];
                stream.Read(bytes, 0, bytes.Length);
                return Encoding.UTF8.GetString(bytes);
            }
        }

        // Ensure the we get the context variables.
        protected void EnsureRequestParamsParsed()
        {
            string strListID = Request.QueryString["List"];
            string strCTID = Request.QueryString["ctype"];
            if (strListID != null)
                List = Web.Lists[new Guid(strListID)];
            if (strCTID != null)
            {
                queryParams = "?ctype=" + strCTID;
                if (List != null)
                {
                    queryParams += "&List=" + strListID;
                    contentType = List.ContentTypes[new SPContentTypeId(strCTID)];
                }
                else
                {
                    contentType = Web.ContentTypes[new SPContentTypeId(strCTID)];
                    isContentTypeTemplate = true;
                }
            }
            else
                queryParams = "?List=" + strListID;
        }
        #endregion

        #region url helper
        public string strGroup = "Group";
        public string StrGetRelativeUrl(System.Web.UI.Page pgIn, string strPage, string strGrpName)
        {
            string strUrl = StrGetWebRelativePath(pgIn) + strPage;
            // No need to UrlEncodeAsUrl strUrl as it would have 
            // already been encoded.
            if (FValidString(strGrpName))
            {
                strUrl += "?" + SPHttpUtility.UrlKeyValueEncode(strGroup) + "=" + SPHttpUtility.UrlKeyValueEncode(strGrpName);
            }
            return strUrl;
        }
        public string StrGetWebRelativePath(System.Web.UI.Page pgIn)
        {
            string strT = SPUtility.OriginalServerRelativeRequestUrl;

            int iLastSlash = strT.LastIndexOf("/");
            if (iLastSlash > 0)
            {
                strT = SPHttpUtility.UrlPathEncode(SPHttpUtility.UrlPathDecode(strT, true), true);
                iLastSlash = strT.LastIndexOf("/");
                return strT.Substring(0, iLastSlash + 1);
            }
            else
            {
                return string.Empty;
            }
        }
        public bool FValidString(string strIn)
        {
            return FValidString(strIn, 2048);
        }
        public bool FValidString(string strIn, uint nMaxLength)
        {
            return (strIn != null && strIn.Length > 0 && strIn.Length <= nMaxLength);
        }
        #endregion

    }
}

using System;
using System.ComponentModel;
using System.ComponentModel.Design;
using System.Collections;
using System.Drawing;
using System.Linq;
using System.Workflow.ComponentModel.Compiler;
using System.Workflow.ComponentModel.Serialization;
using System.Workflow.ComponentModel;
using System.Workflow.ComponentModel.Design;
using System.Workflow.Runtime;
using System.Workflow.Activities;
using System.Workflow.Activities.Rules;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Workflow;
using Microsoft.SharePoint.WorkflowActions;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace shrenky.projects.watermark.workflow
{
    public sealed partial class workflow : StateMachineWorkflowActivity
    {
        public workflow()
        {
            InitializeComponent();
        }

        public SPWorkflowActivationProperties workflowProperties = new SPWorkflowActivationProperties();

        private void onWorkflowActivated1_Invoked(object sender, ExternalDataEventArgs e)
        {
            FormData data = FormDataHelper.DeserializeFormData(workflowProperties.AssociationData);
            SPListItem item = workflowProperties.Item;
            if (item != null && item.File != null)
            {
                SPFile file = item.File;
                string waterMarkText = data.WaterMarkText;
                byte[] byteArray = file.OpenBinary();
                using (MemoryStream memStr = new MemoryStream())
                {
                    memStr.Write(byteArray, 0, byteArray.Length);
                    using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(memStr, true))
                    {
                        Document document = wordDoc.MainDocumentPart.Document;
                        Paragraph firstParagraph = document.Body.Elements<Paragraph>().FirstOrDefault();
                        if (firstParagraph != null)
                        {
                            Paragraph testParagraph = new Paragraph(
                                new Run(
                                    new Text(waterMarkText)));
                            firstParagraph.Parent.InsertBefore(testParagraph,
                                firstParagraph);
                        }
                    }

                    string linkFileName = file.Item["LinkFilename"] as string;
                    file.ParentFolder.Files.Add(linkFileName, memStr, true);
                }
            }
        }
    }
}

using Microsoft.SharePoint.WebControls;
using System;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;

namespace shrenky.projects.flyoutfield.CONTROLTEMPLATES
{
    public partial class FlyoutFieldUserControl : BaseFieldControl
    {
        protected TextBox FlyingFieldControl;

        protected override string DefaultTemplateName
        {
            get
            {
                return "FlyoutFieldRenderingTemplate";
            }
        }

        public override object Value
        {
            get
            {
                EnsureChildControls();
                return FlyingFieldControl;
            }
            set
            {
                EnsureChildControls();
                FlyingFieldControl.Text = value.ToString();
            }
        }

        protected override void CreateChildControls()
        {
            if (Field == null) return;
            base.CreateChildControls();
            if (ControlMode == SPControlMode.Display) return;
            FlyingFieldControl = (TextBox)TemplateContainer.FindControl("FlyingFieldControl");
            if (ControlMode == SPControlMode.New)
            {
                FlyingFieldControl.Text = "";
            }
        }
    }
}

using Microsoft.SharePoint;
using Microsoft.SharePoint.WebControls;
using shrenky.projects.flyoutfield.CONTROLTEMPLATES;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace shrenky.projects.flyoutfield
{
    public class FlyoutField : SPFieldText
    {
        public FlyoutField(SPFieldCollection fields, string fieldName):base(fields, fieldName){}
        public FlyoutField(SPFieldCollection fields, string typeName, string displayName) : base(fields, typeName, displayName) { }

        public override BaseFieldControl FieldRenderingControl
        {
            get
            {
                BaseFieldControl fieldControl = new FlyoutFieldUserControl();
                fieldControl.FieldName = this.InternalName;
                return fieldControl;  
            }
        }

        public override string JSLink
        {
            get
            {
                return "/_layouts/15/shrenky.projects.flyoutfield/flyscript.js";
            }
            set
            {
                base.JSLink = value;
            }
        }
    }
}

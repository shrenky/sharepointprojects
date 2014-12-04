<%@ Assembly Name="shrenky.projects.watermark, Version=1.0.0.0, Culture=neutral, PublicKeyToken=1cd99f874a92087c" %>
<%@ Import Namespace="Microsoft.SharePoint.ApplicationPages" %>
<%@ Register Tagprefix="SharePoint" Namespace="Microsoft.SharePoint.WebControls" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register Tagprefix="Utilities" Namespace="Microsoft.SharePoint.Utilities" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register Tagprefix="asp" Namespace="System.Web.UI" Assembly="System.Web.Extensions, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35" %>
<%@ Import Namespace="Microsoft.SharePoint" %>
<%@ Assembly Name="Microsoft.Web.CommandUI, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="settings.aspx.cs" Inherits="shrenky.projects.watermark.Layouts.shrenky.projects.watermark.settings" DynamicMasterPageFile="~masterurl/default.master" %>
 <%@ Register TagPrefix="wssuc" TagName="InputFormSection" src="/_controltemplates/InputFormSection.ascx" %> 
<%@ Register TagPrefix="wssuc" TagName="InputFormControl" src="/_controltemplates/InputFormControl.ascx" %>
<%@ Register TagPrefix="wssuc" TagName="ButtonSection" src="/_controltemplates/ButtonSection.ascx" %>

<asp:Content ID="PageHead" ContentPlaceHolderID="PlaceHolderAdditionalPageHead" runat="server">
    <script type="text/javascript">
        function _spBodyOnLoad() {
            try {
                window.focus();
            }
            catch (e) {
            }
        }
        function DoValidateAndSubmit() {
            return true;
        }
</script>
</asp:Content>

<asp:Content ID="Main" ContentPlaceHolderID="PlaceHolderMain" runat="server">

        <div class="ms-authoringcontrols">
            <SharePoint:EncodedLiteral runat="server" ID="WaterMarkTextLabel" Text="Please specify watermark text: " EncodeMethod="HtmlEncode"></SharePoint:EncodedLiteral>
            <SharePoint:InputFormTextBox runat="server" ID="WaterMartTextBox" ></SharePoint:InputFormTextBox>
        </div>
                            
        <div style="float:right">
            <asp:Button runat="server" class="ms-ButtonHeightWidth" OnClick="OKButton_Click" Text="OK" id="OKButton" />
            <asp:Button runat="server" ID="CancelButton" Text="Cancel" CssClass="ms-ButtonHeightWidth" OnClick="CancelButton_Click" />
        </div>
    
        <input type="hidden" name="WorkflowDefinition" value='<% SPHttpUtility.NoEncode(SPHttpUtility.HtmlEncode(Request.Form["WorkflowDefinition"]), Response.Output); %>'/>
        <input type="hidden" name="WorkflowName" value='<% SPHttpUtility.NoEncode(SPHttpUtility.HtmlEncode(Request.Form["WorkflowName"]), Response.Output); %>'/>
        <input type="hidden" name="AllowManual" value='<% SPHttpUtility.NoEncode(SPHttpUtility.HtmlEncode(Request.Form["AllowManual"]), Response.Output); %>'/>
        <input type="hidden" name="GuidAssoc" value='<% SPHttpUtility.NoEncode(SPHttpUtility.HtmlEncode(Request.Form["GuidAssoc"]), Response.Output); %>'/>
        <input type="hidden" name="HistoryList" value='<% SPHttpUtility.NoEncode(SPHttpUtility.HtmlEncode(Request.Form["HistoryList"]), Response.Output); %>'/>
        <input type="hidden" name="TaskList" value='<% SPHttpUtility.NoEncode(SPHttpUtility.HtmlEncode(Request.Form["TaskList"]), Response.Output); %>'/>
        <input type="hidden" name="AutoStartCreate" value='<% SPHttpUtility.NoEncode(SPHttpUtility.HtmlEncode(Request.Form["AutoStartCreate"]), Response.Output); %>'/>
        <input type="hidden" name="AutoStartChange" value='<% SPHttpUtility.NoEncode(SPHttpUtility.HtmlEncode(Request.Form["AutoStartChange"]), Response.Output); %>'/>
        
</asp:Content>

<asp:Content ID="PageTitle" ContentPlaceHolderID="PlaceHolderPageTitle" runat="server">
<SharePoint:EncodedLiteral runat="server" ID="PageTitleLiteral" Text="Workflow association form" EncodeMethod="HtmlEncode"></SharePoint:EncodedLiteral>
</asp:Content>

<asp:Content ID="PageTitleInTitleArea" ContentPlaceHolderID="PlaceHolderPageTitleInTitleArea" runat="server" >
    <%
        string pageTitleinTitleArea = "Customize " + workflowName;
        SPHttpUtility.HtmlEncode(pageTitleinTitleArea, Response.Output);
    %>
    :
    <asp:HyperLink ID="returnUrl" runat="server" />
</asp:Content>

<asp:Content ID="PageDescription" ContentPlaceHolderID="PlaceHolderPageDescription" runat="server">
    <%
        string pageDesc = "Use this page to customize this instance of " + workflowName + ".";
        SPHttpUtility.HtmlEncode(pageDesc, Response.Output);
    %>
</asp:Content>
  

<%@ Assembly Name="shrenky.projects.flyoutfield, Version=1.0.0.0, Culture=neutral, PublicKeyToken=1cd99f874a92087c" %>
<%@ Assembly Name="Microsoft.Web.CommandUI, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register Tagprefix="SharePoint" Namespace="Microsoft.SharePoint.WebControls" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register Tagprefix="Utilities" Namespace="Microsoft.SharePoint.Utilities" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register Tagprefix="asp" Namespace="System.Web.UI" Assembly="System.Web.Extensions, Version=3.5.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35" %>
<%@ Import Namespace="Microsoft.SharePoint" %> 
<%@ Register Tagprefix="WebPartPages" Namespace="Microsoft.SharePoint.WebPartPages" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Control Language="C#" AutoEventWireup="true" CodeBehind="FlyoutFieldUserControl.ascx.cs" Inherits="shrenky.projects.flyoutfield.CONTROLTEMPLATES.FlyoutFieldUserControl" %>
<SharePoint:RenderingTemplate ID="FlyoutFieldRenderingTemplate" runat="server">  
    <script src="/_layouts/15/shrenky.projects.flyoutfield/jquery-1.11.1.min.js" type="text/javascript"></script>
    <Template>  
        <asp:TextBox runat="server" ID="FlyingFieldControl"></asp:TextBox>  
    </Template>  
</SharePoint:RenderingTemplate>  

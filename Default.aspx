<%@ Page Title="Home Page" Async="true" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="ImageToForm._Default" %>

<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">

    <div class="jumbotron">
        <h1>Image2Form</h1>
        <p class="lead">A Tool that converts the screenshot you upload to a form accessible in MasterWorks!</p>
        
    </div>

    <div class="row">
        <div class="col-md-4">
            <h2>Getting started</h2>
            <p> 
            Just Upload a MasterWorks form details page screenshot and the tool generates both the XML Template of the form and the Data Dictionary for the uploaded image 
            </p>
            <p>
                <a class="btn btn-default" >Learn more &raquo;</a>
            </p>
        </div>
        <p>  
        Please Select an Image file:     
        <asp:FileUpload ID="FUP_WatermarkImage" runat="server" /> 
            
    </p>  
    <p>  
        <asp:Button ID="btnUpload" runat="server" Text="Upload Image" onclick="btnUpload_Click" />  
    </p>  
    <p>  
         
    </p>  
    </div>

</asp:Content>

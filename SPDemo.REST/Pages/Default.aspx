<%-- The following 4 lines are ASP.NET directives needed when using SharePoint components --%>

<%@ Page Inherits="Microsoft.SharePoint.WebPartPages.WebPartPage, Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" MasterPageFile="~masterurl/default.master" Language="C#" %>

<%@ Register TagPrefix="Utilities" Namespace="Microsoft.SharePoint.Utilities" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register TagPrefix="WebPartPages" Namespace="Microsoft.SharePoint.WebPartPages" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register TagPrefix="SharePoint" Namespace="Microsoft.SharePoint.WebControls" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>

<%-- The markup and script in the following Content element will be placed in the <head> of the page --%>
<asp:Content ContentPlaceHolderID="PlaceHolderAdditionalPageHead" runat="server">
    <script type="text/javascript" src="../Scripts/jquery-1.8.2.min.js"></script>
    <script type="text/javascript" src="/_layouts/15/sp.runtime.js"></script>
    <script type="text/javascript" src="/_layouts/15/sp.js"></script>
    <script type="text/javascript" src="/_layouts/15/SP.RequestExecutor.js"></script>

    <!-- Add your CSS styles to the following file -->
    <link rel="Stylesheet" type="text/css" href="../Content/kendo.common.css" />
    <link rel="Stylesheet" type="text/css" href="../Content/kendo.metro.css" />
    <link rel="Stylesheet" type="text/css" href="../Content/App.css" />

    <!-- Add your JavaScript to the following file -->
    <script type="text/javascript" src="../Scripts/json2.js"></script>
    <script type="text/javascript" src="../Scripts/date.format.js"></script>
    <script type="text/javascript" src="../Scripts/date.js"></script>
    <script type="text/javascript" src="../Scripts/kendo.web.js"></script>
    <script type="text/javascript" src="../Scripts/App.js"></script>
    <script type="text/javascript">
        $(document).ready(function () {
            initializePage();
        });
    </script>
</asp:Content>

<%-- The markup in the following Content element will be placed in the TitleArea of the page --%>
<asp:Content ContentPlaceHolderID="PlaceHolderPageTitleInTitleArea" runat="server">
    SharePoint 2013 REST Demo
</asp:Content>

<%-- The markup and script in the following Content element will be placed in the <body> of the page --%>
<asp:Content ContentPlaceHolderID="PlaceHolderMain" runat="server">
    <div id="tabstrip" class="tabDiv">
        <ul>
            <li class="k-state-active">
                List Items
            </li>
            <li>
                File Upload
            </li>
            <li>
                Site Provisioning
            </li>
            <li>
                Search
            </li>
            <li>
                Social
            </li>
        </ul>
        <div id="listItems" class="contentDiv">
            <div id="demoSelector" style="width:100%;padding-bottom:10px;padding-top:10px;">
                <span style="font-size:14pt;font-weight:bold;">Query Option:</span><br /><br />
                <select id="restOptions" style="width:95%;font-size:14pt;">
                    <option value="Select a query option">Select a query option</option>
                    <option value="/web/lists/getbytitle('Industry GDP by State 2010')/items">Get all items</option>
                    <option value="/web/lists/getbytitle('Industry GDP by State 2010')/items/getbyid(1000)">Get item with ID of '1000'</option>
                    <option value="/web/lists/getbytitle('Industry GDP by State 2010')/items?$top=10">Get top 10 items</option>
                    <option value="/web/lists/getbytitle('Industry GDP by State 2010')/items()?$filter=State eq 'Texas'">Get items where State equals 'Texas'</option>
                    <option value="/web/lists/getbytitle('Industry GDP by State 2010')/items()?$filter=startswith(State,'a')">Get items where State starts with 'A'</option>
                    <option value="/web/lists/getbytitle('Industry GDP by State 2010')/items()?$filter=Amount gt 1000000">Get items where Amount is greater than $1,000,000.00</option>
                    <option value="/web/lists/getbytitle('Industry GDP by State 2010')/items()?$filter=Amount gt 1000000&$orderby=Amount desc">Get items where Amount is greater than $1,000,000.00, order by Amount descending</option>
                    <option value="/web/lists/getbytitle('Industry GDP by State 2010')/items()?$filter=Amount gt 1000000&$skip=20">Get items where Amount is greater than $1,000,000.00, skip first 20 records</option>
                    <option value="demoGrid;https://binarywaveinc.sharepoint.com/sites/dev;Industry GDP by State 2010;A377EC9D-2043-4B69-A6C0-3BF182AD5149">Dynamic View</option>
                </select>
            </div>
            <div id="demoCommand" style="width:95%;">
                <span style="font-size:14pt;font-weight:bold;">Query URL:</span><br/>
                <div id="demoCommand_Text" style="width:100%;height:40px;font-size:14pt;color:blue;padding:5px;"></div>
            </div>
            <div id="demoActions" style="width:100%;">
                <button id="buttonSubmit" style="border:1px solid black;">Submit</button>
            </div>
            <div id="demoBottom" style="width:100%;padding-top:10px;padding-bottom:10px;">
                <br /><span style="font-size:14pt;font-weight:bold;">Query Results:</span><br/><br/>
                <div id="demoGrid"></div>
            </div>
        </div>
        <div id="fileUpload" class="contentDiv">
            <div id="fileSelector" style="width:100%;padding-bottom:10px;padding-top:10px;">
                <span style="font-size:14pt;font-weight:bold;">Document:</span><br /><br />
                <input id="inputFile" type="file" style="width:99%;height:30px;font-size:14pt;" />
            </div>
            <div id="fileActions" style="width:100%;">
                <button id="fileSubmit" style="border:1px solid black;">Submit</button>
            </div>
        </div>
        <div id="siteProvisioning" class="contentDiv">
            <div id="siteName" style="width:100%;padding-bottom:10px;padding-top:10px;">
                <span style="font-size:14pt;font-weight:bold;">Site Name:</span><br /><br />
                <input id="inputSite" type="text" style="width:99%;height:30px;font-size:14pt;" />
            </div>
            <div id="siteActions" style="width:100%;">
                <button id="siteSubmit" style="border:1px solid black;">Submit</button>
                &nbsp;&nbsp;&nbsp;<span id="siteStatus" style="color:blue;"></span>
            </div>
        </div>
        <div id="searchContent" class="contentDiv">
            <div id="searchQuery" style="width:100%;padding-bottom:10px;padding-top:10px;">
                <span style="font-size:14pt;font-weight:bold;">Keywords:</span><br /><br />
                <input id="inputSearch" type="text" style="width:99%;height:30px;font-size:14pt;" />
            </div>
            <div id="searchActions" style="width:100%;">
                <button id="searchSubmit" style="border:1px solid black;">Submit</button>
            </div>
            <div id="searchResults" style="width:100%">
                <br /><span style="font-size:14pt;font-weight:bold;">Search Results:</span><br/><br/>
                <div id="searchOutput" style="width:99%"></div>
            </div>
        </div>
        <div id="socialContent" class="contentDiv">
            <div id="socialActions" style="width:100%;padding-top:10px;">
                <button id="socialSubmit" style="border:1px solid black;">Get Followed Content</button>
            </div>
            <div id="socialResults" style="width:100%;">
                <br /><span style="font-size:14pt;font-weight:bold;">Results:</span><br/><br/>
                <div id="socialOutput" style="width:99%"></div>
            </div>
        </div>
        </div>
    </div>
</asp:Content>

# Sample Office Add-in using ADAL JS and CORS generated using the Yeoman Office Generator

This sample illustrates how to extend an Office Add-in generated using the [Yeoman Office Generator](https://github.com/OfficeDev/generator-office) to connect with Office 365 using ADAL JS and CORS. This sample add-in is a Task Pane Add-in that allows you to search for documents in SharePoint Online directly from Office client applications. 

This sample is based on the documentation provided with the Office Generator repo at [https://github.com/OfficeDev/generator-office/blob/master/docs/adal-config.md](https://github.com/OfficeDev/generator-office/blob/master/docs/adal-config.md).

## Running this sample

1. Clone this repository
1. Register a new web application in your Azure Active Directory
1. Copy the application ID
1. Grant the application the following **Office 365 SharePoint Online** permissions: **Run search queries as a user** and **Read items in all site collections**
1. In the application's manifest enable implicit OAuth flow
1. In the `app/app.config.js` file change the value of the `tenantName` variable to your tenant and the `appId` to the ID of your application as registered in the Azure Active Directory
1. Upload the `manifest.xml` file to the **Apps for Office** in your App Catalog (located at `https://yourtenant.sharepoint.com/sites/apps/_layouts/15/start.aspx#/AgaveCatalog/Forms/AllItems.aspx`)
1. Open a Word/PowerPoint/Excel/Project document and insert the add-in
1. Search for a document. The top 5 documents matching your query will be displayed in the Task pane from where you can open them to view their contents

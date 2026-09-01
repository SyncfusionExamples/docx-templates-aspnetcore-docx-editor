# Document Template Studio with Syncfusion DOCX Editor

## Introduction

**Document Template Studio Service** is an ASP.NET Core web application
for creating, editing, managing, and previewing DOCX document templates.

The sample uses Syncfusion<sup style="font-size:70%">&reg;</sup> [ASP,NET CORE DOCX Editor](https://www.syncfusion.com/docx-editor-sdk/asp-net-core-docx-editor?utm_source=github&utm_medium=listing&utm_campaign=github-github-documenteditor-examples) as
the document editing experience and **Syncfusion DocIO** for DOCX
processing and mail merge operations.

The application provides a server-rendered Razor Pages interface where
users can:

-   View available document templates.
-   Create a new template from a blank document or an existing DOCX
    file.
-   Edit DOCX templates in the Syncfusion Document Editor.
-   Add and manage merge fields.
-   Save and publish edited templates.
-   Preview a template with sample JSON data using mail merge.
-   Delete templates from the template catalog.

The `Input Data` folder is intentionally included as sample input data.
Users can choose any suitable DOCX template from `Input Data/Template`
when creating a template and use the corresponding JSON file from
`Input Data/JSONForMailMerge` for the **Preview With Data (Mail Merge)**
operation.

------------------------------------------------------------------------

## Application workflow

The application follows this general workflow:

``` text
                    ┌─────────────────────┐
                    │   Template Studio   │
                    │     Dashboard       │
                    └──────────┬──────────┘
                               │
             ┌─────────────────┼─────────────────┐
             │                 │                 │
             ▼                 ▼                 ▼
       Create Blank      Upload DOCX       Open Existing
         Template          Template          Template
             │                 │                 │
             └─────────────────┼─────────────────┘
                               ▼
                    ┌─────────────────────┐
                    │   Document Editor   │
                    │  Syncfusion EJ2     │
                    └──────────┬──────────┘
                               │
                 ┌─────────────┼─────────────┐
                 │             │             │
                 ▼             ▼             ▼
            Add Fields      Edit DOCX    Save/Publish
                 │             │             │
                 └─────────────┼─────────────┘
                               ▼
                    ┌─────────────────────┐
                    │ Preview With Data   │
                    │    JSON Mail Merge  │
                    └──────────┬──────────┘
                               ▼
                    ┌─────────────────────┐
                    │ Merged DOCX Preview │
                    └─────────────────────┘
```

------------------------------------------------------------------------

# How to run this sample

## Prerequisites

Install the following before running the application:

1.  **.NET 10 SDK**
3.  A valid **Syncfusion license** for environments where a license is
    required.

Verify the .NET SDK:

``` bash
dotnet --version
```

The project targets:

``` xml
<TargetFramework>net10.0</TargetFramework>
```

------------------------------------------------------------------------

## 1. Restore NuGet packages

Open a terminal in the project root and run:

``` bash
dotnet restore
```

This restores the Syncfusion ASP.NET Core and document-processing
dependencies used by the sample.

------------------------------------------------------------------------

## 2. Configure the Syncfusion license

The application registers the Syncfusion license in `Program.cs`:

``` csharp
Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense("");
```

Replace the empty value with the appropriate Syncfusion license key when
required for your environment.

Do not commit a private license key to a public source repository.

------------------------------------------------------------------------

## 4. Build the project

Run:

``` bash
dotnet build
```

Resolve any build or package errors before starting the application.

------------------------------------------------------------------------

## 5. Run the application

Run:

``` bash
dotnet run
```

The configured project launch profile uses:

``` text
http://localhost:5212
```

Open the application in a browser:

``` text
http://localhost:5212
```

The application should open the Template Studio dashboard.

------------------------------------------------------------------------

# Using the sample

## Create a template

From the dashboard:

1.  Select **Create Template**.

2.  Enter the template name.

3.  Select the template type.

4.  Enter the description if required.

5.  Enter one or more signer roles.

6.  Optionally select a DOCX file from:

    ``` text
    Input Data/Template/
    ```

7.  Create the template.

There are two supported creation flows:

### Create from an existing DOCX

Select a DOCX document from the `Input Data/Template` folder.

The application copies the uploaded document into:

``` text
wwwroot/Templates/
```

and creates the corresponding entry in:

``` text
wwwroot/Data/templates.json
```

### Create a blank template

A template can also be created without uploading a DOCX file.

The Document Editor opens a blank document. The user can author the
document and then save/publish it.

------------------------------------------------------------------------

## Edit a template

Open a template from the dashboard.

The editor page uses the Syncfusion EJ2 DocumentEditor and provides the
document authoring experience.

The editor supports the application's document operations through:

``` text
/api/DocumentEditor/*
```

The application loads the template DOCX, converts it into the format
required by DocumentEditor, and displays it in the browser.

------------------------------------------------------------------------

## Add merge fields

The editor provides a Merge Fields panel.

A merge field can be added to the document and used as a Word mail-merge
field.

For example:

``` text
«DonorName»
«DonorAddress»
«DonationAmount»
«DonationDate»
```


------------------------------------------------------------------------

## Save and publish

When the user saves the edited template, the DocumentEditor sends the
document data to:

``` text
POST /api/DocumentEditor/Save
```

The server converts the editor content to DOCX and stores the published
document under:

``` text
wwwroot/Templates/
```

The corresponding template catalog entry is updated in:

``` text
wwwroot/Data/templates.json
```

------------------------------------------------------------------------

## Preview With Data (Mail Merge)

The sample includes JSON data files under:

``` text
Input Data/JSONForMailMerge/
```

For example:

``` text
Input Data/JSONForMailMerge/Donation Official Tax Receipt.json
```

Use the JSON file that corresponds to the selected DOCX template.

The Preview With Data workflow:

1.  Reads the current document from the editor.
2.  Converts the editor document to DOCX.
3.  Sends the DOCX and JSON data to the server.
4.  Uses Syncfusion DocIO MailMerge to populate the merge fields.
5.  Converts the merged document back into the format required by
    DocumentEditor.
6.  Opens the merged result for preview.

The server-side mail merge operation is:

``` text
POST /api/DocumentEditor/MailMerge
```

------------------------------------------------------------------------

# Input Data

The `Input Data` directory is intentionally part of the sample.

It provides example files that users can use to test the complete
template workflow.

## DOCX sample templates

Located at:

``` text
Input Data/Template/
```

Available samples include:

``` text
Donation Official Tax Receipt.docx
Donation Thank-You Letter.docx
Donor Impact Letter.docx
Pledge Payment Reminder.docx
```

## Mail-merge JSON data

Located at:

``` text
Input Data/JSONForMailMerge/
```

Available samples include:

``` text
Donation Official Tax Receipt.json
Donation Thank-You Letter.json
Donor Impact Letter.json
Pledge Payment Reminder.json
```

The JSON structure should contain keys corresponding to the merge fields
used by the selected DOCX document.

For example, a template containing:

``` text
«DonorName»
«DonationAmount»
«DonationDate»
```

requires matching data keys such as:

``` json
{
  "DonorName": "John Doe",
  "DonationAmount": "$500.00",
  "DonationDate": "September 1, 2026"
}
```

The exact JSON structure used by the sample files should be followed
when adding more complex mail-merge data.

------------------------------------------------------------------------
## Resources

- **Product page:**   [Syncfusion® ASP.NET CORE DOCX Editor](https://www.syncfusion.com/docx-editor-sdk/asp-net-core-docx-editor?utm_source=github&utm_medium=listing&utm_campaign=github-github-documenteditor-examples) 

- **Documentation:**   [Syncfusion® ASP.NET CORE DOCX Editor - Documentation](https://help.syncfusion.com/document-processing/word/word-processor/asp-net-core/overview?utm_source=github&utm_medium=listing&utm_campaign=github-github-documenteditor-examples) 

- **Online demo:**   [Syncfusion® ASP.NET CORE DOCX Editor - Online demo](https://document.syncfusion.com/demos/docx-editor/asp-net-core/documenteditor/default?utm_source=github&utm_medium=listing&utm_campaign=github-github-documenteditor-examples) 

## Support and feedback 

For any other queries, reach our [Syncfusion® support team](https://support.syncfusion.com/?utm_source=github&utm_medium=listing&utm_campaign=github-github-documenteditor-examples) or post the queries through the [community forums](https://www.syncfusion.com/forums?utm_source=github&utm_medium=listing&utm_campaign=github-github-documenteditor-examples). 

Request new feature through [Syncfusion® feedback portal](https://www.syncfusion.com/feedback?utm_source=github&utm_medium=listing&utm_campaign=github-github-documenteditor-examples). 

## License

This is a commercial product and requires a paid license for possession or use Syncfusion's licensed software, including this component, is subject to the terms and conditions of [Syncfusion's EULA](https://www.syncfusion.com/license/studio/34.1.29/syncfusion_essential_studio_eula.pdf?utm_source=github&utm_medium=listing&utm_campaign=github-github-documenteditor-examples). You can purchase a licnense [here](https://www.syncfusion.com/sales/products?utm_source=github&utm_medium=listing&utm_campaign=github-github-documenteditor-examples) or start a free 30\-day trial [here](https://www.syncfusion.com/account/manage-trials/start-trials?utm_source=github&utm_medium=listing&utm_campaign=github-github-documenteditor-examples). 

---
page_type: sample
description: This sample demonstrates how to use the Microsoft Graph JavaScript SDK to send email in Excel from Office Add-ins.
products:
- ms-graph
- microsoft-graph-email-api
languages:
- javascript
---

# One Click Launch - Microsoft Excel Mail Merge Sample Office Add-in

[![Node.js build](https://github.com/microsoftgraph/msgraph-training-office-addin/actions/workflows/node.js.yml/badge.svg)](https://github.com/microsoftgraph/msgraph-training-office-addin/actions/workflows/node.js.yml) ![License.](https://img.shields.io/badge/license-MIT-green.svg)

This sample demonstrates how to one click launch Microsoft Excel Mail Merge Sample Office Add-in, which uses the Microsoft Graph JavaScript SDK to send emails in Excel from Office Add-ins.

## Applies to
- Excel on Windows, Mac, and Online.

# Required Steps - Get an Application Id

## Register a web application with the Azure Active Directory admin center


1. Open a browser and navigate to the [Microsoft Entra admin center](https://aad.portal.azure.com). Login using a **personal account** (aka: Microsoft Account) or **Work or School Account**.

1. Select **Identity** in the left-hand navigation, then select **App registrations** under **Applications**.

1. Select **New registration**. On the **App registrations** page, set the values as follows.

    - Set **Name** to `Office Add-in Graph Tutorial`.
    - Set **Supported account types** to **Accounts in any organizational directory and personal Microsoft accounts**.
    - Under **Redirect URI**, set the first drop-down to `Single-page application (SPA)` and set the value to `https://officedev.github.io/Excel-Scenario-based-Add-in-Samples/Mail-Merge-Sample-Add-in/consent.html`.

1. Select **Register**. On the **Office Add-in Graph Tutorial** page, copy the value of the **Application (client) ID** and save it, you will need it in the next step.


**Note**: This step needs to be **performed only once** by add-in developer, aiming to integrate your app with the Microsoft identity platform and establishing the information that it uses to get tokens. After successful registration and add-in published, **customer can use it directly**, do not need to register again. 

## Additional resources
You may explore additional resources at the following links:
- [Office Add-ins code samples](https://github.com/OfficeDev/Office-Add-in-samples)
- [Office Add-ins documentation](https://learn.microsoft.com/en-us/office/dev/add-ins/)

## Questions and comments

- Did you experience any problems with the sample? [Create an issue](https://github.com/OfficeDev/Excel-Scenario-based-Add-in-Samples/issues/new) and we'll help you out.
- We'd love to get your feedback about this sample. Go to our [Office samples survey](https://forms.office.com/Pages/ResponsePage.aspx?id=v4j5cvGGr0GRqy180BHbR8GFRbAYEV9Hmqgjcbr7lOdUNVAxQklNRkxCWEtMMFRFN0xXUFhYVlc5Ni4u) to give feedback and suggest improvements.
- For general questions about developing Office Add-ins, go to [Microsoft Q&A](https://learn.microsoft.com/answers/topics/office-js-dev.html) using the office-js-dev tag.

## Code of conduct

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**
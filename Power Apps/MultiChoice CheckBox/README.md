# MultiChoice CheckBox

## Summary

This sample Power App demonstrates how you can use check boxes for a multi-choice column in SharePoint.

![](./power%20app%20check%20boxes.png)

There are 3 samples in this application
1. Basic Control Example
2. Component Example
3. Form Example

## Setup
First, you will need to create a new SharePoint list in which you will build out your sample content. Add a multi-choice column and make sure to name it **ChoiceCol**. The sample app does not provide a way to create new list items, only update existin list items. I recommend that you create two sample items in your list.

![](./SharePoint%20example%20list.png)

Next, go to [https://make.powerapps.com](https://make.powerapps.com), select your desired environment, open solutions in the left rail, and import the solution [MultiChoiceCheckboxes_1_0_0_1.zip](MultiChoiceCheckboxes_1_0_0_1.zip). During solution import, you will be prompted to select or create a new SharePoint connection. You will then be prompted to provide values for 2 environment variables. 

* SharePoint Check Box Site - Select the site from the drop down or paste the site URL.
* SharePoint Check Box List - Once the site is populated, you can select the list you created in the drop down.

![](./solution%20import%20environment%20variables.png)

## Using the Sample App

To view how the app was constructed it is best to open the app in Edit mode so that you can inspect the various formulas. However, each sample has an overview of the settings to create the sample as text on the screen.

![](./Power%20app%20sample%20screen%201.png)

![](./Power%20app%20sample%20screen%202.png)

![](./Power%20app%20sample%20screen%203.png)

![](./Power%20app%20sample%20screen%204.png)

## Version history

Version|Date|Contributor|Comments
-------|----|----|----
1.0|7/28/2022|Travis Lingenfelder|Initial release

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

**This sample code, scripts, and other resources are not supported under any Microsoft standard support program or service and are meant for illustrative purposes only. The sample code, scripts, and resources are provided AS IS without warranty of any kind. Microsoft further disclaims all implied warranties including, without limitation, any implied warranties of merchantability or of fitness for a particular purpose. The entire risk arising out of the use or performance of this material and documentation remains with you. In no event shall Microsoft, its authors, or anyone else involved in the creation, production, or delivery of the sample be liable for any damages whatsoever (including, without limitation, damages for loss of business profits, business interruption, loss of business information, or other pecuniary loss) arising out of the use of or inability to use the samples or documentation, even if Microsoft has been advised of the possibility of such damages.**

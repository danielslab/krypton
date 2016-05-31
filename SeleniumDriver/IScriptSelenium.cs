/***************************************************************************
** Application: Krypton: ThinkSys Automation Generic System
** File Name: Driver.IScriptSelenium.cs
** Copyright 2012, ThinkSys Software
** All rights reserved
** Developed by: ThinkSys Software 
** Description: Interface to create selenium script in programming language
*****************************************************************************/

using System.Collections.Generic;

namespace Driver
{
    /// <summary>
    ///Interface to provide contract to generate selenium script in programming language
    /// </summary>
    public interface IScriptSelenium
    {
        void Click();
        void SendKeys(string text);
        void Check();
        void Uncheck();
        void Wait();
        void Close();
        void Navigate(string testData);
        void FireEvent();
        void GoBack();
        void GoForward();
        void Refresh();
        void Clear();
        void EnterUniqueData();
        void KeyPress(string testData);
        void SelectItem(string testData);
        void SelectItemByIndex(string testData);
        void WaitForObject();
        void WaitForObjectNotPresent();
        void WaitForObjectProperty();
        void AcceptAlert();
        void DismissAlert();
        void VerifyAlertText();
        void GetObjectProperty(string property);
        void SetAttribute(string property, string propertyValue);
        void VerifyPageProperty(string property, string propertyVal);
        void VerifyTextPresentOnPage(string text);
        void VerifyTextNotPresentOnPage(string text);
        void VerifyListItemPresent(string listItem);
        void VerifyListItemNotPresent(string listItem);
        void VerifyObjectPresent();
        void VerifyObjectNotPresent();
        void VerifyObjectProperty(string property, string propertyVal);
        void VerifyObjectPropertyNot(string property, string propertyVal);
        void GetPageProperty(string propertyType);
        void VerifyPageDisplay();
        void VerifyTextInPageSource(string text);
        void VerifyTextNotInPageSource(string text);
        void VerifyObjectDisplay();
        void VerifyObjectNotDisplay();
        void ExecuteScript(string script, string testObject);
        void DeleteCookies(string testData);
        void ShutDownDriver();
        void SetObjDataRow(Dictionary<string, string> objDataRow);
        void SaveScript();
    }
}

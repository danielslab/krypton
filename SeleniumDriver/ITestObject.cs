/***************************************************************************
** Application: Krypton: ThinkSys Automation Generic System
** File Name: ITestObject.cs
** Copyright 2012, ThinkSys Software
** All rights reserved
** Developed by: ThinkSys Software 
** Description: Interface Contain the Method Require for Browser Automation. 
*****************************************************************************/

using System.Collections.Generic;

namespace Driver
{
    /// <summary>
    ///Interface to provide contract to Perform action on Browser.
    /// </summary>
    public interface ITestObject
    {
        void SetObjDataRow(Dictionary<string, string> objDataRow,string currentStepAction="");
        bool VerifySortOrder(string propertyForSorting, string sortOrder = "");
        void MouseMove();
        void MouseClick();
        void MouseOver();
        void AddAction(string actionToAdd = "", string data = "");
        void PerformAction();
        void ClearText();
        void KeyPress(string key);
        void Click(Dictionary<int, string> keyWordDic = null, string data = null);
        void ClickInThread();
        void CheckMultiple(string[] dataContents);
        void UncheckMultiple(string[] dataContent);
        void Check(string checkStatus = "");
        void UnCheck();
        void SendKeys(string text);
        void Submit();
        void FireEvent(string eventName);
        string WaitForObject(string waitTime, string globalWaitTime, Dictionary<int, string> optionDic, string modifier);
        void WaitForObjectNotPresent(string waitTime, string globalWaitTime, Dictionary<int, string> optionDic);
        void WaitForObjectProperty(string propertyParam, string propertyValueParam, string globalWaitTime, Dictionary<int, string> optionDic);
        void DoubleClick();
        void SelectItem(string[] itemList, bool selectMultiple = false);
        string SelectItemByIndex(string index);
        string GetObjectProperty(string property);
        bool VerifyObjectDisplayed();
        bool VerifyListItemPresent(string listItemName);
        bool VerifyListItemNotPresent(string listItemName);
        bool VerifyObjectPresent();
        bool VerifyObjectNotPresent();
        bool VerifyObjectProperty(string property, string propertyValue, Dictionary<int, string> keywordDic);
        bool VerifyObjectPropertyNot(string property, string propertyValue, Dictionary<int, string> KeywordDic);
        string ExecuteStatement(string scriptToExecute);
        void SetAttribute(string property, string propertyValue);
        void DragAndDrop(Dictionary<string, string> targetObjDic);
        bool UploadFile(string path,string element);
        void WaitForPage(string timeOut);
    }
}
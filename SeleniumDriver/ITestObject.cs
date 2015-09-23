/***************************************************************************
** Application: Krypton: ThinkSys Automation Generic System
** File Name: ITestObject.cs
** Copyright 2012, ThinkSys Software
** All rights reserved
** Developed by: ThinkSys Software 
** Description: Interface Contain the Method Require for Browser Automation. 
*****************************************************************************/
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System.Text.RegularExpressions;
using OpenQA.Selenium.Remote;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Interactions.Internal;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Interfaces;
using OpenQA.Selenium.Appium.MultiTouch;
using OpenQA.Selenium.Appium.Appium;
namespace Driver
{
    /// <summary>
    ///Interface to provide contract to Perform action on Browser.
    /// </summary>
    public interface ITestObject
    {
        void SetObjDataRow(Dictionary<string, string> objDataRow,string CurrentStepAction="");
        bool verifySortOrder(string propertyForSorting, string sortOrder = "");
        void mouseMove();
        void mouseClick();
        void mouseOver();
        void addAction(string actionToAdd = "", string data = "");
        void performAction();
        void ClearText();
        void KeyPress(string key);
        void Click(Dictionary<int, string> keyWordDic = null, string data = null);
        void clickInThread();
        void checkMultiple(string[] dataContents);
        void uncheckMultiple(string[] dataContent);
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
        bool VerifyObjectProperty(string property, string propertyValue, Dictionary<int, string> KeywordDic);
        bool VerifyObjectPropertyNot(string property, string propertyValue, Dictionary<int, string> KeywordDic);
        string ExecuteStatement(string scriptToExecute);
        void SetAttribute(string property, string propertyValue);
        void DragAndDrop(Dictionary<string, string> targetObjDic);
        bool UploadFile(string path,string element);
        
    }
}
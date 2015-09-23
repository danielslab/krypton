/***************************************************************************
** Application: Krypton: ThinkSys Automation Generic System
** File Name: TestDriver.APIRelatedMethods.cs
** Copyright 2012, ThinkSys Software
** All rights reserved
** Developed by: ThinkSys Software
** Description: Contains the functionalities related to API calls.
*****************************************************************************/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Net;
using System.IO;
using Common;
using Newtonsoft.Json;
using System.Net.Security;
namespace TestDriver
{
    class WebAPI
    {
        /// <summary>
        /// This method can be used to request for an API and collect responses
        /// </summary>
        /// <param name="requestHeader">Request header that needs to be sent along with request. 
        /// Header name and value should be saparated by column, while multiple headers should be saparated by pipe | </param>
        /// <param name="responseFormat"></param>
        /// <param name="requestMethod"></param>
        /// <param name="requestUrl"></param>
        /// <param name="requestBody"></param>
        /// <returns></returns>
        public static bool requestWebService(string requestHeader, string responseFormat, string requestMethod, string requestUrl, string requestBody)
        {
            //name of the file, which would contain exact response returned by api call
            string apiResponseFileName = Property.ApiXmlFile + Common.Property.StepNumber + "." + responseFormat;

            //name of the file, which would contain xml equivalent of the api response, may not be converted always
            string xmlFileName = Property.ApiXmlFile + Common.Property.StepNumber + ".xml";

            //Final location where API resonse will be saved
            string XmlFilePath = Property.ResultsSourcePath + "\\" + apiResponseFileName;

            try
            {
                if (File.Exists(XmlFilePath))
                {
                    File.Delete(XmlFilePath);
                }

                //Create an object to request webservice
                HttpWebRequest webServiceRequest = (HttpWebRequest)WebRequest.Create(requestUrl);

                //Assign headers to webservice object
                string[] arrHeaderList = requestHeader.Split(Property.SEPERATOR);
                for (int i = 0; i < arrHeaderList.Length; i++)
                {
                    string[] arrHeader = arrHeaderList[i].Trim().Split(':');
                    string strHeaderName = arrHeader[0].Trim();
                    string strHeaderValue = arrHeader[1].Trim();
                    switch (strHeaderName.ToLower())
                    {
                        case "content-type":
                            webServiceRequest.ContentType = strHeaderValue;
                            break;
                        case "authorization":
                            webServiceRequest.Headers.Add(HttpRequestHeader.Authorization, strHeaderValue);
                            break;
                        default:
                            if(!WebHeaderCollection.IsRestricted(strHeaderName))
                            {
                            webServiceRequest.Headers.Add(arrHeaderList[i].Trim());
                            }
                            break;
                    }
                }
                ServicePointManager.ServerCertificateValidationCallback = new RemoteCertificateValidationCallback
                (
                        delegate { return true; }
                );

                //Fire request
                webServiceRequest.Method = requestMethod;

                if (!requestBody.Equals(string.Empty))
                {
                    byte[] byteRequestBody = Encoding.UTF8.GetBytes(requestBody);
                    Stream requestStream = webServiceRequest.GetRequestStream();
                    requestStream.Write(byteRequestBody, 0, byteRequestBody.Length);
                    requestStream.Close();
                    Common.Property.Remarks = requestBody;
                }
                WebResponse responseXml = webServiceRequest.GetResponse();

                //Collect returns from webservice
                string statusCode = ((HttpWebResponse)responseXml).StatusCode.ToString();
                string statusDescription = ((HttpWebResponse)responseXml).StatusDescription;
                Console.WriteLine("API call response: " + statusCode + "\t" + statusDescription);
                               
                Stream responseData = responseXml.GetResponseStream();
                StreamReader reader = new StreamReader(responseData, true);
                String responseFromServer = reader.ReadToEnd();
                responseFromServer.Replace('\\',' ');

                //Parse API response
                //Store return data to a file and created a variable also
                try
                {
                        File.WriteAllText(XmlFilePath, responseFromServer);
                    
                }
                catch
                {
                   
                }

               
                Utility.SetVariable("APIFilePath", XmlFilePath);

                Property.Attachments = apiResponseFileName;

                // for Bug# 186 (To handle Json which is returned without object and the url having "api")
                #region Converting json to xml in case json was the original format
                if (responseFormat.ToLower().Equals("json") || requestUrl.ToLower().Contains("json") || requestUrl.ToLower().Contains("api"))
                {
                    try
                    {
                        //Converting json to xml
                        XmlNode jsonToXml = null;
                        string jsonString = File.ReadAllText(XmlFilePath);
                        jsonToXml = JsonConvert.DeserializeXmlNode("{\"root\":" + jsonString + "}", Property.ProductName);
                        //New path exclusively for xml file, 
                        //this is achieved by replacing json file name with xml file name in full length path
                        XmlFilePath = XmlFilePath.Replace(apiResponseFileName, xmlFileName);
                        //Write converted xml to the document
                        XmlDocument convertedXmlDoc = new XmlDocument();
                        convertedXmlDoc.CreateXmlDeclaration("1.0", string.Empty, string.Empty);
                        convertedXmlDoc.LoadXml(jsonToXml.OuterXml);
                        convertedXmlDoc.Save(XmlFilePath);
                        //Going forward xml will be used to validate and read attributes
                        Utility.SetVariable("APIFilePath", XmlFilePath);

                    }
                    catch (Exception exJsonToXml)
                    {
                        throw new Exception("Unable to convert Json Response to xml. Error: " + exJsonToXml.Message);
                    }
                }
                #endregion

                return true;
                }
            catch (Exception e)
            {
                if (Property.SnapshotOption.Equals("on failure", StringComparison.OrdinalIgnoreCase) && Property.DebugMode.Equals("false", StringComparison.OrdinalIgnoreCase))
                {
                    Property.Attachments = xmlFileName;
                }
                throw e;
            }
        }

        private static bool IsXMLConvertable(string responseFromServer)
        {
            bool isLoadSuccessful = false;
            try
            {
                XmlDocument convertedXmlDoc = new XmlDocument();
                convertedXmlDoc.CreateXmlDeclaration("1.0", string.Empty, string.Empty);
                convertedXmlDoc.LoadXml(responseFromServer);
                isLoadSuccessful = true;
            }
            catch(Exception ex)
            {
            
            }
            return isLoadSuccessful;
        }


        /// <summary>
        /// This method will download any file from internet location and save as local file
        /// </summary>
        /// <param name="webLocation">
        /// Full internet url
        /// </param>
        /// <param name="localFileName">
        /// File name to be used after file has been downloaded, if same file already exists, it will be overwritten
        /// </param>
        /// <returns></returns>
        public static bool downloadFile(string webLocation, string localFileName = "FileDownload.file")
        {
            localFileName = localFileName.Trim().Replace(".", Common.Property.StepNumber + ".");

            try
            {
                //Delete any existing file with same name
                string XmlFilePath = Property.ResultsSourcePath + "\\" + localFileName;
                string locationVariable = "DownloadedFile";

                if (File.Exists(XmlFilePath))
                {
                    File.Delete(XmlFilePath);
                }
                //Download file from web resources and store on local disk
                try
                {
                    WebClient wc = new WebClient();
                    wc.DownloadFile(webLocation, XmlFilePath);
                    Utility.SetVariable(locationVariable, XmlFilePath);

                    Common.Property.Remarks = "File has been downloaded and its location has been store in variable '" +
                                              locationVariable + "'.";
                }
                catch (Exception fileDownloadError)
                {
                    Common.Property.Remarks = "Unable to download file: " + fileDownloadError.Message;
                    return false;
                }

                Property.Attachments = localFileName;
                return true;
            }
            catch (Exception e)
            {
                Common.Property.Remarks = "Unable to download file: " + e.Message;
                return false;
            }
        }

        public static bool uploadFiletoFTP(string fileName, Uri uploadUrl, string user, string pswd)
        {
            Stream requestStream = null;
            FileStream fileStream = null;
            FtpWebResponse uploadResponse = null;
            try
            {
                FtpWebRequest uploadRequest = (FtpWebRequest)WebRequest.Create(uploadUrl);
                uploadRequest.Method = WebRequestMethods.Ftp.UploadFile;

                ICredentials credentials = new NetworkCredential(user, pswd);
                uploadRequest.Credentials = credentials;

                requestStream = uploadRequest.GetRequestStream();
                fileStream = File.Open(fileName, FileMode.Open);
                byte[] buffer = new byte[1024];
                int bytesRead;
                while (true)
                {
                    bytesRead = fileStream.Read(buffer, 0, buffer.Length);
                    if (bytesRead == 0)
                        break;
                    requestStream.Write(buffer, 0, bytesRead);
                }

                requestStream.Close();
                uploadResponse = (FtpWebResponse)uploadRequest.GetResponse();
                Common.Property.Remarks = "File :'" + fileName + "' has been uploaded to '" + uploadUrl + "'.";
            }
            catch (UriFormatException ex)
            {
                Common.Property.Remarks = "Error: " + ex.Message;
                return false;
            }
            catch (IOException ex)
            {
                Common.Property.Remarks = "Error: " + ex.Message;
                return false;
            }
            catch (WebException ex)
            {
                Common.Property.Remarks = "Error: " + ex.Message;
                return false;
            }
            finally
            {
                if (uploadResponse != null)
                    uploadResponse.Close();
                if (fileStream != null)
                    fileStream.Close();
                if (requestStream != null)
                    requestStream.Close();
            }
            return true;
        }


        /// <summary>
        /// Verify the value of node specified against expectedValue.
        /// </summary>
        /// <param name="attributeName">String : Name of Attribute (node) to verify</param>
        /// <param name="expectedValue">String : Expected Value against which verification will be done</param>
        /// <returns>bool value true OR false</returns>
        public static bool verifyXmlAttribute(string attributeName, string expectedValue, int nodeIndex = 1)
        {

            FileInfo fileInfo = new FileInfo(Utility.GetVariable("APIFilePath"));

            try
            {
                //replacing variable in case variable is used instead of value.
                expectedValue = Utility.ReplaceVariablesInString(expectedValue);
                //Read the value for specified node.
                bool readAttribute = readXmlAttribute(attributeName, nodeIndex);
                if (!readAttribute)
                {
                    return false;
                }

                //get the actual value for specified node.
                string actualValue = Utility.GetVariable(attributeName);


                if (Property.SnapshotOption.Equals("always", StringComparison.OrdinalIgnoreCase) || Property.DebugMode.Equals("true", StringComparison.OrdinalIgnoreCase))
                {
                    Property.Attachments = fileInfo.Name;
                }

                if (Utility.doKeywordMatch(expectedValue, actualValue))
                {
                    return true;
                }
                else
                {
                    if (Property.SnapshotOption.Equals("on failure", StringComparison.OrdinalIgnoreCase) && Property.DebugMode.Equals("false", StringComparison.OrdinalIgnoreCase))
                    {
                        Property.Attachments = fileInfo.Name;
                    }
                    Property.Remarks = "Actual value '" +
                        actualValue + "' of attribute '" + attributeName + "' does not match its expected value '" + expectedValue + "'.";
                    return false;
                }
            }
            catch (Exception e)
            {
                if (Property.SnapshotOption.Equals("on failure", StringComparison.OrdinalIgnoreCase) && Property.DebugMode.Equals("false", StringComparison.OrdinalIgnoreCase))
                {
                    Property.Attachments = fileInfo.Name;
                }
                throw e;
            }
        }
        /// <summary>
        /// Return the integer count of Xml node attribute
        /// </summary>
        /// <param name="attributeName">string : Name of attribute(node) to read</param>
        /// <returns>Count of ndoe</returns>
        public static int ReturnCountofXmlNodes(string attributeName)
        {
            //Connect with xml document first
            XmlDocument doc = new XmlDocument();
            doc.Load(Utility.GetVariable("APIFilePath"));

            //Gets the list of nodes having specified attribute name.
            XmlNodeList list = doc.GetElementsByTagName(attributeName);

            //Store count of all attributes to variable
            Utility.SetVariable(attributeName + "Count", list.Count.ToString());

            return list.Count;
        }
        /// <summary>
        /// Reads the attribute(node) value and store the value in runtimedic with attribute name as key.
        /// </summary>
        /// <param name="attributeName">string : Name of attribute(node) to read.</param>
        /// <param name="nodeIndex">int : position of the node to be read, starting with 1</param>
        /// <param name="valueLookingFor">string : In case position of the node needs to be found, 
        /// pass this as an expected value of attribute</param>
        /// <returns>string value of specified attribute(node).</returns>
        /// Added different handlings and cleaned code
        /// Added Return CountXmlNodes function to return the count of Xml Nodes.
        public static bool readXmlAttribute(string attributeName, int nodeIndex = 1, string valueLookingFor = "")
        {
            try
            {
                bool blnNodeFound = false;
                string varNameForAttributeLocation = attributeName.Trim() + "_location";

                //Handle various possible node index values, node index (input) should always start with 1 and is one based
                if (nodeIndex <= 0)
                    nodeIndex = 1;

                //Connect with xml document first
                XmlDocument doc = new XmlDocument();
                doc.Load(Utility.GetVariable("APIFilePath"));

                //Gets the list of nodes having specified attribute name.
                XmlNodeList list = doc.GetElementsByTagName(attributeName);

                int cnt = ReturnCountofXmlNodes(attributeName);

                //When total count of attributes was found to be less than expected
                if (cnt < nodeIndex)
                {
                    blnNodeFound = false;
                    Property.Remarks = "Total instances of attribute '" + attributeName + "' were less than requested index '" +
                                        nodeIndex + "'. First occurance has been stored to variable '" + attributeName + "'. ";
                    nodeIndex = 1;
                }

                //storing the value to runtimedic.
                if (valueLookingFor.Equals(string.Empty))
                {
                    blnNodeFound = true;
                    Utility.SetVariable(attributeName, list.Item(nodeIndex - 1).InnerText.ToString().Trim());
                }
                //else loop will be hit only when we are looking for a specific value of attribute anywhere in the document
                else
                {
                    //Iterate through each node and check if expected value is found
                    for (int i = 0; i < cnt; i++)
                    {
                        if (list.Item(i).InnerText.ToLower().Trim().Equals(valueLookingFor.ToLower().Trim()))
                        {
                            blnNodeFound = true;
                            nodeIndex = i + 1;
                            Utility.SetVariable(varNameForAttributeLocation, nodeIndex.ToString());
                            break;
                        }
                    }

                    //If node was found, store its location to a variable also for later usage
                    if (blnNodeFound)
                    {
                        Property.Remarks = "Value '" + valueLookingFor + "' of attribute '" + attributeName +
                                           "' was found at position :" + nodeIndex + " in api response. " +
                                           "This position was saved to variable " + varNameForAttributeLocation + ".";
                        return true;
                    }
                    else
                    {
                        Property.Remarks = "Value '" + valueLookingFor + "' of attribute '" + attributeName +
                                           "' was NOT found anywhere in api response";
                        return false;
                    }
                }

                Property.Remarks = Property.Remarks + "Stored value:= " + Utility.GetVariable(attributeName);
                return true;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        
        /// <summary>
        /// Web Api calls goes here.
        /// </summary>
        /// <param name="Header">string: string containing pipe seperated value for content type and authorization.</param>
        /// <param name="Format">string: format in which response would be saved.</param>
        /// <param name="Method">string: defines the type of reuest.</param>
        /// <param name="url">string: API url</param>
        /// <param name="Body">string: Body to be send with request.</param>
        /// <returns>bool: true OR false</returns>

        /// <summary>
        /// Count number of nodes present in XML response with specified node name 
        /// and store the value in runtimedic dictionary with key=nodecount.
        /// </summary>
        /// <param name="attributeNodeName">string : Node name to count. </param>
        /// <returns>bool with true OR false</returns>
        /// Reused ReadXmlAttribute to count nodes also
        public static bool CountXMLNodes(string attributeNodeName)
        { 
            try
            {
                int cnt = ReturnCountofXmlNodes(attributeNodeName);
                if (cnt == 0)
                {
                    Property.Remarks = "Count of the "+attributeNodeName+ "is found to be 0";
                    return true;
                }
                if (readXmlAttribute(attributeNodeName))
                {
                    
                    Property.Remarks= "Total " + Utility.GetVariable(attributeNodeName + "Count") + " instances of attribute named '" +
                                        attributeNodeName + " were found and stored to variable '" + attributeNodeName + "Count'";  //  Kafaltiya: Updated Remark to display the Variable name on which the node count will be stored.
                    return true;
                }
                else
                {
                    //Property.Remarks is being set in readXmlAttribute method itself when attribute could not be located
                    return false;
                }
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// verify XMLnode count for specified node.
        /// </summary>
        /// <param name="data">String : Combination of node and count to verify</param>
        /// <returns>bool : true OR false</returns>
        public static bool verifyxmlnodecount(string attributeName, string expectedCount)
        {
            try
            {
                if (CountXMLNodes(attributeName))
                {
                    string actualCount = Utility.GetVariable(attributeName + "Count").Trim();
                    if (actualCount.Equals(expectedCount))
                    {
                        Property.Remarks = string.Empty;
                        return true;
                    }
                    else
                    {
                        Property.Remarks = "Actual instance count '" +
                        actualCount + "' of attribute '" + attributeName + "' does not match its expected count '" + expectedCount + "'.";
                        return false;
                    }
                }
                else
                {
                    Property.Remarks = "Count of attribute " + attributeName + "" + "is" + "0";
                    //Property.Remarks is being set in readXmlAttribute method itself when attribute could not be located
                    return true;
                }
            }
            catch (Exception e)
            {
                throw e;
            }
        }
    }
}

/***************************************************************************
** Application: Krypton: ThinkSys Automation Generic System
** File Name: TestDriver.APIRelatedMethods.cs
** Copyright 2012, ThinkSys Software
** All rights reserved
** Developed by: ThinkSys Software
** Description: Contains the functionalities related to API calls.
*****************************************************************************/

using System;
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
        public static bool RequestWebService(string requestHeader, string responseFormat, string requestMethod, string requestUrl, string requestBody)
        {
            //name of the file, which would contain exact response returned by api call
            string apiResponseFileName = Property.ApiXmlFile + Property.StepNumber + "." + responseFormat;

            //name of the file, which would contain xml equivalent of the api response, may not be converted always
            string xmlFileName = Property.ApiXmlFile + Property.StepNumber + ".xml";

            //Final location where API resonse will be saved
            string xmlFilePath = Property.ResultsSourcePath + "\\" + apiResponseFileName;

            try
            {
                if (File.Exists(xmlFilePath))
                {
                    File.Delete(xmlFilePath);
                }

                //Create an object to request webservice
                HttpWebRequest webServiceRequest = (HttpWebRequest)WebRequest.Create(requestUrl);

                //Assign headers to webservice object
                string[] arrHeaderList = requestHeader.Split(Property.Seprator);
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
                        case "user-agent":
                            webServiceRequest.UserAgent  = strHeaderValue;// .Headers.Add(HttpRequestHeader.UserAgent,strHeaderValue);
                            break;
                        default:
                            if (!WebHeaderCollection.IsRestricted(strHeaderName))
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
                    Property.Remarks = requestBody;
                }
                WebResponse responseXml = webServiceRequest.GetResponse();
                //var headers = responseXml.Headers;
                //Collect returns from webservice
                string statusCode = ((HttpWebResponse)responseXml).StatusCode.ToString();
                string statusDescription = ((HttpWebResponse)responseXml).StatusDescription;
                Console.WriteLine("API call response: " + statusCode + "\t" + statusDescription);
                String responseFromServer = null;
                using(Stream responseData = responseXml.GetResponseStream())
                    if (responseData != null)
                        using (StreamReader reader = new StreamReader(responseData, true)) 
                        {
                            responseFromServer = reader.ReadToEnd();
                            responseFromServer.Replace('\\', ' ');
                        }

                //Parse API response
                //Store return data to a file and created a variable also
                File.WriteAllText(xmlFilePath, responseFromServer);

                Utility.SetVariable("APIFilePath", xmlFilePath);

                Property.Attachments = apiResponseFileName;

                //To handle Json which is returned without object and the url having "api"
                #region Converting json to xml in case json was the original format
                if (responseFormat.ToLower().Equals("json") || requestUrl.ToLower().Contains("json") || requestUrl.ToLower().Contains("api"))
                {
                    try
                    {
                        //Converting json to xml
                        XmlNode jsonToXml = null;
                        string jsonString = File.ReadAllText(xmlFilePath);
                        if (jsonString==string.Empty)
                        {
                            StoreResponseHeader(responseXml);
                        }
                        else
                        {
                            jsonToXml = JsonConvert.DeserializeXmlNode("{\"root\":" + jsonString + "}", Property.ProductName);
                            //New path exclusively for xml file, 
                            //this is achieved by replacing json file name with xml file name in full length path.
                            xmlFilePath = xmlFilePath.Replace(apiResponseFileName, xmlFileName);
                            //Write converted xml to the document
                            XmlDocument convertedXmlDoc = new XmlDocument();
                            convertedXmlDoc.CreateXmlDeclaration("1.0", string.Empty, string.Empty);
                            convertedXmlDoc.LoadXml(jsonToXml.OuterXml);
                            convertedXmlDoc.Save(xmlFilePath);
                            //Going forward xml will be used to validate and read attributes
                            Utility.SetVariable("APIFilePath", xmlFilePath);
                            StoreResponseHeader(responseXml);
                        }
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
                throw;
            }
        }


        public static bool VerifyXmlHeader(string headerField, string expectedValue)
        {
            try
            {
                string headerFilePath = Property.ResultsSourcePath + "\\" + Property.ApiXmlHeaderFile;
                using (StreamReader sr = new StreamReader(headerFilePath))
                {
                    string line = string.Empty;
                    while ((line=sr.ReadLine()) != null)
                    {
                        string[] headerInfo = line.Split(':');
                        if (headerInfo[0].Trim().Equals(headerField.Trim(),StringComparison.OrdinalIgnoreCase))
                        {
                            if (expectedValue.Trim() == headerInfo[1].ToString().Trim())
                                return true;
                        }
                    }
                }
            }
            catch (FileNotFoundException e)
            {
                Console.WriteLine("ResponseHeader Not Found.");
                return false;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return false;
        }

        /// <summary>
        /// This method use to store the response header to the file.
        /// </summary>
        private static void StoreResponseHeader(WebResponse response)
        {
            HttpWebResponse responseheader = (HttpWebResponse)response;
            string xmlFilePath = Property.ResultsSourcePath + "\\" + Property.ApiXmlHeaderFile;
            using (StreamWriter sw = new StreamWriter(xmlFilePath))
            {
                sw.WriteLine("ResponseHeader : Found");
                sw.WriteLine("protocolversion: {0}", responseheader.ProtocolVersion);
                int iStatusCode = (int)responseheader.StatusCode;
                sw.WriteLine("statuscode: {0}", iStatusCode.ToString());
                sw.WriteLine("statusdescription: {0}", responseheader.StatusDescription);
                sw.WriteLine("contentencoding: {0}", responseheader.ContentEncoding);
                sw.WriteLine("contentlength: {0}", responseheader.ContentLength);
                sw.WriteLine("contenttype: {0}", responseheader.ContentType);
                sw.WriteLine("lastmodified: {0}", responseheader.LastModified);
                sw.WriteLine("server: {0}", responseheader.Server);
            }
        }

        private static bool IsXmlConvertable(string responseFromServer)
        {
            bool isLoadSuccessful = false;
            try
            {
                XmlDocument convertedXmlDoc = new XmlDocument();
                convertedXmlDoc.CreateXmlDeclaration("1.0", string.Empty, string.Empty);
                convertedXmlDoc.LoadXml(responseFromServer);
                isLoadSuccessful = true;
            }
            catch (Exception ex)
            {
                // ignored
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
        public static bool DownloadFile(string webLocation, string localFileName = "FileDownload.file")
        {
            localFileName = localFileName.Trim().Replace(".", Property.StepNumber + ".");

            try
            {
                //Delete any existing file with same name
                string xmlFilePath = Property.ResultsSourcePath + "\\" + localFileName;
                string locationVariable = "DownloadedFile";

                if (File.Exists(xmlFilePath))
                {
                    File.Delete(xmlFilePath);
                }
                //Download file from web resources and store on local disk
                try
                {
                    WebClient wc = new WebClient();
                    wc.DownloadFile(webLocation, xmlFilePath);
                    Utility.SetVariable(locationVariable, xmlFilePath);

                    Property.Remarks = "File has been downloaded and its location has been store in variable '" +
                                              locationVariable + "'.";
                }
                catch (Exception fileDownloadError)
                {
                    Property.Remarks = "Unable to download file: " + fileDownloadError.Message;
                    return false;
                }

                Property.Attachments = localFileName;
                return true;
            }
            catch (Exception e)
            {
                Property.Remarks = "Unable to download file: " + e.Message;
                return false;
            }
        }

        public static bool UploadFiletoFtp(string fileName, Uri uploadUrl, string user, string pswd)
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
                while (true)
                {
                    var bytesRead = fileStream.Read(buffer, 0, buffer.Length);
                    if (bytesRead == 0)
                        break;
                    requestStream.Write(buffer, 0, bytesRead);
                }

                requestStream.Close();
                uploadResponse = (FtpWebResponse)uploadRequest.GetResponse();
                Property.Remarks = "File :'" + fileName + "' has been uploaded to '" + uploadUrl + "'.";
            }
            catch (UriFormatException ex)
            {
                Property.Remarks = "Error: " + ex.Message;
                return false;
            }
            catch (IOException ex)
            {
                Property.Remarks = "Error: " + ex.Message;
                return false;
            }
            catch (WebException ex)
            {
                Property.Remarks = "Error: " + ex.Message;
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
        /// <param name="nodeIndex"></param>
        /// <returns>bool value true OR false</returns>
        public static bool VerifyXmlAttribute(string attributeName, string expectedValue, int nodeIndex = 1)
        {

            FileInfo fileInfo = new FileInfo(Utility.GetVariable("APIFilePath"));

            try
            {
                //replacing variable in case variable is used instead of value.
                expectedValue = Utility.ReplaceVariablesInString(expectedValue);
                //Read the value for specified node.
                bool readAttribute = ReadXmlAttribute(attributeName, nodeIndex);
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

                if (Utility.DoKeywordMatch(expectedValue, actualValue))
                {
                    return true;
                }
                if (Property.SnapshotOption.Equals("on failure", StringComparison.OrdinalIgnoreCase) && Property.DebugMode.Equals("false", StringComparison.OrdinalIgnoreCase))
                {
                    Property.Attachments = fileInfo.Name;
                }
                Property.Remarks = "Actual value '" +
                                   actualValue + "' of attribute '" + attributeName + "' does not match its expected value '" + expectedValue + "'.";
                return false;
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
        public static bool ReadXmlAttribute(string attributeName, int nodeIndex = 1, string valueLookingFor = "")
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
                    var xmlNode = list.Item(nodeIndex - 1);
                    if (xmlNode != null)
                        Utility.SetVariable(attributeName, xmlNode.InnerText.ToString().Trim());
                }
                //else loop will be hit only when we are looking for a specific value of attribute anywhere in the document
                else
                {
                    //Iterate through each node and check if expected value is found
                    for (int i = 0; i < cnt; i++)
                    {
                        var xmlNode = list.Item(i);
                        if (xmlNode != null && xmlNode.InnerText.ToLower().Trim().Equals(valueLookingFor.ToLower().Trim()))
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
                    Property.Remarks = "Value '" + valueLookingFor + "' of attribute '" + attributeName +
                                       "' was NOT found anywhere in api response";
                    return false;
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
        /// <returns>bool: true OR false</returns>
        /// <summary>
        /// Count number of nodes present in XML response with specified node name 
        /// and store the value in runtimedic dictionary with key=nodecount.
        /// </summary>
        /// <param name="attributeNodeName">string : Node name to count. </param>
        /// <returns>bool with true OR false</returns>
        /// Reused ReadXmlAttribute to count nodes also
        public static bool CountXmlNodes(string attributeNodeName)
        {
            int cnt = ReturnCountofXmlNodes(attributeNodeName);
            if (cnt == 0)
            {
                Property.Remarks = "Count of the " + attributeNodeName + "is found to be 0";
                return true;
            }
            if (ReadXmlAttribute(attributeNodeName))
            {

                Property.Remarks = "Total " + Utility.GetVariable(attributeNodeName + "Count") + " instances of attribute named '" +
                                   attributeNodeName + " were found and stored to variable '" + attributeNodeName + "Count'";
                return true;
            }
            //Property.Remarks is being set in readXmlAttribute method itself when attribute could not be located
            return false;
        }

        /// <summary>
        /// verify XMLnode count for specified node.
        /// </summary>
        /// <param name="attributeName"></param>
        /// <param name="expectedCount"></param>
        /// <returns>bool : true OR false</returns>
        public static bool Verifyxmlnodecount(string attributeName, string expectedCount)
        {
            try
            {
                if (CountXmlNodes(attributeName))
                {
                    string actualCount = Utility.GetVariable(attributeName + "Count").Trim();
                    if (actualCount.Equals(expectedCount))
                    {
                        Property.Remarks = string.Empty;
                        return true;
                    }
                    Property.Remarks = "Actual instance count '" +
                                       actualCount + "' of attribute '" + attributeName + "' does not match its expected count '" + expectedCount + "'.";
                    return false;
                }
                Property.Remarks = "Count of attribute " + attributeName + "" + "is" + "0";
                //Property.Remarks is being set in readXmlAttribute method itself when attribute could not be located
                return true;
            }
            catch (Exception e)
            {
                throw e;
            }
        }
    }
}

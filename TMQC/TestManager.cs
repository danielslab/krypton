/***************************************************************************
** Application: Krypton: ThinkSys Automation Generic System
** File Name: TestManager.cs
** Copyright 2012, ThinkSys Software
** All rights reserved
** Developed by: ThinkSys Software
** Description: File to Interact with Quality Center.
*****************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using TDAPIOLELib;


namespace Krypton
{
    public class TestManager
    {

        //Assign all Quality Center related variables here. These will be used in methods written below
        private static string qcServerName = "da0itas049";
        private static string qcServer = "http://" + qcServerName + "/qcbin";
        private static string strDomainName = "DEFAULT";
        private static string strProjectName = "Match_Domestic_Automation";
        private static string strQCUserName = "qcuserv1";
        private static string strQCUserPassword = "Uxr<Fx5%";
        private static string LocalFileLocation = string.Empty;

        /// <summary>
        /// This method will download an attachment from a folder in Quality Center
        /// </summary>
        /// <param name="qcFolderLocation">Folder Location from where attachment needs to be downloaded</param>
        /// <param name="strFileName">Name of the file that needs to be downloaded from specified location. 
        /// Name should be specified along with extenstion </param>
        /// <returns>Location of local (usually temp) path where file has been downloaded. Returns empty string if file could not be downloaded.</returns>
        public string DownloadAttachment(string qcFolderLocation, string strFileName)
        {
            //Connect with Quality Center
            TDConnection qctd = new TDConnection();
            qctd.InitConnectionEx(qcServer);
            qctd.ConnectProjectEx(strDomainName, strProjectName, strQCUserName, strQCUserPassword);

            if (qctd.Connected)
            {
                //Define objects that will be used to download files
                SubjectNode otaSysTreeNode = new SubjectNode();
                AttachmentFactory otaAttachmentFactory = new AttachmentFactory();
                TDFilter otaAttachmentFilter = new TDFilter();
                List otaAttachmentList = new List();
                ExtendedStorage attStorage = new ExtendedStorage();

                otaSysTreeNode = qctd.TreeManager.NodeByPath(qcFolderLocation);     //Returns node object from test plan in Quality Center
                otaAttachmentFactory = otaSysTreeNode.Attachments();                //Returns all attachments for the folder in QC
                otaAttachmentFilter = otaAttachmentFactory.Filter();                //Can be used to filter list of attachments
                otaAttachmentList = otaAttachmentFilter.NewList();                  //Creates list of attached files

                //Check if there is any attachment available for the specified folder
                if (otaAttachmentList.Count > 0)
                {
                    foreach (Attachment otaAttachment in otaAttachmentList)
                    {
                        //Check if file names are same
                        if (otaAttachment.FileName.ToLower() == strFileName.ToLower())
                        {
                            attStorage = otaAttachment.AttachmentStorage();
                            LocalFileLocation = otaAttachment.DirectLink;

                            //Load method will download file to local workstation. true to used for synchronised download.
                            attStorage.Load(LocalFileLocation, true);

                            //Client path refers to local path where file has been downloaded
                            LocalFileLocation = attStorage.ClientPath;
                            break;
                        }
                    }
                }
            }

            //Return empty string if connection to QC was not successfull.
            else
            {
                LocalFileLocation = string.Empty;
            }

            return LocalFileLocation;
        }




        public static bool UploadAttachment(string resultFilesList, string resultDestinationPath)
        {
            return true;
        }
    }
}
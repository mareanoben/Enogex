using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint;
using System.IO;
using System.Collections;
using System.Data.SqlClient;

namespace FileLoaderTimerJob
{
   // enum siteColl { okc, tul };
    //enum webs { cust1, cust2, cust3, cust4, cust5, cust6, cust7, cust8, cust9, cust10 };

    public class FileLoader : SPJobDefinition
    {

        public const string JOB_DEFINITION_NAME = "File Loader TIMER_JOB";
        public const string JOB_DEFINITION_TITLE = "FileLoader";
        
        public static string logFilePath = DateTime.Now.ToString();

        /// <summary>
        /// Initializes a new instance of the <see cref="FileLoader"/> class.
        /// </summary>
        public FileLoader()
        {
            Title = JOB_DEFINITION_TITLE;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="FileLoader"/> class.
        /// </summary>
        /// <param name="webApplication">The web application.</param>
        public FileLoader(SPWebApplication webApplication)
            : base(JOB_DEFINITION_NAME, webApplication, null, SPJobLockType.Job)
        {
            Title = JOB_DEFINITION_TITLE;
        }

        public override void Execute(Guid targetInstanceId)
        {
            logFilePath = formatDate(DateTime.Now.ToString());


            try
            {

                //SPWebApplication webApp = SPWebService.ContentService.WebApplications.Cast<SPWebApplication>().Where(w => w.GetResponseUri(SPUrlZone.Default).AbsoluteUri == "http://sp10-dev:2000/").FirstOrDefault();

                //// foreach (SPWebApplication webApp in SPWebService.ContentService.WebApplications)
                //if (webApp != null)
                //{

                   // string webAppUrl = webApp.GetResponseUri(SPUrlZone.Default).AbsoluteUri;

                   // string siteCol = webAppUrl + "sites/okc";
                    string siteCol = "http://devecportal";
                    string[] filePaths = Directory.GetFiles(@"c:\Enogex\", "*", SearchOption.AllDirectories);
                    bool isWebExist = false;

                    if (filePaths.Length != 0)
                    {
                        foreach (string filePath in filePaths)
                        {

                            string fileName = Path.GetFileNameWithoutExtension(filePath);
                            string customerName = getCustomerName(fileName);
                            using (SPSite site = new SPSite(siteCol)) 
                            { 
                             isWebExist = webExists(site,customerName);
                            }
                           
                            string documentLibrary = getLibraryName(filePath);
                            if (documentLibrary != null && isWebExist)
                            {
                                fileUpload(filePath, documentLibrary, customerName, siteCol);
                               
                            }
                            else
                                fileNotUploaded(filePath);
                        }

                   

                    }


                    // }
                //}
            }
            catch
            { }




        }
        private bool webExists(SPSite site, string web_)
        {
            return site.AllWebs.Cast<SPWeb>().Any(web => string.Equals(web.Name, web_));
        }

        private string formatDate(string path)
        {
            string formattedDate = path.Replace(@"/", "_").Replace(" ", "--").Replace(":", "_");
            try
            {
                string date = path.Replace(@"/", "_").Replace(" ", "--");
                int colon = date.IndexOf(":");
                int dash2 = date.LastIndexOf("--");

                string pmAm = date.Substring(dash2 + 2);

                if (colon != -1 && dash2 != -1)
                    formattedDate = date.Substring(0, colon) + pmAm;
            }
            catch
            {
                formattedDate = path.Replace(@"/", "_").Replace(" ", "--").Replace(":", "_");
            }
            return formattedDate;
        }
     
        static string renameFile(string fileName)
        {
            string newFileName = string.Empty; ;

            //cust_1_;FileName;ContentType.txt
            try
            {
                string[] array = fileName.Split(';');
                if (array[4] != null)
                    newFileName = array[6];
            }
            catch
            {
                newFileName = fileName;
            }
            return newFileName;
        }
        private static string getLibraryName(string filePath)
        {
            string docName = string.Empty;
            #region old code
            //try
            //{
            //    int positionOfSemicolon1 = filePath.IndexOf(";");
            //    int positionOfSemicolon2 = filePath.IndexOf(";", positionOfSemicolon1 + 1);
            //    int positionOfSemicolon3 = filePath.LastIndexOf(";");

            //    if (positionOfSemicolon1 != -1 && positionOfSemicolon2 != -1 && positionOfSemicolon3 != -1)
            //    {
            //        docName = filePath.Substring(positionOfSemicolon2 + 1, (positionOfSemicolon3 - positionOfSemicolon2) - 1);

            //    }
            //    else
            //    {
            //        docName = filePath;
            //    }
            //}
            //catch {
            //    docName = filePath;
            //}
            #endregion
            try
            {
                int positionOfSlash1 = filePath.IndexOf(@"\");
                int positionOfSlash2 = filePath.IndexOf(@"\", positionOfSlash1 + 1);
                int positionOfSlash3 = filePath.LastIndexOf(@"\");

                if (positionOfSlash1 != -1 && positionOfSlash2 != -1 && positionOfSlash3 != -1)
                    docName = filePath.Substring(positionOfSlash2 + 1, (positionOfSlash3 - positionOfSlash2) - 1);

            }
            catch
            {
                docName = filePath;
            }
            return docName;

        }
        private static string getCustomerName(string fileName)
        {
            string systID = string.Empty;
            string custID = string.Empty;
            string cust = string.Empty;
            string customerName = string.Empty;
            try
            {
                string[] arr = fileName.Split(';');
                if (arr[0] != null && arr[1] != null && arr[2] != null)
                {
                    systID = arr[0];
                    custID = arr[1];
                    cust = arr[2];
                }


               
                else
                {
                    customerName = fileName;
                }

                using (SPSite site = new SPSite("http://devecportal"))
                {
                    using (SPWeb web = site.OpenWeb())
                    {
                        SPList list = web.Lists.TryGetList("CustomerMap");
                        if (list != null) 
                        {
                            SPListItemCollection listCol = list.Items;
                            foreach (SPListItem lst in listCol)
                            {
                                if (lst["SystemID"].ToString() == systID && lst["CustomerID"].ToString() == custID && lst["EnogexCustomerID"].ToString() == cust)
                                {
                                    customerName = lst["CustomerURL"].ToString();
                                    break;
                                    
                                }
                            }
                            //SPListItem item = listCol.Cast<SPListItem>().Where(it => it["CustomerID"] == custID && it["EnogexCustID"] == cust).FirstOrDefault();
                            //customerName = item["CustomerURL"].ToString();
                        
                        }
                    }
                }
                
            }

            catch
            {
                customerName = fileName;
            }
            return customerName;
        }
        private static string getContentTypeName(string fileName)
        {
            string contTypeName = string.Empty;
            
            try
            {
                string[] arr = fileName.Split(';');

                if (arr[3] != null)
                    contTypeName = arr[2];
                
                else
                {
                    contTypeName = fileName;
                }
            }
            catch
            {
                contTypeName = fileName;
            }
            return contTypeName;
        }
        private static string getReportDate(string fileName)
        {
            string reportDate = string.Empty;
            //cust_1_;FileName2;FolderName;exploration
            try
            {
               string[] arr = fileName.Split(';');
               if (arr[5] != null)
                   reportDate = arr[7];

               else
               {
                   reportDate = fileName;
               }
            }
            catch
            {
                reportDate = fileName;
            }
            return reportDate;
        }


        #region GetContentType

        static string getContentType(string fileName, string documentLibrary, string contentTypeName, SPSite site)
        {
            string typeId = string.Empty;
            using (SPWeb web = site.OpenWeb())
            {
                try
                {
                    SPContentTypeCollection cTypeColl = web.ContentTypes;
                   
                    

                    foreach(SPContentType type in cTypeColl)
                    {
                        if (type.Name.ToLower() == contentTypeName.ToString().ToLower() && type.Group == "Customer Reports")
                            typeId = type.Id.ToString();
                    }
                    //List<SPContentType> contColl = web.ContentTypes.Cast<SPContentType>().Where(c => c.Group == "Customer Portal Reports").ToList();

                    //return contColl.Where(c => c.Name.ToLower() == contentTypeName.ToString().ToLower()).FirstOrDefault().Id.ToString();

                   
                }
                catch { }

            }
            return typeId;
        }


        #endregion

        private static void fileUpload(string filePath, string docLibrary, string siteUrl, string siteColUrl)
        {

            string contentTypeName = "Document";
            List<string> logInfo = new List<string>();
            using (SPSite site = new SPSite(siteColUrl))
            {
                using (SPWeb web = site.OpenWeb(siteUrl))
                {
                    if (System.IO.File.Exists(@filePath))
                    {
                        try
                        {
                           

                            SPList list = web.Lists.TryGetList(docLibrary);

                            if (list != null)
                            {
                                SPFolder myLibrary = web.Folders[docLibrary];
                                Boolean replaceExistingFiles = true;
                                string fileName = System.IO.Path.GetFileName(@filePath);
                                FileStream fileStream = File.OpenRead(@filePath);

                                contentTypeName = getContentTypeName(fileName);
                                string customerName = getCustomerName(fileName);

                                string renamedFile = renameFile(fileName);
                                string cTypeId = getContentType(fileName, docLibrary, contentTypeName, site);


                                // Upload document                 
                                SPFile spfile = myLibrary.Files.Add(renamedFile, fileStream, replaceExistingFiles);
                                SPListItem item = spfile.Item;



                                item["ContentTypeId"] = cTypeId;

                                item.Update();

                                
                                //item["Value1"] = getValueOne(fileName);
                                //item["Value2"] = getValueTwo(fileName);
                                //item["Value3"] = getValueThree(fileName);

                                item.Update();

                                myLibrary.Update();

                                string extention = ".";
                                string modifiedBy = "NA";

                                string editor = item["Editor"].ToString();
                                int positionOfSlash = editor.LastIndexOf("#");
                                int positionOfExt = fileName.LastIndexOf(".");

                                if (positionOfExt != -1)
                                    extention = fileName.Substring(positionOfExt);

                                if (positionOfSlash != -1)
                                    modifiedBy = editor.Substring(positionOfSlash + 1);

                                //CsvFileWriter log = new CsvFileWriter();
                                //log.csvWrite(logInfo);


                                bool fileUploaded = checkFileUploaded(spfile, item.Url, web);
                                if (fileUploaded == false)
                                {
                                    //Log
                                    fileNotUploaded(filePath);

                                }
                                else
                                {
                                    //Log

                                    #region logInfo List<string> old code
                                    // @"C:\LoadedFiles.csv";
                                    string path = @"C:\Log\Success\LoadedFiles_" + logFilePath.Replace(@"/", "_").Replace(":", "_") + ".csv";
                                    List<string> fileInfo = new List<string>();
                                    fileInfo.Add(customerName);
                                    fileInfo.Add(docLibrary);
                                    fileInfo.Add(contentTypeName);
                                    fileInfo.Add(item["Name"].ToString());
                                    fileInfo.Add(item["Modified"].ToString());
                                    fileInfo.Add(DateTime.Now.ToString());
                                    fileInfo.Add(modifiedBy);
                                    fileInfo.Add(extention);

                                    CsvWriter.CreateCsvFile(fileInfo, path);


                                    #endregion

                                    deleteFile(fileUploaded, fileName, docLibrary, filePath, fileStream);
                                }
                            }

                              //Create a new Document Library if not exist
                            else 
                            {

                                SPDocLibrary docLib = new SPDocLibrary();

                                docLibrary = docLib.createNewDocumentLibrary(site, siteUrl, docLibrary);

                                SPFolder library = web.Folders[docLibrary];

                                Boolean replaceExistingFiles = true;
                                String fileName = System.IO.Path.GetFileName(@filePath);
                                FileStream fileStream = File.OpenRead(@filePath);

                                contentTypeName = getContentTypeName(fileName);
                                string cTypeId = docLib.setContentType(site, docLibrary, contentTypeName);

                                string renamedFile = renameFile(fileName);

                                // Upload document                 
                                SPFile spfile = library.Files.Add(renamedFile, fileStream, replaceExistingFiles);
                                SPListItem item = spfile.Item;
                                item["ContentTypeId"] = cTypeId;

                                item.Update();
                                library.Update();
                                bool fileUploaded = checkFileUploaded(spfile, item.Url, web);

                                if (fileUploaded == false)
                                {
                                    //ToDo: Log

                                }
                                else
                                {
                                    //ToDo: Log
                                    deleteFile(fileUploaded, fileName, docLibrary, filePath, fileStream);
                                }
                            
                            }
                        }
                        catch
                        {

                        }

                    }

                }
            }
        }

        private static void fileNotUploaded(string filePath)
        {
            List<string> unloadedFiles = new List<string>();
            string path = @"C:\Log\Failed\NotLoaded_" + logFilePath.Replace(@"/", "_").Replace(":", "_") + ".csv";

            string fileName = System.IO.Path.GetFileName(@filePath);

            string extention = ".";

            int positionOfExt = fileName.LastIndexOf(".");
            string file = renameFile(fileName);

            if (positionOfExt != -1)
                extention = fileName.Substring(positionOfExt);

            unloadedFiles.Add(file);
            unloadedFiles.Add(DateTime.Now.ToString());
            unloadedFiles.Add(extention);

            CsvWriter.CreateCsvFile(unloadedFiles, path);

            deleteUnloadedFile(fileName, filePath);
           

        }
      

        private static bool checkFileUploaded(SPFile spfile, string filePath, SPWeb web)
        {
            bool fileUploaded = false;
            spfile = web.GetFile(filePath);
            if (spfile.Exists)
            {
                fileUploaded = true;
            }
            return fileUploaded;
        }

        private static void deleteFile(bool fileUploaded, string fileName,string docLibrary, string filePath, FileStream fileStream)
        {
            if (fileUploaded == true)
            {
                try
                {

                    if (!File.Exists(@"c:\Archive\Loaded\" + docLibrary+ @"\" + fileName))
                    {
                        File.Copy(@filePath, @"c:\Archive\Loaded\" + docLibrary + @"\" + fileName);
                    }
                    fileStream.Close();
                    File.Delete(@filePath);
                }
                catch { }
            }


        }

        private static void deleteUnloadedFile(string fileName, string filePath)
        {

            try
            {
                FileStream fileStream = File.OpenRead(@filePath);
                if (!File.Exists(@"c:\Archive\Unloaded\" + fileName))
                {
                    File.Copy(@filePath, @"c:\Archive\Unloaded\" + fileName);
                }
                fileStream.Close();
                File.Delete(@filePath);
            }
            catch { }



        }

    }
}

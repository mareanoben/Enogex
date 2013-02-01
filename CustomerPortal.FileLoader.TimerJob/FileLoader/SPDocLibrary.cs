using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint;

namespace CustomerPortal.FileLoader.TimerJob.FileLoader
{

    enum contTypes
    {
        General
    };
    class SPDocLibrary
    {
        List<string> webs = new List<string>();
        public string createNewDocumentLibrary(SPSite site, string _web, string docLib)
        {
            //foreach (webs web_ in Enum.GetValues(typeof(webs)))
            //{
            try
            {
                bool webExist = webExists(site, _web);

                if (webExist == true)
                {

                    using (SPWeb web = site.OpenWeb(_web))
                    {
                        SPList list = web.Lists.TryGetList(docLib);
                        if (list == null)
                        {
                            SPListTemplateType tempType = SPListTemplateType.DocumentLibrary;
                            web.Lists.Add(docLib, null, tempType);
                            return docLib;
                        }
                    }
                }
                //}
            }
            catch { }
            return docLib;
        }
        public string setContentType(SPSite site, string docLib, string cont)
        {

            using (SPWeb web = site.OpenWeb())
            {
                SPList list = web.Lists.TryGetList("CustomerMap");
                if (list != null)
                {
                    SPListItemCollection listCol = list.Items;
                    foreach (SPListItem lst in listCol)
                    {
                        webs.Add(lst["CustomerURL"].ToString());


                    }
                    //SPListItem item = listCol.Cast<SPListItem>().Where(it => it["CustomerID"] == custID && it["EnogexCustID"] == cust).FirstOrDefault();
                    //customerName = item["CustomerURL"].ToString();

                }
            }



            string contentTypeGroup = "Customer Portal Reports";
            string contId = string.Empty;
            try
            {
                using (SPWeb mainSite = site.OpenWeb())
                {
                    //foreach (contTypes cont in Enum.GetValues(typeof(contTypes)))
                    //{


                    SPContentType documentCType = mainSite.AvailableContentTypes[SPBuiltInContentTypeId.Document];

                    List<SPContentType> contColl = mainSite.ContentTypes.Cast<SPContentType>().Where(c => c.Group == "Customer Reports").ToList();
                    //SPContentType availableCType = mainSite.AvailableContentTypes[cont];
                    SPContentType availableCType = contColl.Where(c => c.Name == cont).FirstOrDefault();
                    if (availableCType != null)
                    {
                        availableCType = mainSite.ContentTypes[availableCType.Id];
                        ensureContentTypeAddedToList(site, availableCType, docLib);
                        contId = availableCType.Id.ToString();

                    }

                   //ToDo: We're not creating content Types authomatically
                    else
                    {

                        SPContentType newCType = new SPContentType(documentCType, mainSite.ContentTypes, cont);
                        mainSite.ContentTypes.Add(newCType);
                        contId = newCType.Id.ToString();
                        newCType = mainSite.ContentTypes[newCType.Id];
                        newCType.Group = contentTypeGroup;
                        attachContentType(site, newCType, docLib);

                        newCType.Update();

                    }

                }
            }
            catch { }

            return contId;

        }



        private void ensureContentTypeAddedToList(SPSite site, SPContentType availableContType, string docLib)
        {

            foreach (string web_ in webs)
            {
                bool webExist = webExists(site, web_.ToString());

                if (webExist == true)
                {

                    using (SPWeb web = site.OpenWeb(web_.ToString()))
                    {
                        SPList list = web.Lists.TryGetList(docLib);
                        if (list == null)
                        {
                            SPListTemplateType tempType = SPListTemplateType.DocumentLibrary;
                            web.Lists.Add(docLib, null, tempType);
                            list = web.Lists.TryGetList(docLib);
                            list.OnQuickLaunch = true;
                            list.ContentTypesEnabled = true;

                            SPContentTypeCollection cTypeColl = list.ContentTypes;
                            List<string> contTypeLists = new List<string>();

                            for (int i = 0; i < cTypeColl.Count; i++)
                            {
                                contTypeLists.Add(cTypeColl[i].Name.ToString());

                            }

                            if (!contTypeLists.Contains(availableContType.Name.ToString()))
                            {
                                list.ContentTypesEnabled = true;
                                list.ContentTypes.Add(availableContType);

                            }


                            list.Update();
                        }
                        else
                        {
                            SPContentTypeCollection cTypeColl = list.ContentTypes;
                            List<string> contTypeLists = new List<string>();

                            for (int i = 0; i < cTypeColl.Count; i++)
                            {
                                contTypeLists.Add(cTypeColl[i].Name.ToString());

                            }

                            if (!contTypeLists.Contains(availableContType.Name.ToString()))
                            {
                                list.ContentTypesEnabled = true;
                                list.ContentTypes.Add(availableContType);
                                list.Update();

                            }

                        }
                    }
                }
                else
                {
                    continue;
                }
            }
        }




        private bool webExists(SPSite site, string web_)
        {
            return site.AllWebs.Cast<SPWeb>().Any(web => string.Equals(web.Name, web_));
        }




        private void attachContentType(SPSite site, SPContentType contType, string docLib)
        {

            foreach (string web_ in webs)
            {
                bool webExist = webExists(site, web_.ToString());

                if (webExist == true)
                {
                    using (SPWeb web = site.OpenWeb(web_.ToString()))
                    {
                        SPList list = web.Lists.TryGetList(docLib);
                        if (list == null)
                        {
                            SPListTemplateType tempType = SPListTemplateType.DocumentLibrary;
                            web.Lists.Add(docLib, null, tempType);
                            list = web.Lists.TryGetList(docLib);
                            list.OnQuickLaunch = true;
                            list.ContentTypesEnabled = true;
                            list.ContentTypes.Add(contType);

                            list.Update();
                        }
                        else
                        {
                            list.ContentTypesEnabled = true;
                            list.ContentTypes.Add(contType);

                            list.Update();
                        }
                    }
                }
            }
        }



    }
}

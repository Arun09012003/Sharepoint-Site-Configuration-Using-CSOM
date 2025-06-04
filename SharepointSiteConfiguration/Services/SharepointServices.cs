using Microsoft.Graph;
using Microsoft.SharePoint.Client;
using PnP.Framework.Provisioning.Model;
using PnP.Framework.Provisioning.ObjectHandlers;
using PnP.Framework.Provisioning.Providers.Xml;
using SharepointSiteConfiguration.Models;
using System.Reflection;

namespace SharepointSiteConfiguration.Services
{
    internal class SharepointServices
    {

        public void CreateListItem(AppSettings settings, string accessToken)
        {

            using (var context = new ClientContext(settings.SiteUrl))
            {
                context.ExecutingWebRequest += (sender, e) =>
                {
                    e.WebRequestExecutor.WebRequest.Headers["Authorization"] = "Bearer " + accessToken;
                };


                var list = context.Web.Lists.GetByTitle(settings.ListName);
                var itemCreateInfo = new ListItemCreationInformation();
                var newItem = list.AddItem(itemCreateInfo);
                newItem["DepartmentList"] = "Department1";
                newItem.Update();

                context.ExecuteQuery();
                Console.WriteLine("Item created successfully.");
            }
        }

        public void FetchListItems(AppSettings settings, string accessToken)
        {
            using (var context = new ClientContext(settings.SiteUrl))
            {
                context.ExecutingWebRequest += (sender, e) =>
                {
                    e.WebRequestExecutor.WebRequest.Headers["Authorization"] = "Bearer " + accessToken;
                };

                var list = context.Web.Lists.GetByTitle(settings.ListName);

                CamlQuery query = CamlQuery.CreateAllItemsQuery();
                ListItemCollection items = list.GetItems(query);

                context.Load(items);
                context.ExecuteQuery();

                Console.WriteLine($"Items in '{settings.ListName}':");

                foreach (var item in items)
                {
                    Console.WriteLine($"ID: {item.Id}, Department: {item["DepartmentList"]}");
                }
            }
        }

        public void UpdateListItems(AppSettings settings, string accessToken, string clientNumber, string planUrl, string documentUrl, string siteUrl)
        {
            using (var context = new ClientContext(settings.SiteUrl))
            {
                context.ExecutingWebRequest += (sender, e) =>
                {
                    e.WebRequestExecutor.WebRequest.Headers["Authorization"] = "Bearer " + accessToken;
                };

                var clientList = context.Web.Lists.GetByTitle(settings.ListName);
                var items = clientList.GetItems(CamlQuery.CreateAllItemsQuery());
                context.Load(items);
                context.ExecuteQuery();

                foreach (var item in items)
                {
                    if ((string)item["ClientNumber"] == clientNumber)
                    {
                        item["ClientSiteURL"] = siteUrl;
                        item["ClientPlannerURL"] = planUrl;
                        item["ClientDocumentURL"] = documentUrl;

                        item.Update();
                        context.ExecuteQuery();

                        Console.WriteLine("Client List updated: " + clientNumber);
                        break;
                    }
                }

                
            }
        }

        public void UpdateBranchListItems(AppSettings settings, string accessToken, string branchId, string branchPlanner, string branchDocLibrary, string siteUrl)
        {
            using (var context = new ClientContext(settings.SiteUrl))
            {
                context.ExecutingWebRequest += (sender, e) =>
                {
                    e.WebRequestExecutor.WebRequest.Headers["Authorization"] = "Bearer " + accessToken;
                };

                var branchList = context.Web.Lists.GetByTitle("Branch List");
                var branch = branchList.GetItems(CamlQuery.CreateAllItemsQuery());
                context.Load(branch);
                context.ExecuteQuery();

                foreach (var item in branch)
                {
                    if ((string)item["BranchId"] == branchId)
                    {
                        item["Siteurl"] = siteUrl;
                        item["LibraryUrl"] = branchDocLibrary;
                        item["Plannerurl"] = branchPlanner;

                        item.Update();
                        context.ExecuteQuery();

                        Console.WriteLine("Branch List updated: " + branchId);
                        break;
                    }
                }
            }
        }

        public List<(string ClientName, string ClientNumber)> GetClients(AppSettings settings, string accessToken)
        {
            var clients = new List<(string, string)>();

            using (var context = new ClientContext(settings.SiteUrl))
            {
                context.ExecutingWebRequest += (sender, e) =>
                {
                    e.WebRequestExecutor.WebRequest.Headers["Authorization"] = "Bearer " + accessToken;
                };

                var list = context.Web.Lists.GetByTitle("Client List");

                CamlQuery query = CamlQuery.CreateAllItemsQuery();
                var items = list.GetItems(query);

                context.Load(items);
                context.ExecuteQuery();

                foreach (var item in items)
                {
                    var clientName = item["ClientName"]?.ToString() ?? "";
                    var clientNumber = item["ClientNumber"]?.ToString() ?? "";

                    clients.Add((clientName, clientNumber));
                }
            }

            return clients;
        }

        public List<(string BranchName, string branchId)> GetBranch(AppSettings settings, string accessToken, string clientNumber)
        {
            var branches = new List<(string, string)>();

            using (var context = new ClientContext(settings.SiteUrl))
            {
                context.ExecutingWebRequest += (sender, e) =>
                {
                    e.WebRequestExecutor.WebRequest.Headers["Authorization"] = "Bearer " + accessToken;
                };

                var list = context.Web.Lists.GetByTitle("Branch List");

                CamlQuery query = CamlQuery.CreateAllItemsQuery();
                var items = list.GetItems(query);

                context.Load(items);
                context.ExecuteQuery();

                foreach (var item in items)
                {
                    var clientField = item["Client"] as Microsoft.SharePoint.Client.FieldLookupValue;
                    if (clientField != null && clientField.LookupValue.ToString() == clientNumber)
                    {
                        var BranchName = item["BranchName"]?.ToString() ?? "";
                        var branchId = item["BranchId"]?.ToString() ?? "";
                        branches.Add((BranchName, branchId));
                    }
                }
            }
            return branches;
        }

        public async Task CreateTeamSite(AppSettings settings, string accessToken, GraphServiceClient graphClient)
        {
            var clients = GetClients(settings, accessToken);


            foreach (var client in clients)
            {

                //var groupBody = new Microsoft.Graph.Models.Group
                //{
                //    DisplayName = client.ClientName,
                //    Description = "Site "+ client.ClientName,
                //    GroupTypes = new List<string> { "Unified" },
                //    MailEnabled = true,
                //    MailNickname = client.ClientNumber,
                //    SecurityEnabled = false,
                //    Visibility = "Private"
                //};

                //var createdGroup = await graphClient.Groups.PostAsync(groupBody);

                var groupBody = new Microsoft.Graph.Group
                {
                    DisplayName = client.ClientName,
                    Description = "Site " + client.ClientName,
                    GroupTypes = new List<string> { "Unified" },
                    MailEnabled = true,
                    MailNickname = client.ClientNumber,
                    SecurityEnabled = false,
                    Visibility = "Private"
                };

                var createdGroup = await graphClient.Groups.Request().AddAsync(groupBody);

                //// Add Member to Microsoft 365 Group
                //await graphClient.Groups[createdGroup.Id].Members.Ref.PostAsync(new Microsoft.Graph.Models.ReferenceCreate
                //{
                //    OdataId = $"https://graph.microsoft.com/v1.0/users/{"arunjatak@dgneaseteq.onmicrosoft.com"}"
                //});

                //// Add Owner to Microsoft 365 Group
                //await graphClient.Groups[createdGroup.Id].Owners.Ref.PostAsync(new Microsoft.Graph.Models.ReferenceCreate
                //{
                //    OdataId = $"https://graph.microsoft.com/v1.0/users/{"arunjatak@dgneaseteq.onmicrosoft.com"}"
                //});


                var memberEmails = new List<string>
                {
                    "arunjatak@dgneaseteq.onmicrosoft.com"
                };

                //Add owner
                //foreach (var email in memberEmails) {
                //    var ownerBody = new ReferenceCreate
                //    {
                //        OdataId = $"https://graph.microsoft.com/v1.0/users/{email}",
                //    };

                //    await graphClient.Groups[createdGroup.Id].Owners.Ref.PostAsync(ownerBody);        
                //}
                foreach (var email in memberEmails)
                {
                    // Get the user object by email
                    var user = await graphClient.Users[email].Request().GetAsync();

                    // Add user as an owner
                    await graphClient.Groups[createdGroup.Id].Owners.References
                        .Request()
                        .AddAsync(user);
                }


                //Add members
                //foreach (var email in memberEmails)
                //{
                //    var memberBody = new ReferenceCreate
                //    {
                //        OdataId = $"https://graph.microsoft.com/v1.0/users/{email}",
                //    };

                //    await graphClient.Groups[createdGroup.Id].Members.Ref.PostAsync(memberBody);    
                //}
                foreach (var email in memberEmails)
                {
                    // Get the user object by email
                    var user = await graphClient.Users[email].Request().GetAsync();

                    // Add user as a member
                    await graphClient.Groups[createdGroup.Id].Members.References
                        .Request()
                        .AddAsync(user);
                }

                Console.WriteLine($"Created group: {createdGroup.Id}");

                Console.WriteLine("Waiting for SharePoint site to provision...");
                await Task.Delay(TimeSpan.FromSeconds(15));

                //var site = await graphClient.Groups[createdGroup.Id].Sites["root"].GetAsync();
                var site = await graphClient.Groups[createdGroup.Id].Sites["root"].Request().GetAsync();
                var siteUrl = site.WebUrl;
                Console.WriteLine($"Site URL: {siteUrl}");


                using (var context = new ClientContext(siteUrl))
                {
                    context.ExecutingWebRequest += (sender, e) =>
                    {
                        e.WebRequestExecutor.WebRequest.Headers["Authorization"] = "Bearer " + accessToken;
                    };

                    var lists = context.Web.Lists;
                    context.Load(lists,
                        l => l.Include(
                            x => x.Title,
                            x => x.BaseTemplate,
                            x => x.Hidden,
                            x => x.RootFolder.ServerRelativeUrl
                        )
                    );
                    context.ExecuteQuery();

                    foreach (var list in lists)
                    {
                        string documentUrl;
                        string planUrl;
                        if (list.BaseTemplate == (int)Microsoft.SharePoint.Client.ListTemplateType.DocumentLibrary && list.Title == "Documents")
                        {
                            // Get Documents library
                            var docLibrary = context.Web.Lists.GetByTitle(list.Title);
                            context.Load(docLibrary, l => l.ContentTypesEnabled);
                            context.ExecuteQuery();

                            // Enable content types if not already enabled
                            if (!docLibrary.ContentTypesEnabled)
                            {
                                docLibrary.ContentTypesEnabled = true;
                                docLibrary.Update();
                                context.ExecuteQuery();
                                Console.WriteLine("Content types enabled on " + list.Title + " library.");
                                await Task.Delay(TimeSpan.FromSeconds(5));
                                context.Load(list, l => l.ContentTypes);
                                context.ExecuteQuery();

                                // Find and remove default "Document" content type
                                var defaultContentType = list.ContentTypes.FirstOrDefault(ct => ct.Name == "Document");
                                if (defaultContentType != null)
                                {
                                    defaultContentType.DeleteObject();
                                    context.ExecuteQuery();
                                    Console.WriteLine("Default 'Document' content type removed from the library.");

                                    Web web = context.Web;
                                    context.Load(web, w => w.ContentTypes);
                                    context.ExecuteQuery();

                                }
                                else
                                {
                                    Console.WriteLine("'Document' content type not found in the library.");
                                }

                                string contentTypeId = "0x01010097E7D4CF65DFB04EA47382EADE69C850".Trim();
                                int retry = 0;
                                while (retry < 10)
                                {
                                    var contentTypeSubscriber = new Microsoft.SharePoint.Client.Taxonomy.ContentTypeSync.ContentTypeSubscriber(context);
                                    context.Load(contentTypeSubscriber);
                                    context.ExecuteQuery();

                                    var contentTypeImportResponse = contentTypeSubscriber.SyncContentTypesFromHubSite2(siteUrl, new List<string> { contentTypeId });
                                    await Task.Delay(TimeSpan.FromSeconds(5));
                                    var siteContentTypes = context.Web.ContentTypes;
                                    context.Load(siteContentTypes);
                                    context.ExecuteQuery();

                                    var contentType = siteContentTypes.FirstOrDefault(ct => ct.Id.StringValue.StartsWith(contentTypeId));

                                    if (contentType != null)
                                    {
                                        list.ContentTypes.AddExistingContentType(contentType);
                                        list.Update();
                                        context.ExecuteQuery();
                                        Console.WriteLine($"Added content type '{contentType.Name}' to the library.");
                                        break;
                                    }
                                    else
                                    {
                                        Console.WriteLine("Content type not found in this site. Waiting for sync...");
                                        retry++;
                                        await Task.Delay(TimeSpan.FromSeconds(5));
                                    }
                                }



                            }
                            else
                            {
                                Console.WriteLine("Content types are already enabled.");
                            }

                            // Enable folder creation if it is not
                            docLibrary.EnableFolderCreation = true;
                            docLibrary.Update();
                            context.ExecuteQuery();

                            // Create Folder
                            ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                            string[] folderName = ["Folder1", "Folder2"];
                            foreach (string folder in folderName)
                            {
                                itemCreateInfo.UnderlyingObjectType = FileSystemObjectType.Folder;
                                itemCreateInfo.LeafName = folder;

                                Microsoft.SharePoint.Client.ListItem newItem = list.AddItem(itemCreateInfo);
                                newItem["Title"] = folderName;
                                newItem.Update();
                                context.ExecuteQuery();
                            }
                            documentUrl = siteUrl + "/Shared%20Documents";

                            //Setting default column value in document library 
                            Web web1 = context.Web;
                            Microsoft.SharePoint.Client.List list1 = web1.Lists.GetByTitle(list.Title);
                            Microsoft.SharePoint.Client.Field field = list1.Fields.GetByInternalNameOrTitle("ClientNumber1");
                            field.DefaultValue = client.ClientNumber;
                            field.Update();
                            context.Load(field);
                            context.ExecuteQuery();

                            // Creating planner plan
                            var planBody = new PlannerPlan
                            {
                                Owner = createdGroup.Id,
                                Title = client.ClientName
                            };
                            var result = await graphClient.Planner.Plans.Request().AddAsync(planBody);
                            planUrl = "https://planner.cloud.microsoft/webui/plan/" + result.Id + "/view";
                            Console.WriteLine("Planner plan cretaed successfully, Plan Name: " + result.Title + " Plan Id: " + result.Id);

                            UpdateListItems(settings, accessToken, client.ClientNumber, planUrl, documentUrl, siteUrl);

                        }
                        else
                        {
                            continue;
                        }
                    }
                }
            }
        }

        public async Task CreateGroupSite(AppSettings settings, string accessToken, GraphServiceClient graphClient)
        {
            var clients = GetClients(settings, accessToken);
            

            foreach (var client in clients)
            {
                string clientDocumentURL="";
                string clientPlannerURL="";
                var groupBody = new Microsoft.Graph.Group
                {
                    DisplayName = client.ClientName,
                    Description = "Site " + client.ClientName,
                    GroupTypes = new List<string> { "Unified" },
                    MailEnabled = true,
                    MailNickname = client.ClientNumber,
                    SecurityEnabled = false,
                    Visibility = "Private"
                };

                var createdGroup = await graphClient.Groups.Request().AddAsync(groupBody);

                var memberEmails = new List<string> { "arunjatak@dgneaseteq.onmicrosoft.com" };

                foreach (var email in memberEmails)
                {
                    // Get the user object by email
                    var user = await graphClient.Users[email].Request().GetAsync();

                    // Add user as an owner
                    await graphClient.Groups[createdGroup.Id].Owners.References
                        .Request()
                        .AddAsync(user);
                }
                foreach (var email in memberEmails)
                {
                    // Get the user object by email
                    var user = await graphClient.Users[email].Request().GetAsync();

                    // Add user as an member
                    await graphClient.Groups[createdGroup.Id].Members.References
                        .Request()
                        .AddAsync(user);
                }

                Console.WriteLine($"Created group: {createdGroup.Id}");
                Console.WriteLine("Waiting for SharePoint site to provision...");
                await Task.Delay(TimeSpan.FromSeconds(15));

                var site = await graphClient.Groups[createdGroup.Id].Sites["root"].Request().GetAsync();
                var siteUrl = site.WebUrl;
                Console.WriteLine($"Site URL: {siteUrl}");

                using (var context = new ClientContext(siteUrl))
                {
                    context.ExecutingWebRequest += (sender, e) =>
                    {
                        e.WebRequestExecutor.WebRequest.Headers["Authorization"] = "Bearer " + accessToken;
                    };

                    var lists = context.Web.Lists;
                    context.Load(lists,
                        l => l.Include(
                            x => x.Title,
                            x => x.BaseTemplate,
                            x => x.Hidden,
                            x => x.RootFolder.ServerRelativeUrl
                        )
                    );
                    context.ExecuteQuery();

                    var docLibrary = lists.FirstOrDefault(l => l.BaseTemplate == 101 && l.Title == "Documents");

                    if (docLibrary != null)
                    {
                        var branches = GetBranch(settings, accessToken, client.ClientNumber);


                        foreach (var branch in branches)
                        {
                            string branchDocLib;
                            string branchPlannerPlan;
                            // Create new document library
                            ListCreationInformation lci = new ListCreationInformation
                            {
                                Title = branch.BranchName,
                                Description = "DocLib",
                                TemplateType = 101
                            };
                            var newLib = context.Web.Lists.Add(lci);
                            context.Load(newLib);
                            context.ExecuteQuery();
                            branchDocLib = siteUrl+"/"+branch.BranchName;
                            await Task.Delay(TimeSpan.FromSeconds(5));


                            var branchPlan = new PlannerPlan
                            {
                                Owner = createdGroup.Id,
                                Title = branch.BranchName
                            };
                            var branchResult = await graphClient.Planner.Plans.Request().AddAsync(branchPlan);
                            branchPlannerPlan = "https://planner.cloud.microsoft/webui/plan/" + branchResult.Id + "/view";
                            Console.WriteLine("Planner plan created successfully for branch, Plan Name: " + branchResult.Title + ", Plan Id: " + branchResult.Id);


                            context.Load(newLib, l => l.ContentTypesEnabled, l => l.ContentTypes);
                            context.ExecuteQuery();

                            // Enable content types
                            if (!newLib.ContentTypesEnabled)
                            {
                                newLib.ContentTypesEnabled = true;
                                newLib.Update();
                                context.ExecuteQuery();
                                Console.WriteLine("Content types enabled on " + branch.BranchName + " library.");
                                await Task.Delay(TimeSpan.FromSeconds(3));

                                var defaultCT = newLib.ContentTypes.FirstOrDefault(ct => ct.Name == "Document");
                                if (defaultCT != null)
                                {
                                    defaultCT.DeleteObject();
                                    context.ExecuteQuery();
                                    Console.WriteLine("Default 'Document' content type removed.");
                                }

                                string contentTypeId = "0x01010097E7D4CF65DFB04EA47382EADE69C850".Trim();
                                int retry = 0;
                                while (retry < 10)
                                {
                                    var subscriber = new Microsoft.SharePoint.Client.Taxonomy.ContentTypeSync.ContentTypeSubscriber(context);
                                    context.Load(subscriber);
                                    context.ExecuteQuery();

                                    subscriber.SyncContentTypesFromHubSite2(siteUrl, new List<string> { contentTypeId });
                                    await Task.Delay(5000);

                                    context.Load(context.Web.ContentTypes);
                                    context.ExecuteQuery();
                                    var importedCT = context.Web.ContentTypes.FirstOrDefault(ct => ct.Id.StringValue.StartsWith(contentTypeId));
                                    if (importedCT != null)
                                    {
                                        newLib.ContentTypes.AddExistingContentType(importedCT);
                                        newLib.Update();
                                        context.ExecuteQuery();
                                        Console.WriteLine($"Added content type '{importedCT.Name}' to the library.");
                                        break;
                                    }
                                    else
                                    {
                                        Console.WriteLine("Waiting for content type to sync...");
                                        retry++;
                                        await Task.Delay(5000);
                                    }
                                }
                            }

                            // Set default field value
                            Microsoft.SharePoint.Client.Field clientField = newLib.Fields.GetByInternalNameOrTitle("ClientNumber1");
                            clientField.DefaultValue = client.ClientNumber;
                            clientField.Update();
                            context.Load(clientField);
                            context.ExecuteQuery();

                            // Create folders
                            string[] folderNames = ["Folder1", "Folder2"];
                            foreach (var folder in folderNames)
                            {
                                var itemCreateInfo = new ListItemCreationInformation
                                {
                                    UnderlyingObjectType = Microsoft.SharePoint.Client.FileSystemObjectType.Folder,
                                    LeafName = folder
                                };
                                Microsoft.SharePoint.Client.ListItem folderItem = newLib.AddItem(itemCreateInfo);
                                folderItem["Title"] = folder;
                                folderItem.Update();
                                context.ExecuteQuery();
                            }
                            Console.WriteLine("Folder created in document library...");



                            //Create site page
                            Microsoft.SharePoint.Client.List Library = context.Web.Lists.GetByTitle("site pages");
                            context.Load(context.Web, w => w.ServerRelativeUrl);
                            context.ExecuteQuery();
                            string serverRelativeUrl = $"{context.Web.ServerRelativeUrl.TrimEnd('/')}/SitePages/" + branch.BranchName + ".aspx";
                            Microsoft.SharePoint.Client.ListItem oItem = Library.RootFolder.Files
                                .AddTemplateFile(serverRelativeUrl, TemplateFileType.ClientSidePage)
                                .ListItemAllFields;

                            oItem["ContentTypeId"] = "0x0101009D1CB255DA76424F860D91F20E6C4118";
                            oItem["Title"] = System.IO.Path.GetFileNameWithoutExtension(branch.BranchName + ".aspx");
                            oItem["ClientSideApplicationId"] = "b6917cb1-93a0-4b97-a84d-7cf49975d4ec";
                            oItem["PageLayoutType"] = "Article";
                            oItem["PromotedState"] = "0";
                            oItem["CanvasContent1"] = "<div></div>";
                            oItem["BannerImageUrl"] = "/_layouts/15/images/sitepagethumbnail.png";

                            oItem.Update();
                            context.Load(oItem, item => item.Id);
                            context.ExecuteQuery();
                            Console.WriteLine("Successfully created modern page in library");
                            Console.WriteLine("Page ID: " + oItem.Id);

                            await Task.Delay(TimeSpan.FromSeconds(5));

                            //Add webpart to page
                            using (var fileStream = System.IO.File.OpenRead(@"C:\Projects\PnP Powershell\Project\CSOM\CreateSiteUsingPnPFramework\CreateSiteUsingPnPFramework\File\TeamSiteStructure.xml"))
                            {
                                using (var memoryStream = new System.IO.MemoryStream())
                                {
                                    fileStream.CopyTo(memoryStream);
                                    memoryStream.Position = 0;

                                    try
                                    {
                                        var provider = new XMLStreamTemplateProvider();
                                        ProvisioningTemplate template = provider.GetTemplate(memoryStream);

                                        var applyingInformation = new ProvisioningTemplateApplyingInformation();
                                        template.Parameters["BranchId"] = branch.branchId;
                                        template.Parameters["BranchName"] = branch.BranchName;
                                        context.Web.ApplyProvisioningTemplate(template, applyingInformation);
                                        Console.WriteLine("Webpart added successfully.");

                                    }
                                    catch (ReflectionTypeLoadException ex)
                                    {
                                        foreach (var loaderEx in ex.LoaderExceptions)
                                        {
                                            Console.WriteLine(loaderEx.Message);
                                        }
                                    }
                                }
                            }
                            UpdateBranchListItems(settings, accessToken, branch.branchId, branchPlannerPlan, branchDocLib, siteUrl + "/SitePages/Home.aspx");

                        }
                    }
                }

                // Creating planner plan
                var planBody = new PlannerPlan
                {
                    Owner = createdGroup.Id,
                    Title = client.ClientName
                };
                var result = await graphClient.Planner.Plans.Request().AddAsync(planBody);
                clientPlannerURL = "https://planner.cloud.microsoft/webui/plan/" + result.Id + "/view";
                Console.WriteLine("Planner plan created successfully for Client, Plan Name: " + result.Title + ", Plan Id: " + result.Id);
                clientDocumentURL= siteUrl + "/Shared%20Documents";
                UpdateListItems(settings, accessToken, client.ClientNumber, clientPlannerURL, clientDocumentURL, siteUrl+ "/SitePages/Home.aspx");
            }
        }
    }
}

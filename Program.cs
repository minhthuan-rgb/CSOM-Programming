using Azure.Identity;
using Microsoft.Extensions.Configuration;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Search.Query;
using Microsoft.SharePoint.Client.Taxonomy;
using Microsoft.SharePoint.Client.UserProfiles;
using Microsoft.SharePoint.Client.Utilities;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CSOM_Programming
{
    class Program
    {
        static async Task Main(string[] args)
        {
            try
            {
                using (var clientContextHelper = new ClientContextHelper())
                {
                    ClientContext ctx = GetContext(clientContextHelper);

                    // Get Site Title
                    ctx.Load(ctx.Web);
                    await ctx.ExecuteQueryAsync();

                    Console.OutputEncoding = Encoding.UTF8;
                    Console.WriteLine($"Connected Site: {ctx.Web.Title}");


                    #region CSOM Programming
                    //await CreateList(ctx, "CSOM Test", ListTemplateType.Announcements); // Create List "CSOM Test"

                    //await CreateTermSetAndTerms(ctx); // Create Term Set And 2 Terms

                    //await CreateTextField(ctx, "hoabout"); // Create site field "about" type text
                    //await CreateTaxonomyField(ctx, "hocity"); // Create site field "city" type taxonomy

                    //await CreateContentType(ctx, "CSOM Test content type"); // Create Content Type "CSOM Test content type"

                    //await AddContentTypeToList(ctx, "CSOM Test"); // Add Content Type To List "CSOM Test"

                    //await AddFieldsToContentType(ctx); // Add 2 Fields "about" And "city" To Content Type "CSOM Test content type"

                    //await SetDefaultContentType(ctx, "CSOM Test"); // Set "CSOM Test content type" As Default Content Type In List "CSOM test"

                    //await BindTaxonomyFieldToTermSet(ctx, "hocity"); // Bind Taxonomy Field "city" To Term Set

                    //await DisplayAllItemsListView(ctx); // Display All Items List View

                    //await AddItemsToList(ctx, 5); // Add 5 Items To List "CSOM Test"

                    //await SetDefaultValueForAboutField(ctx); // Set Default Value For "about" site field

                    //await AddItemsToList(ctx, 2, true); // Add 2 Items With Default Value Of "about" site field

                    //await SetDefaultValueForCityField(ctx); // Set Default Value For "city" site field

                    //await AddItemsToList(ctx, 2, true, true); // Add 2 Items With Default Value Of "city" site field

                    //await GetListItems(ctx); // Get List Items Where Field "about" is not "about default"

                    //await CreateListView(ctx); // Create List View by CSOM 

                    //await UpdateListItems(ctx); // Update List View Items

                    //await CreatePeopleField(ctx); // Create People Field In List "CSOM Test"

                    //await SetUserAdminToAuthorField(ctx); // Migrate all list items to set user admin to field "author"

                    //await CreateTaxonomyFieldMulti(ctx, "hocities"); // Create site field "cities" type taxonomy multi values

                    //await BindTaxonomyFieldToTermSet(ctx, "hocities"); // Bind Taxonomy Field "cities" To Term Set

                    //await AddCitiesFieldToContentType(ctx); // Add Field "cities" To Content Type "CSOM Test Content Type"

                    //await AddItemsToList(ctx, 3, true, true, true); // Add 3 Items With Field "Cities" Multi Value

                    //await DisplayCitiesField(ctx); // Display "cities" Field 

                    //await CreateList(ctx, "Document Test", ListTemplateType.DocumentLibrary); // Create List "Document Test"

                    //await AddContentTypeToList(ctx, "Document Test"); // Add Content Type TO List "Document Test"

                    //await CreateFolders(ctx); // Create Folders For List "Document Test"

                    //await AddFilesInsideFolder(ctx, 3, "FolderTest"); // Add 3 Files In "Folder 2" With Value "Folder test" In Field "about"

                    //await AddFilesInsideFolder(ctx, 2, "CitiesTest", true); // Add 2 Files In "Folder 2" With Value "Stockholm" In Field "cities"

                    //await GetListItemInFolderOnly(ctx); // Get All List Items Just In "Folder 2" And Have Value "Stockholm" in "cities" field

                    //await CreateListItemByUploadFile(ctx); // Create List Item In "Document Test" By Upload A File Document.docx

                    //await DisplayAllDocumentsListView(ctx); // Display All Documents List View

                    //await CreateFolderStructureView(ctx); // Create Folder Structure View

                    ///await LoadUser(ctx, "59Tese"); // Load User From User Email Or Name

                    //await GetTaxonomyHiddenListItems(ctx); // Load TaxonomyHiddenList Items

                    //await LoadSiteUsers(ctx); // Load All Site Users

                    //await RemoveDeletedUser(ctx); // Remove Site Users That Was Deleted In Server
                    #endregion


                    #region Permissions
                    //await CreateSubSite(ctx); // Create Subsite "Finance And Accounting"

                    //await CreateListAtSubsite(ctx, "Accounts", ListTemplateType.Announcements); // Create List "Accounts" at subsite

                    //await StopInheritingPermissions(ctx); // Stop Inheriting Permissions In List "Account"

                    //await GrantDesignPermissionForUser(ctx, "thien.pham.minh", RoleType.WebDesigner); // Grant "Design" Permission For A User

                    //await DeleteUniquePermissions(ctx); // Re-establish Inheritance In List "Account"

                    //await CreateCustomPermissionLevel(ctx); // Create Custom Permission Level

                    //await CreateCustomSecurityGroup(ctx); // Create Custom Secure Group

                    // Add 3 Users To Custom Security Group
                    //await AddUserToSecurityGroup(ctx, "thanh.pham.minh", "Test Group");
                    //await AddUserToSecurityGroup(ctx, "thao.pham.nguyen.phuong", "Test Group");
                    //await AddUserToSecurityGroup(ctx, "thien.pham.minh", "Test Group");

                    //await CheckInheritedPermissionLevel(ctx); // Check That Permission Level Of Group Has Been Inherited From The Root Site
                    #endregion


                    #region User Profile
                    //await DisplaySomePropertiesForAllUsers(ctx); // Display Some Properties For All Users In The Tenant

                    //await UpdateUserProperty(ctx, "59Tese@HenoldMK.onmicrosoft.com"); // Update User Property
                    #endregion


                    #region KQL
                    await KQLFilter(ctx);
                    #endregion
                }

                Console.WriteLine($"Press Any Key To Stop!");
                Console.ReadKey();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }


        #region Client Context
        static ClientContext GetContext(ClientContextHelper clientContextHelper)
        {
            var builder = new ConfigurationBuilder().AddJsonFile($"appsettings.json", true, true);
            IConfiguration config = builder.Build();
            var info = config.GetSection("SharepointInfo").Get<SharepointInfo>();
            return clientContextHelper.GetContext(new Uri(info.SiteUrl), info.Username, info.Password);
        }
        #endregion


        #region List
        private static async Task CreateList(ClientContext ctx, string title, ListTemplateType type, Web web = null)
        {
            var creationInfo = new ListCreationInformation
            {
                Title = title,
                TemplateType = (int)type
            };
            List list;
            if (web == null)
                list = ctx.Web.Lists.Add(creationInfo);
            else list = web.Lists.Add(creationInfo);

            list.Description = $"This Is {title} List, That Was Created From Client Side";
            list.OnQuickLaunch = true;
            list.Update();

            await ctx.ExecuteQueryAsync();

            Console.WriteLine($"Successfully Created '{list.Title}' List!");
        }

        private static async Task AddContentTypeToList(ClientContext ctx, string title)
        {
            List myList = ctx.Web.Lists.GetByTitle(title);
            ctx.Load(myList, list => list.ContentTypesEnabled);
            await ctx.ExecuteQueryAsync();

            if (!myList.ContentTypesEnabled)
            {
                myList.ContentTypesEnabled = true;
                myList.Update();
                await ctx.ExecuteQueryAsync();
            }
            ContentTypeCollection contentTypes = ctx.Web.ContentTypes;
            ctx.Load(contentTypes, cts => cts.Include(ct => ct.Name));
            await ctx.ExecuteQueryAsync();

            ContentType contentType = contentTypes.First(c => c.Name.Equals("CSOM Test content type"));
            myList.ContentTypes.AddExistingContentType(contentType);
            myList.Update();

            await ctx.ExecuteQueryAsync();

            Console.WriteLine("Finished!");
        }

        private static async Task SetDefaultContentType(ClientContext ctx, string title)
        {
            List myList = ctx.Web.Lists.GetByTitle(title);
            ContentTypeCollection contentTypes = myList.ContentTypes;
            ctx.Load(contentTypes, cts => cts.Include(ct => ct.Name,
                                                      ct => ct.Id));
            await ctx.ExecuteQueryAsync();

            IList<ContentTypeId> reverse = new List<ContentTypeId>();
            foreach (ContentType ct in contentTypes)
                if (ct.Name.Equals("CSOM Test content type"))
                    reverse.Add(ct.Id);

            myList.RootFolder.UniqueContentTypeOrder = reverse;
            myList.RootFolder.Update();
            myList.EnableAttachments = false;
            myList.Update();
            await ctx.ExecuteQueryAsync();

            Console.WriteLine("Finished!");
        }

        private static async Task AddItemsToList(ClientContext ctx, int amount, bool isAboutDefault = false, bool isCityDefault = false, bool isMulti = false)
        {
            for (int i = 0; i < amount; i++)
            {
                await AddItemToList(ctx, isAboutDefault ? null : (i + 1).ToString(), isCityDefault, isMulti);
            }
        }

        private static async Task AddItemToList(ClientContext ctx, string about = null, bool isCityDefault = false, bool isMulti = false)
        {
            List myList = ctx.Web.Lists.GetByTitle("CSOM Test");

            var creationInfo = new ListItemCreationInformation();
            ListItem newItem = myList.AddItem(creationInfo);
            if (about != null)
                newItem["hoabout"] = $"Item {about}";

            if (!isCityDefault)
            {
                var cityField = ctx.Web.Fields.GetByTitle("hocity");

                var taxCityField = ctx.CastTo<TaxonomyField>(cityField);
                taxCityField.SetFieldValueByTerm(newItem, GetTermByName(ctx, "Stockholm"), 1033);
            }

            if (isMulti)
            {
                var citiesField = ctx.Web.Fields.GetByTitle("hocities");

                var taxCitiesField = ctx.CastTo<TaxonomyField>(citiesField);
                taxCitiesField.SetFieldValueByTermCollection(newItem, GetAllTerms(ctx), 1033);
            }

            newItem.Update();
            await ctx.ExecuteQueryAsync();

            Console.WriteLine($"Created Item!");
        }

        private static async Task CreateListView(ClientContext ctx)
        {
            List myList = ctx.Web.Lists.GetByTitle("CSOM Test");
            ctx.Load(myList, ml => ml.Title);

            var creationInfo = new ViewCreationInformation
            {
                Title = "CSOM Test View",
                //creationInfo.SetAsDefaultView = true;
                ViewTypeKind = ViewType.Html,
                Query = @$"<OrderBy>
                                    <FieldRef Name='Created' Ascending='False'/>
                                </OrderBy>
                                <Where>
                                    <Eq>
                                        <FieldRef Name='hocity'/>
                                        <Value Type='{FieldType.TaxonomyFieldType.ToString()}'>Ho Chi Minh</Value>
                                    </Eq>
                                </Where>"
            };
            string commaSeparateColumnNames = "ID, Title, hoabout, hocity, Created";
            creationInfo.ViewFields = commaSeparateColumnNames.Split(", ");

            View listView = myList.Views.Add(creationInfo);
            ctx.Load(listView, l => l.Title);

            await ctx.ExecuteQueryAsync();

            Console.WriteLine($"Successfully Created List View '{listView.Title}' For List '{myList.Title}'!");
        }

        private static async Task DisplayAllItemsListView(ClientContext ctx)
        {
            List myList = ctx.Web.Lists.GetByTitle("CSOM Test");
            ctx.Load(myList, ml => ml.Title);

            View allItemsView = await GetListView(ctx, myList, "All items");

            allItemsView.ViewFields.RemoveAll();
            allItemsView.ViewFields.Add("ID");
            allItemsView.ViewFields.Add("hoabout");
            allItemsView.ViewFields.Add("hocity");

            allItemsView.DefaultView = true;
            allItemsView.Update();

            await ctx.ExecuteQueryAsync();

            Console.WriteLine($"Successfully Updated List View '{allItemsView.Title}' For List '{myList.Title}'!");
        }

        private static async Task DisplayCitiesField(ClientContext ctx)
        {
            List myList = ctx.Web.Lists.GetByTitle("CSOM Test");
            ctx.Load(myList, ml => ml.Title);

            View allItemsView = await GetListView(ctx, myList, "All items");

            allItemsView.ViewFields.Add("hocities");
            allItemsView.Update();

            await ctx.ExecuteQueryAsync();

            Console.WriteLine($"Successfully Added 'cities' Field To List View '{allItemsView.Title}' For List '{myList.Title}'!");
        }

        private static async Task<View> GetListView(ClientContext ctx, List list, string title)
        {
            ViewCollection views = list.Views;
            ctx.Load(views, vs => vs.Include(v => v.Title));
            await ctx.ExecuteQueryAsync();

            View allItemsView = views.First(view => view.Title.Equals(title));

            return allItemsView;
        }

        private static async Task CreatePeopleField(ClientContext ctx)
        {
            List myList = ctx.Web.Lists.GetByTitle("CSOM Test");

            await CreateField(ctx, "hoauthor", FieldType.User, myList.Fields);
        }

        private static async Task SetUserAdminToAuthorField(ClientContext ctx)
        {
            List myList = ctx.Web.Lists.GetByTitle("CSOM Test");

            var items = myList.GetItems(new CamlQuery()
            {
                ViewXml = @"<View>
                                <Query>
                                    <OrderBy>
                                        <FieldRef Name='ID'/>
                                    </OrderBy>
                                </Query>
                                <ViewFields>
                                    <FieldRef Name='hoauthor'/>
                                </ViewFields>
                            </View>"
            });
            ctx.Load(items, its => its.Include(it => it["hoauthor"]));

            User user = ctx.Web.EnsureUser(await LoadCurrentUserEmail(ctx));
            ctx.Load(user, u => u.Id);

            await ctx.ExecuteQueryAsync();

            FieldUserValue userValue = new()
            {
                LookupId = user.Id,
            };

            foreach (var item in items)
            {
                item["hoauthor"] = userValue;
                item.Update();
            }

            await ctx.ExecuteQueryAsync();

            Console.WriteLine("Finished!");
        }

        private static async Task<string> LoadCurrentUserEmail(ClientContext ctx)
        {
            //User currentUser = ctx.Web.CurrentUser;
            PeopleManager peopleManager = new(ctx);
            PersonProperties properties = peopleManager.GetMyProperties();
            ctx.Load(properties, p => p.DisplayName,
                                 p => p.Email);
            await ctx.ExecuteQueryAsync();
            return properties.Email;
        }

        private static async Task DeleteList(ClientContext ctx, string title)
        {
            List list = ctx.Web.Lists.GetByTitle(title);
            ctx.Load(list, l => l.Title);
            list.DeleteObject();

            await ctx.ExecuteQueryAsync();

            Console.WriteLine($"Successfully Deleted {list.Title} List!");
        }

        private static async Task CreateFolders(ClientContext ctx)
        {
            List myList = ctx.Web.Lists.GetByTitle("Document Test");

            var folder = myList.RootFolder;
            folder = folder.Folders.Add("Folder 1");
            folder.Folders.Add("Folder 2");

            await ctx.ExecuteQueryAsync();

            Console.WriteLine("Finished!");
        }

        private static async Task AddFilesInsideFolder(ClientContext ctx, int amount, string title, bool isCities = false)
        {
            Folder folder = ctx.Web.GetFolderByServerRelativeUrl(ctx.Web.ServerRelativeUrl + "/Document%20Test/Folder%201/Folder%202");

            for (int i = 0; i < amount; i++)
                await AddFileToFolder(ctx, $"{title}{i + 1}", folder, isCities);
        }

        private static async Task AddFileToFolder(ClientContext ctx, string title, Folder folder, bool isCities)
        {
            FileCollection files = folder.Files;
            ctx.Load(files, fs => fs.Include(f => f.Name));

            ContentTypeCollection contentTypes = ctx.Web.ContentTypes;
            ctx.Load(contentTypes, cts => cts.Include(ct => ct.Name,
                                                      ct => ct.Id));
            await ctx.ExecuteQueryAsync();

            var temp = files.FirstOrDefault(f => f.Name.Equals($"{title}.docx"));

            if (temp != null)
                throw new Exception($"'{title}.docx' Has Existed!");

            var creationInfo = new FileCreationInformation()
            {
                Content = Encoding.ASCII.GetBytes("Folder test"),
                Overwrite = true,
                Url = $"{title}.docx"
            };
            var addedFile = files.Add(creationInfo);

            ContentType contentType = contentTypes.First(c => c.Name.Equals("CSOM Test content type"));
            var newItem = addedFile.ListItemAllFields;
            newItem["ContentTypeId"] = contentType.Id;
            newItem.Update();

            if (isCities)
            {
                var citiesField = ctx.Web.Fields.GetByTitle("hocities");
                var taxCitiesField = ctx.CastTo<TaxonomyField>(citiesField);

                taxCitiesField.SetFieldValueByTerm(newItem, GetTermByName(ctx, "Stockholm"), 1033);
            }
            else newItem["hoabout"] = "Folder Test";

            newItem.Update();

            await ctx.ExecuteQueryAsync();

            Console.WriteLine($"Created Item {addedFile.Name}");
        }

        private static async Task CreateListItemByUploadFile(ClientContext ctx)
        {
            List myList = ctx.Web.Lists.GetByTitle("Document Test");
            var folder = myList.RootFolder;

            ContentTypeCollection contentTypes = ctx.Web.ContentTypes;
            ctx.Load(contentTypes, cts => cts.Include(ct => ct.Name,
                                                      ct => ct.Id));
            await ctx.ExecuteQueryAsync();

            var creationInfo = new FileCreationInformation()
            {
                Content = System.IO.File.ReadAllBytes("D:\\Document.docx"),
                Overwrite = true,
                Url = Path.GetFileName("D:\\Document.docx")
            };
            var uploadFile = folder.Files.Add(creationInfo);

            ContentType contentType = contentTypes.First(c => c.Name.Equals("CSOM Test content type"));
            var newItem = uploadFile.ListItemAllFields;
            newItem["ContentTypeId"] = contentType.Id;
            newItem.Update();

            await ctx.ExecuteQueryAsync();

            Console.WriteLine($"Created Item {uploadFile.Name}");
        }

        private static async Task DisplayAllDocumentsListView(ClientContext ctx)
        {
            List myList = ctx.Web.Lists.GetByTitle("Document Test");
            ctx.Load(myList, ml => ml.Title);
            View allDocumentsView = await GetListView(ctx, myList, "All Documents");

            allDocumentsView.ViewFields.Add("hoabout");
            allDocumentsView.ViewFields.Add("hocity");
            allDocumentsView.ViewFields.Add("hocities");
            allDocumentsView.DefaultView = true;
            allDocumentsView.Update();

            await ctx.ExecuteQueryAsync();

            Console.WriteLine($"Successfully Updated List View '{allDocumentsView.Title}' For List '{myList.Title}'!");
        }

        private static async Task CreateFolderStructureView(ClientContext ctx)
        {
            List myList = ctx.Web.Lists.GetByTitle("Document Test");
            ctx.Load(myList, ml => ml.Title);

            ViewCollection views = myList.Views;
            ctx.Load(views, vs => vs.Include(v => v.Title));
            await ctx.ExecuteQueryAsync();

            View temp = views.FirstOrDefault(view => view.Title.Equals("Folders"));

            if (temp != null)
                throw new Exception($"List View '{temp.Title}' Has Existed!");

            var creationInfo = new ViewCreationInformation
            {
                Title = "Folders",
                ViewTypeKind = ViewType.Html,
                Query = @$"<Where>
                                    <Eq>
                                        <FieldRef Name='FSObjType'/>
                                        <Value Type='{FieldType.Integer.ToString()}'>1</Value>
                                    </Eq>
                                </Where>"
            };
            string commaSeparateColumnNames = "Type, Name, Modified, Modified By";
            creationInfo.ViewFields = commaSeparateColumnNames.Split(", ");

            View listView = views.Add(creationInfo);
            listView.Scope = ViewScope.DefaultValue;
            listView.DefaultView = true;
            listView.Update();

            ctx.Load(listView, l => l.Title);
            await ctx.ExecuteQueryAsync();

            Console.WriteLine($"Successfully Created List View '{listView.Title}' For List '{myList.Title}'!");
        }
        #endregion


        #region Term Set And Terms
        private static async Task CreateTermSetAndTerms(ClientContext ctx)
        {
            TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(ctx);
            TermStore termStore = taxonomySession.GetDefaultSiteCollectionTermStore();

            // Create Group Test
            //TermGroup termGroup = termStore.CreateGroup("Test", Guid.NewGuid());
            TermGroup termGroup = termStore.Groups.GetByName("Test");

            // Create Term Set
            TermSet termSet = termGroup.CreateTermSet("city-PhamMinhThuan", Guid.NewGuid(), 1033);

            // Create 2 Terms
            termSet.CreateTerm("Ho Chi Minh", 1033, Guid.NewGuid());
            termSet.CreateTerm("Stockholm", 1033, Guid.NewGuid());

            await ctx.ExecuteQueryAsync();

            Console.WriteLine("Finished!");
        }

        private static Term GetTermByName(ClientContext ctx, string termName)
        {
            TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(ctx);
            TermStore termStore = taxonomySession.GetDefaultSiteCollectionTermStore();
            TermGroup termGroup = termStore.Groups.GetByName("Test");
            TermSet termSet = termGroup.TermSets.GetByName("city-PhamMinhThuan");

            return termSet.Terms.GetByName(termName);
        }

        private static TermCollection GetAllTerms(ClientContext ctx)
        {
            TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(ctx);
            TermStore termStore = taxonomySession.GetDefaultSiteCollectionTermStore();
            TermGroup termGroup = termStore.Groups.GetByName("Test");
            TermSet termSet = termGroup.TermSets.GetByName("city-PhamMinhThuan");

            return termSet.GetAllTerms();
        }
        #endregion


        #region Fields
        private static async Task CreateTextField(ClientContext ctx, string name)
        {
            await CreateField(ctx, name, FieldType.Text);
        }

        private static async Task CreateTaxonomyField(ClientContext ctx, string name)
        {
            await CreateField(ctx, name, FieldType.TaxonomyFieldType);
        }

        private static async Task CreateTaxonomyFieldMulti(ClientContext ctx, string name)
        {
            await CreateField(ctx, name, FieldType.TaxonomyFieldTypeMulti);
        }

        private static async Task CreateField(ClientContext ctx, string name, FieldType fieldType, FieldCollection fields = null)
        {
            string isMultiStr = fieldType == FieldType.TaxonomyFieldTypeMulti ? "Mult='TRUE' " : "";
            string isUserString = fieldType == FieldType.User ? "UserSelectionMode='PeopleOnly' " : "";

            string schemaField = $"<Field ID='{Guid.NewGuid()}' " +
                                $"Type='{fieldType.ToString()}' " +
                                isMultiStr +
                                $"Name='{name}' " +
                                $"StaticName='{name}' " +
                                $"DisplayName='{name}' " +
                                isUserString +
                                $"Group='59Tese'/>";

            Field field = null;
            if (fields != null)
                field = fields.AddFieldAsXml(schemaField, true, AddFieldOptions.AddFieldInternalNameHint);
            else field = ctx.Web.Fields.AddFieldAsXml(schemaField, true, AddFieldOptions.AddFieldInternalNameHint);

            ctx.Load(field, f => f.StaticName);
            await ctx.ExecuteQueryAsync();

            Console.WriteLine($"Successfully Created '{field.StaticName}' Field!");
        }

        private static async Task BindTaxonomyFieldToTermSet(ClientContext ctx, string fieldName)
        {
            TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(ctx);
            TermStore termStore = taxonomySession.GetDefaultSiteCollectionTermStore();

            // Get Group
            TermGroup termGroup = termStore.Groups.GetByName("Test");

            // Get Term Set
            TermSet termSet = termGroup.TermSets.GetByName("city-PhamMinhThuan");

            ctx.Load(termStore, t => t.Id);
            ctx.Load(termSet, t => t.Id);
            await ctx.ExecuteQueryAsync();

            Field field = ctx.Web.Fields.GetByTitle(fieldName);
            TaxonomyField taxCityField = ctx.CastTo<TaxonomyField>(field);
            taxCityField.SspId = termStore.Id;
            taxCityField.TermSetId = termSet.Id;
            taxCityField.TargetTemplate = String.Empty;
            taxCityField.AnchorId = Guid.Empty;
            taxCityField.UpdateAndPushChanges(true);

            await ctx.ExecuteQueryAsync();

            Console.WriteLine("Finished!");
        }

        private static async Task SetDefaultValueForAboutField(ClientContext ctx)
        {
            Field field = ctx.Web.Fields.GetByTitle("hoabout");
            field.DefaultValue = "about default";
            field.UpdateAndPushChanges(true);

            ctx.Load(field, f => f.Title,
                            f => f.DefaultValue);

            await ctx.ExecuteQueryAsync();

            Console.WriteLine($"Successfully Set '{field.Title}' Site Field Default Value To '{field.DefaultValue}'!");
        }

        private static async Task SetDefaultValueForCityField(ClientContext ctx)
        {
            Field field = ctx.Web.Fields.GetByTitle("hocity");
            TaxonomyField taxCityField = ctx.CastTo<TaxonomyField>(field);
            var term = GetTermByName(ctx, "Ho Chi Minh");
            ctx.Load(term, t => t.Id,
                           t => t.Name);
            await ctx.ExecuteQueryAsync();

            TaxonomyFieldValue taxonomyFieldValue = new()
            {
                WssId = -1,
                Label = term.Name,
                TermGuid = term.Id.ToString()
            };
            ClientResult<string> value = taxCityField.GetValidatedString(taxonomyFieldValue);
            await ctx.ExecuteQueryAsync();

            taxCityField.DefaultValue = value.Value;
            taxCityField.UpdateAndPushChanges(true);

            ctx.Load(taxCityField, t => t.Title,
                            t => t.DefaultValue);
            await ctx.ExecuteQueryAsync();

            Console.WriteLine($"Successfully Set '{taxCityField.Title}' Site Field Default Value To '{taxCityField.DefaultValue}'!");
        }
        #endregion


        #region Content Types
        private static async Task CreateContentType(ClientContext ctx, string name)
        {
            ContentTypeCollection contentTypes = ctx.Web.ContentTypes;
            ctx.Load(contentTypes, cts => cts.Include(ct => ct.Name));
            await ctx.ExecuteQueryAsync();

            ContentType temp = contentTypes.FirstOrDefault(c => c.Name.Equals(name));
            if (temp != null)
                throw new Exception($"Content Type Name '{name}' Has Existed!");

            var creationInfo = new ContentTypeCreationInformation
            {
                Name = name,
                Group = "59Tese Content Types",
                ParentContentType = contentTypes.First(c => c.Name.Equals("Item"))
            };

            ContentType myContentType = contentTypes.Add(creationInfo);

            ctx.Load(myContentType, c => c.Name,
                                    c => c.Id);

            await ctx.ExecuteQueryAsync();

            Console.WriteLine($"Successfully Created '{myContentType.Name}' With ID '{myContentType.Id}'!");
        }

        private static async Task AddFieldsToContentType(ClientContext ctx)
        {
            ContentTypeCollection contentTypes = ctx.Web.ContentTypes;
            FieldCollection fields = ctx.Web.Fields;

            ctx.Load(contentTypes, cts => cts.Include(ct => ct.Name));
            ctx.Load(fields, fs => fs.Include(f => f.Group));

            await ctx.ExecuteQueryAsync();

            ContentType contentType = contentTypes.First(c => c.Name.Equals("CSOM Test content type"));
            var groupFields = fields.Where(f => f.Group.Equals("59Tese")).Select(f => f);

            foreach (Field field in groupFields)
            {
                var creationInfo = new FieldLinkCreationInformation
                {
                    Field = field
                };
                contentType.FieldLinks.Add(creationInfo);
            }
            contentType.Update(true);

            FieldLinkCollection ctFields = contentType.FieldLinks;
            ctx.Load(ctFields, fs => fs.Include(f => f.Name));
            await ctx.ExecuteQueryAsync();

            ctFields.First(f => f.Name.Equals("Title")).Hidden = true;
            contentType.Update(true);

            await ctx.ExecuteQueryAsync();

            Console.WriteLine("Finished!");
        }

        private static async Task AddCitiesFieldToContentType(ClientContext ctx)
        {
            ContentTypeCollection contentTypes = ctx.Web.ContentTypes;
            FieldCollection fields = ctx.Web.Fields;

            ctx.Load(contentTypes, cts => cts.Include(ct => ct.Name));
            ctx.Load(fields, fs => fs.Include(f => f.StaticName));

            await ctx.ExecuteQueryAsync();

            ContentType contentType = contentTypes.First(c => c.Name.Equals("CSOM Test content type"));
            Field myField = fields.First(f => f.StaticName.Equals("hocities"));

            var creationInfo = new FieldLinkCreationInformation
            {
                Field = myField
            };
            contentType.FieldLinks.Add(creationInfo);

            contentType.Update(true);

            await ctx.ExecuteQueryAsync();

            Console.WriteLine("Finished!");
        }
        #endregion


        #region CAML Query
        private static async Task GetListItems(ClientContext ctx)
        {
            List myList = ctx.Web.Lists.GetByTitle("CSOM Test");

            var items = myList.GetItems(new CamlQuery()
            {
                ViewXml = @$"<View>
                                <Query>
                                    <Where>
                                        <Neq>
                                            <FieldRef Name='hoabout'/>
                                            <Value Type='{FieldType.Text.ToString()}'>about default</Value>
                                        </Neq>
                                    </Where>
                                </Query>
                            </View>"
            });
            ctx.Load(items, its => its.Include(it => it.Id,
                                               it => it["hoabout"],
                                               it => it["hocity"]
                                               ));
            await ctx.ExecuteQueryAsync();

            foreach (var item in items)
            {
                TaxonomyFieldValue taxCityFieldValue = item["hocity"] as TaxonomyFieldValue;
                Console.WriteLine($"ID: {item.Id} \tAbout: {item["hoabout"]} \tCity: {taxCityFieldValue.Label}");
            }

            Console.WriteLine("Finished!");
        }

        private static async Task UpdateListItems(ClientContext ctx)
        {
            List myList = ctx.Web.Lists.GetByTitle("CSOM Test");

            var items = myList.GetItems(new CamlQuery()
            {
                ViewXml = @$"<View>
                                <Query>
                                    <Where>
                                        <Eq>
                                            <FieldRef Name='hoabout'/>
                                            <Value Type='{FieldType.Text.ToString()}'>about default</Value>
                                        </Eq>
                                    </Where>
                                </Query>
                                <RowLimit>2</RowLimit>
                            </View>"
            });
            ctx.Load(items, its => its.Include(it => it["hoabout"]));
            await ctx.ExecuteQueryAsync();

            if (items.Count > 0)
            {
                foreach (var item in items)
                {
                    item["hoabout"] = "Update script";
                    item.Update();
                }
                await ctx.ExecuteQueryAsync();
                Console.WriteLine("Finished!");
            }
            else Console.WriteLine("There Is No Item To Update!");
        }

        private static async Task GetListItemInFolderOnly(ClientContext ctx)
        {
            List myList = ctx.Web.Lists.GetByTitle("Document Test");

            Folder folder = ctx.Web.GetFolderByServerRelativeUrl(ctx.Web.ServerRelativeUrl + "/Document%20Test/Folder%201/Folder%202");
            ctx.Load(folder, f => f.ServerRelativeUrl);
            await ctx.ExecuteQueryAsync();

            var items = myList.GetItems(new CamlQuery()
            {
                ViewXml = @$"<View>
                                <Query>
                                    <Where>
                                        <Includes>
                                            <FieldRef Name='hocities'/>
                                            <Value Type='{FieldType.TaxonomyFieldTypeMulti.ToString()}'>Stockholm</Value>
                                        </Includes>
                                    </Where>
                                </Query>
                            </View>",
                FolderServerRelativeUrl = folder.ServerRelativeUrl
            });
            ctx.Load(items, its => its.Include(it => it.Id,
                                               it => it["hoabout"],
                                               it => it["hocities"]
                                               ));
            await ctx.ExecuteQueryAsync();

            foreach (var item in items)
            {
                TaxonomyFieldValueCollection taxCitiesFieldValues = item["hocities"] as TaxonomyFieldValueCollection;
                string res = $"ID: {item.Id} \tAbout: {item["hoabout"]} \tCities: ";
                foreach (var value in taxCitiesFieldValues)
                    res += "| " + value.Label + " |";
                Console.WriteLine(res);
            }
            Console.WriteLine("Finished!");
        }
        #endregion


        #region Other Methods
        // Load User From User Email Or Name
        private static async Task LoadUser(ClientContext ctx, string logonName)
        {
            ClientResult<PrincipalInfo> principal = Utility.ResolvePrincipal(ctx, ctx.Web, logonName, PrincipalType.User, PrincipalSource.All, ctx.Web.SiteUsers, true);
            await ctx.ExecuteQueryAsync();

            if (principal.Value != null)
            {
                var user = ctx.Web.SiteUsers.GetByEmail(principal.Value.Email);
                ctx.Load(user, u => u.LoginName,
                               u => u.Email,
                               u => u.Id,
                               u => u.Title);
                await ctx.ExecuteQueryAsync();

                Console.WriteLine($"Id: {user.Id} \nTitle: {user.Title} \nEmail: {user.Email} \nLoginName: {user.LoginName}");
            }
            else throw new Exception($"User With Logon Name '{logonName}' Not Found!");
        }

        // Load TaxonomyHiddenList Items
        private static async Task GetTaxonomyHiddenListItems(ClientContext ctx)
        {
            List myList = ctx.Web.Lists.GetByTitle("TaxonomyHiddenList");

            var items = myList.GetItems(CamlQuery.CreateAllItemsQuery());

            ctx.Load(items, its => its.Include(it => it["ID"],
                                               it => it["Title"],
                                               it => it["IdForTerm"]));
            await ctx.ExecuteQueryAsync();

            foreach (var item in items)
                Console.WriteLine($"ID: {item["ID"]} \tTitle: {item["Title"]} \tIdForTerm: {item["IdForTerm"]}");
        }

        // Load Site Users
        private static async Task LoadSiteUsers(ClientContext ctx)
        {
            var users = ctx.Web.SiteUsers;
            ctx.Load(users, us => us.Include(u => u.Id,
                                             u => u.Title,
                                             u => u.Email));
            await ctx.ExecuteQueryAsync();

            foreach (var user in users)
                Console.WriteLine($"Id: {user.Id} \tTitle: {user.Title} \tEmail: {user.Email}");

            Console.WriteLine($"\tTotal Users: {users.Count}");
        }

        // Remove Site Users That Was Deleted In Server
        private static async Task RemoveDeletedUser(ClientContext ctx)
        {
            var users = ctx.Web.SiteUsers;
            ctx.Load(users, us => us.Include(u => u.Email,
                                             u => u.LoginName));
            await ctx.ExecuteQueryAsync();

            foreach (var user in users)
            {
                if (user.Email != "" && user.Email != "CSOMTest@HenoldMK.onmicrosoft.com")
                {
                    ClientResult<PrincipalInfo> principal = Utility.ResolvePrincipal(ctx, ctx.Web, user.Email, PrincipalType.User, PrincipalSource.All, users, true);
                    await ctx.ExecuteQueryAsync();
                    if (principal.Value == null)
                    {
                        users.RemoveByLoginName(user.LoginName);
                        await ctx.ExecuteQueryAsync();
                        Console.WriteLine($"Successfully Deleted User With Login Name {user.LoginName}");
                    }
                }
            }
        }
        #endregion


        #region Permission


        #region Exercise 3
        private static async Task CreateSubSite(ClientContext ctx)
        {
            var creationInfo = new WebCreationInformation
            {
                Url = "FinanceAndAccounting",
                Title = "Finance And Accounting",
                Description = "Finance And Accounting subsite for Permission exercise",
                UseSamePermissionsAsParentSite = true,
                WebTemplate = "STS#3",
                Language = 1033
            };

            Web web = ctx.Web.Webs.Add(creationInfo);
            ctx.Load(web, w => w.Title,
                          w => w.Url,
                          web => web.Description);
            await ctx.ExecuteQueryAsync();

            Console.WriteLine($"Subsite Information:\n Title: {web.Title}\n URL: {web.Url}\n Description: {web.Description}");
        }

        private static async Task CreateListAtSubsite(ClientContext ctx, string title, ListTemplateType type)
        {
            await CreateList(ctx, title, type, await GetSubsite(ctx, "Finance And Accounting"));
        }

        private static async Task<Web> GetSubsite(ClientContext ctx, string title)
        {
            WebCollection webs = ctx.Web.Webs;
            ctx.Load(webs, ws => ws.Include(w => w.Title));
            await ctx.ExecuteQueryAsync();

            return webs.FirstOrDefault(w => w.Title.Equals(title));
        }

        private static async Task StopInheritingPermissions(ClientContext ctx)
        {
            Web subSite = await GetSubsite(ctx, "Finance And Accounting");

            List myList = subSite.Lists.GetByTitle("Accounts");
            ctx.Load(myList, ml => ml.HasUniqueRoleAssignments,
                             ml => ml.Title);
            await ctx.ExecuteQueryAsync();

            if (!myList.HasUniqueRoleAssignments)
            {
                myList.BreakRoleInheritance(false, false);
                myList.Update();

                await ctx.ExecuteQueryAsync();

                Console.WriteLine($"Finished! Stop Inheriting Permission From Its Parent!");
            }
            else throw new Exception($"List {myList.Title} Already Has Unique Permissions!");
        }

        private static async Task GrantDesignPermissionForUser(ClientContext ctx, string logonName, RoleType roleType)
        {
            Web subSite = await GetSubsite(ctx, "Finance And Accounting");

            List myList = subSite.Lists.GetByTitle("Accounts");
            ctx.Load(myList, ml => ml.HasUniqueRoleAssignments,
                             ml => ml.Title);
            await ctx.ExecuteQueryAsync();

            if (myList.HasUniqueRoleAssignments)
            {
                User user = ctx.Web.EnsureUser(logonName);
                ctx.Load(user, u => u.Email);

                try
                {
                    var listRoleDefinitionBinding = new RoleDefinitionBindingCollection(ctx)
                    {
                        ctx.Web.RoleDefinitions.GetByType(roleType)
                    };

                    var roleAssignment = myList.RoleAssignments.GetByPrincipal(user);
                    ctx.Load(roleAssignment, r => r.RoleDefinitionBindings);

                    await ctx.ExecuteQueryAsync();

                    var roleDefinition = roleAssignment.RoleDefinitionBindings.FirstOrDefault(r => r.RoleTypeKind.Equals(roleType));
                    if (roleDefinition == null)
                    {
                        myList.RoleAssignments.Add(user, listRoleDefinitionBinding);

                        await ctx.ExecuteQueryAsync();

                        Console.WriteLine("Finished!");
                    }
                    else throw new Exception($"User '{user.Email}' Already Have This Permission!");
                }
                catch (ServerException)
                {
                    var listRoleDefinitionBinding = new RoleDefinitionBindingCollection(ctx)
                    {
                        ctx.Web.RoleDefinitions.GetByType(roleType)
                    };

                    myList.RoleAssignments.Add(user, listRoleDefinitionBinding);

                    await ctx.ExecuteQueryAsync();

                    Console.WriteLine("Finished!");
                }
            }
            else throw new Exception($"List '{myList.Title}' Doesn't Have Uniqe Permission!");
        }

        private static async Task DeleteUniquePermissions(ClientContext ctx)
        {
            Web subSite = await GetSubsite(ctx, "Finance And Accounting");

            List myList = subSite.Lists.GetByTitle("Accounts");
            ctx.Load(myList, ml => ml.HasUniqueRoleAssignments,
                             ml => ml.Title);
            await ctx.ExecuteQueryAsync();

            if (myList.HasUniqueRoleAssignments)
            {
                myList.ResetRoleInheritance();
                myList.Update();

                await ctx.ExecuteQueryAsync();

                Console.WriteLine($"Finished! Re-establish Inheritance!");
            }
            else throw new Exception($"List {myList.Title} Already Has Inherited Permissions From Its Parent!");
        }
        #endregion


        #region Exercise 4
        private static async Task CreateCustomPermissionLevel(ClientContext ctx)
        {
            BasePermissions basePermissions = new();
            basePermissions.Set(PermissionKind.ManageLists);
            basePermissions.Set(PermissionKind.CreateAlerts);

            var creationInfo = new RoleDefinitionCreationInformation
            {
                Name = "Test Level",
                Description = "Custom Permission Level 'Test', granted 'manage lists' and 'create alerts' permissions",
                BasePermissions = basePermissions
            };

            var roleDefinition = ctx.Web.RoleDefinitions.Add(creationInfo);
            ctx.Load(roleDefinition, r => r.Name,
                                     r => r.Description);
            await ctx.ExecuteQueryAsync();

            Console.WriteLine($"Successfully Created Custom Permission Level\n Name: {roleDefinition.Name}\n Description: {roleDefinition.Description}");
        }

        private static async Task CreateCustomSecurityGroup(ClientContext ctx)
        {
            var creationInfo = new GroupCreationInformation
            {
                Title = "Test Group"
            };
            Group group = ctx.Web.SiteGroups.Add(creationInfo);
            
            await ctx.ExecuteQueryAsync();

            var siteRoleDefinitionBinding = new RoleDefinitionBindingCollection(ctx)
            {
                ctx.Web.RoleDefinitions.GetByName("Test Level")
            };
            ctx.Web.RoleAssignments.Add(group, siteRoleDefinitionBinding);

            await ctx.ExecuteQueryAsync();

            Console.WriteLine("Finished!");
        }

        private static async Task AddUserToSecurityGroup(ClientContext ctx, string logonName, string groupName)
        {
            Group group = ctx.Web.SiteGroups.GetByName(groupName);
            ctx.Load(group, g => g.Title);

            User user = ctx.Web.EnsureUser(logonName);
            ctx.Load(user, u => u.Email,
                           u => u.LoginName,
                           u => u.Title);
            await ctx.ExecuteQueryAsync();

            User addedUser = group.Users.Add(new UserCreationInformation()
            {
                Email = user.Email,
                LoginName = user.LoginName,
                Title = user.Title
            });
            ctx.Load(addedUser, u => u.Email);

            await ctx.ExecuteQueryAsync();

            Console.WriteLine($"Successfully Added User '{addedUser.Email}' To Group {group.Title}");
        }

        private static async Task CheckInheritedPermissionLevel(ClientContext ctx)
        {
            Web subSite = await GetSubsite(ctx, "Finance And Accounting");

            Group myGroup = subSite.SiteGroups.GetByName("Test Group");
            ctx.Load(myGroup, g => g.Title);

            var res = subSite.RoleAssignments.GetByPrincipal(myGroup);
            var roles = res.RoleDefinitionBindings;
            ctx.Load(roles, rs => rs.Include(r => r.Name));
            await ctx.ExecuteQueryAsync();

            foreach(var role in roles)
                if (role.Name.Equals("Test Level"))
                {
                    Console.WriteLine($"'{role.Name}' Has Been Inherited From The Root Site For The Security Group '{myGroup.Title}'");
                    return;
                }
        }
        #endregion


        #endregion


        #region User Profile
        private static async Task DisplaySomePropertiesForAllUsers(ClientContext ctx)
        {
            var builder = new ConfigurationBuilder().AddJsonFile($"appsettings.json", true, true);
            IConfiguration config = builder.Build();
            var info = config.GetSection("AzureAdInfo").Get<AzureAdInfo>();
            

            Microsoft.Graph.GraphServiceClient graphServiceClient = new(new ClientSecretCredential(info.TenantId, info.ClientId, info.ClientSecret));
            var users = graphServiceClient.Users.Request().Select(u => u.Mail).GetAsync().Result;

            PeopleManager peopleManager = new(ctx);

            foreach (var u in users)
            {
                var user = ctx.Web.EnsureUser(u.Mail);
                ctx.Load(user, u => u.LoginName);
                await ctx.ExecuteQueryAsync();

                var props = peopleManager.GetPropertiesFor(user.LoginName);
                ctx.Load(props, p => p.UserProfileProperties);
                await ctx.ExecuteQueryAsync();

                var userProfileProps = props.UserProfileProperties;
                Console.WriteLine($" Account Name: {userProfileProps["AccountName"]}\n" +
                                    $" First Name: {userProfileProps["FirstName"]}\n" +
                                    $" Last Name: {userProfileProps["LastName"]}\n" +
                                    $" Work Phone: {userProfileProps["WorkPhone"]}");
            }
        }

        private static async Task UpdateUserProperty(ClientContext ctx, string logonName)
        {
            User user = ctx.Web.EnsureUser(logonName);
            ctx.Load(user, u => u.LoginName);
            await ctx.ExecuteQueryAsync();

            PeopleManager peopleManager = new(ctx);
            var props = peopleManager.GetPropertiesFor(user.LoginName);
            ctx.Load(props, p => p.UserProfileProperties);
            await ctx.ExecuteQueryAsync();

            peopleManager.SetSingleValueProfileProperty(user.LoginName, "Henold-MiddleName", "Minh");
            await ctx.ExecuteQueryAsync();

            Console.WriteLine("Finished!");
        }
        #endregion


        #region KQL
        private static async Task KQLFilter(ClientContext ctx)
        {
            SearchExecutor searchExecutor = new(ctx);
            KeywordQuery keywordQuery = new(ctx)
            {
                QueryText = @"LastName=Phạm",
                EnableSorting = true,
                RowsPerPage = 100,
                RowLimit = 100,
                SourceId = new Guid("b09a7990-05ea-4af9-81ef-edfab16c4e31")
            };
            keywordQuery.SelectProperties.Add("LastName");
            keywordQuery.SelectProperties.Add("FirstName");
            keywordQuery.SelectProperties.Add("Henold-MiddleName");
            keywordQuery.SelectProperties.Add("WorkPhone");
            ClientResult<ResultTableCollection> results = searchExecutor.ExecuteQuery(keywordQuery);
            await ctx.ExecuteQueryAsync();

            if (results.Value[0].TotalRows == 0)
                Console.WriteLine("No Record Found!");
            else
            {
                foreach(var resultRow in results.Value[0].ResultRows)
                {
                    Console.WriteLine();
                    Console.WriteLine($" First Name: {resultRow["FirstName"]}\n " +
                        $"Middle Name: {resultRow["Henold-MiddleName"]}\n " +
                        $"Last Name: {resultRow["LastName"]}\n " +
                        $"Work Phone: {resultRow["WorkPhone"]}");
                }
            }
        }
        #endregion
    }
}

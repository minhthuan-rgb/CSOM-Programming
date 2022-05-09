using Microsoft.Extensions.Configuration;
using Microsoft.SharePoint.Client;
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

                    Console.WriteLine($"Connected Site: {ctx.Web.Title}");

                    // Create List "CSOM Test"
                    //await CreateList(ctx, "Test", ListTemplateType.Announcements);

                    // Delete List "CSOM Test"
                    //await DeleteList(ctx, "CSOM Test");

                    // Create Term Set And 2 Terms
                    //await CreateTermSetAndTerms(ctx);

                    // Create 2 site fields: "about" type text and "city" type taxonomy
                    //await CreateTextField(ctx, "about");
                    //await CreateTaxonomyField(ctx, "city");

                    // Create Content Type "CSOM Test content type"
                    //await CreateContentType(ctx, "CSOM Test content type");

                    // Add Content Type To List "CSOM Test"
                    //await AddContentTypeToList(ctx, "CSOM Test");

                    // Add 2 Fields "about" And "city" To Content Type "CSOM Test content type"
                    //await AddFieldsToContentType(ctx);

                    // Set "CSOM Test content type" As Default Content Type In List "CSOM test"
                    //await SetDefaultContentType(ctx, "CSOM Test");

                    // Bind Taxonomy Field "city" To Term Set
                    //await BindTaxonomyFieldToTermSet(ctx, "city");

                    // Display All Items List View
                    //await DisplayAllItemsListView(ctx);

                    // Add 5 Items To List "CSOM Test"
                    //await AddItemsToList(ctx, 5);

                    // Set Default Value For "about" site field
                    //await SetDefaultValueForAboutField(ctx);

                    // Add 2 Items With Default Value Of "about" site field
                    //await AddItemsToList(ctx, 2, true);

                    // Set Default Value For "city" site field
                    //await SetDefaultValueForCityField(ctx);

                    // Add 2 Items With Default Value Of "city" site field
                    //await AddItemsToList(ctx, 2, true, true);

                    // Get List Items Where Field "about" is not "about default"
                    //await GetListItems(ctx);

                    // Create List View by CSOM 
                    //await CreateListView(ctx);

                    // Update List View Items
                    //await UpdateListItems(ctx);

                    // Create People Field In List "CSOM Test"
                    //await CreatePeopleField(ctx);

                    // Migrate all list items to set user admin to field "author"
                    //await SetUserAdminToAuthorField(ctx);

                    // Create site field "cities" type taxonomy multi values
                    //await CreateTaxonomyFieldMulti(ctx, "cities");

                    // Bind Taxonomy Field "cities" To Term Set
                    //await BindTaxonomyFieldToTermSet(ctx, "cities");

                    // Add Field "cities" To Content Type "CSOM Test Content Type"
                    //await AddCitiesFieldToContentType(ctx);

                    // Add 3 Items With Field "Cities" Multi Value
                    //await AddItemsToList(ctx, 3, true, true, true);

                    // Display "cities" Field 
                    //await DisplayCitiesField(ctx);

                    // Create List "Document Test"
                    //await CreateList(ctx, "Document Test", ListTemplateType.DocumentLibrary);

                    // Add Content Type TO List "Document Test"
                    //await AddContentTypeToList(ctx, "Document Test");

                    // Create Folders For List "Document Test"
                    //await CreateFolders(ctx);

                    // Set "CSOM Test content type" As Default Content Type In List "Document Test"
                    //await SetDefaultContentType(ctx, "Document Test");

                    // Add 3 Files In "Folder 2" With Value "Folder test" In Field "about"
                    //await AddFilesInsideFolder(ctx, 3, "FolderTest");

                    // Add 2 Files In "Folder 2" With Value "Stockholm" In Field "cities"
                    //await AddFilesInsideFolder(ctx, 2, "CitiesTest", true);

                    // Get All List Items Just In "Folder 2" And Have Value "Stockholm" in "cities" field
                    //await GetListItemInFolderOnly(ctx);

                    // Create List Item In "Document Test" By Upload A File Document.docx
                    //await CreateListItemByUploadFile(ctx);

                    // Display All Documents List View
                    //await DisplayAllDocumentsListView(ctx);

                    // Create Folder Structure View
                    await CreateFolderStructureView(ctx);

                    // Load User From User Email Or Name
                    //await LoadUser(ctx, "59Tese");
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
        private static async Task CreateList(ClientContext ctx, string title, ListTemplateType type)
        {
            var creationInfo = new ListCreationInformation();
            creationInfo.Title = title;
            creationInfo.TemplateType = (int)type;
            List list = ctx.Web.Lists.Add(creationInfo);

            list.Description = $"This is {title} List, that was created from client side";
            list.OnQuickLaunch = true;
            list.Update();

            await ctx.ExecuteQueryAsync();

            Console.WriteLine($"Successfully Created {list.Title} List!");
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
                newItem["about"] = $"Item {about}";

            if (!isCityDefault)
            {
                var cityField = ctx.Web.Fields.GetByTitle("city");

                var taxCityField = ctx.CastTo<TaxonomyField>(cityField);

                taxCityField.SetFieldValueByValue(newItem, new TaxonomyFieldValue()
                {
                    WssId = -1,
                    Label = "Stockholm",
                    TermGuid = "9661521e-608d-42c5-83f6-e96c674a32db"
                });

                //var clientRuntimeContext = newItem.Context;
                //var field = myList.Fields.GetByTitle("city");
                //var taxCityField = clientRuntimeContext.CastTo<TaxonomyField>(field);
            }

            if (isMulti)
            {
                var citiesField = ctx.Web.Fields.GetByTitle("cities");
                var taxCitiesField = ctx.CastTo<TaxonomyField>(citiesField);

                string fieldValues = "1;#Stockholm|9661521e-608d-42c5-83f6-e96c674a32db;#2;#Ho Chi Minh|8d6eea46-5de6-441c-9740-aca1928ba368";
                taxCitiesField.SetFieldValueByValueCollection(newItem, new TaxonomyFieldValueCollection(ctx, fieldValues, taxCitiesField));
            }

            newItem.Update();
            await ctx.ExecuteQueryAsync();

            Console.WriteLine($"Created Item!");
        }

        private static async Task CreateListView(ClientContext ctx)
        {
            List myList = ctx.Web.Lists.GetByTitle("CSOM Test");
            ctx.Load(myList, ml => ml.Title);

            ViewCollection views = myList.Views;
            ctx.Load(views, vs => vs.Include(v => v.Title));
            await ctx.ExecuteQueryAsync();

            View temp = views.FirstOrDefault(view => view.Title.Equals("CSOM Test View"));

            if (temp != null)
            {
                Console.WriteLine($"List View '{temp.Title}' has existed!");
                return;
            }

            var creationInfo = new ViewCreationInformation();
            creationInfo.Title = "CSOM Test View";
            //creationInfo.SetAsDefaultView = true;
            creationInfo.ViewTypeKind = ViewType.Html;
            creationInfo.Query = "<OrderBy>" +
                                    "<FieldRef Name='Created' Ascending='False'/>" +
                                "</OrderBy>" +
                                "<Where>" +
                                    "<Eq>" +
                                        "<FieldRef Name='city'/>" +
                                        "<Value Type = 'Taxonomy'>Ho Chi Minh</Value>" +
                                    "</Eq>" +
                                "</Where>";
            string commaSeparateColumnNames = "ID, Title, about, city, Created";
            creationInfo.ViewFields = commaSeparateColumnNames.Split(", ");

            View listView = views.Add(creationInfo);
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
            allItemsView.ViewFields.Add("about");
            allItemsView.ViewFields.Add("city");

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

            allItemsView.ViewFields.Add("cities");
            allItemsView.Update();

            await ctx.ExecuteQueryAsync();

            Console.WriteLine($"Successfully Add 'cities' Field To List View '{allItemsView.Title}' For List '{myList.Title}'!");
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

            await CreateField(ctx, "authorr", FieldType.User, myList.Fields);
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
                                    <FieldRef Name='authorr'/>
                                </ViewFields>
                            </View>"
            });
            ctx.Load(items, its => its.Include(it => it["authorr"]));

            User user = ctx.Web.EnsureUser(await LoadCurrentUserEmail(ctx));
            ctx.Load(user, u => u.Id);

            await ctx.ExecuteQueryAsync();

            FieldUserValue userValue = new FieldUserValue()
            {
                LookupId = user.Id,
            };

            foreach (var item in items)
            {
                item["authorr"] = userValue;
                item.Update();
            }

            await ctx.ExecuteQueryAsync();

            Console.WriteLine("Finished!");
        }

        private static async Task<string> LoadCurrentUserEmail(ClientContext ctx)
        {
            //User currentUser = ctx.Web.CurrentUser; 

            PeopleManager peopleManager = new PeopleManager(ctx);
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
            List myList = ctx.Web.Lists.GetByTitle("Document Test");

            var folder = myList.RootFolder;
            folder = folder.Folders.GetByUrl("Folder 1");
            folder = folder.Folders.GetByUrl("Folder 2");

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
            {
                Console.WriteLine($"'{title}.docx' has existed!");
                return;
            }

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

            if (isCities)
            {
                var citiesField = ctx.Web.Fields.GetByTitle("cities");
                var taxCitiesField = ctx.CastTo<TaxonomyField>(citiesField);

                string fieldValues = "1;#Stockholm|9661521e-608d-42c5-83f6-e96c674a32db";
                taxCitiesField.SetFieldValueByValueCollection(newItem, new TaxonomyFieldValueCollection(ctx, fieldValues, taxCitiesField));
            }
            else newItem["about"] = "Folder Test";

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

            allDocumentsView.ViewFields.Add("about");
            allDocumentsView.ViewFields.Add("city");
            allDocumentsView.ViewFields.Add("cities");

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
            {
                Console.WriteLine($"List View '{temp.Title}' has existed!");
                return;
            }

            var creationInfo = new ViewCreationInformation();
            creationInfo.Title = "Folders";
            creationInfo.ViewTypeKind = ViewType.Html;
            creationInfo.Query = "<Where>" +
                                    "<Eq>" +
                                        "<FieldRef Name='FSObjType'/>" +
                                        "<Value Type='Integer'>1</Value>" +
                                    "</Eq>" +
                                "</Where>";
            string commaSeparateColumnNames = "Type, Name, Modified, Modified By";
            creationInfo.ViewFields = commaSeparateColumnNames.Split(", ");

            View listView = views.Add(creationInfo);
            listView.Scope = ViewScope.RecursiveAll;
            //listView.DefaultView = true;
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
            TermGroup termGroup = termStore.CreateGroup("Test", Guid.NewGuid());

            // Create Term Set
            TermSet termSet = termGroup.CreateTermSet("city-PhamMinhThuan", Guid.NewGuid(), 1033);

            // Create 2 Terms
            Term hcm = termSet.CreateTerm("Ho Chi Minh", 1033, Guid.NewGuid());

            Term stockholm = termSet.CreateTerm("Stockholm", 1033, Guid.NewGuid());

            await ctx.ExecuteQueryAsync();

            Console.WriteLine("Finished!");
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
            Field field = ctx.Web.Fields.GetByTitle("about");
            field.DefaultValue = "about default";
            field.UpdateAndPushChanges(true);

            ctx.Load(field, f => f.Title,
                            f => f.DefaultValue);

            await ctx.ExecuteQueryAsync();

            Console.WriteLine($"Successfully set '{field.Title}' site field default value to '{field.DefaultValue}'!");
        }

        private static async Task SetDefaultValueForCityField(ClientContext ctx)
        {
            Field field = ctx.Web.Fields.GetByTitle("city");
            TaxonomyField taxCityField = ctx.CastTo<TaxonomyField>(field);
            taxCityField.DefaultValue = "2;#Ho Chi Minh|8d6eea46-5de6-441c-9740-aca1928ba368";
            taxCityField.UpdateAndPushChanges(true);

            ctx.Load(taxCityField, t => t.Title,
                            t => t.DefaultValue);

            await ctx.ExecuteQueryAsync();

            Console.WriteLine($"Successfully set '{taxCityField.Title}' site field default value to '{taxCityField.DefaultValue}'!");
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
            {
                Console.WriteLine($"Content type name '{name}' has existed!");
                return;
            }

            var creationInfo = new ContentTypeCreationInformation();
            creationInfo.Name = name;
            creationInfo.Group = "59Tese Content Types";
            creationInfo.ParentContentType = contentTypes.First(c => c.Name.Equals("Item"));

            ContentType myContentType = contentTypes.Add(creationInfo);

            ctx.Load(myContentType, c => c.Name,
                                    c => c.Id);

            await ctx.ExecuteQueryAsync();

            Console.WriteLine($"Successfully Created {myContentType.Name} with ID {myContentType.Id}!");
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
                var creationInfo = new FieldLinkCreationInformation();
                creationInfo.Field = field;
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
            Field myField = fields.First(f => f.StaticName.Equals("cities"));

            var creationInfo = new FieldLinkCreationInformation();
            creationInfo.Field = myField;
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
                                            <FieldRef Name='about'/>
                                            <Value Type='{FieldType.Text.ToString()}'>about default</Value>
                                        </Neq>
                                    </Where>
                                </Query>
                            </View>"
            });
            ctx.Load(items, its => its.Include(it => it.Id,
                                               it => it["about"],
                                               it => it["city"]
                                               ));
            await ctx.ExecuteQueryAsync();

            foreach (var item in items)
            {
                TaxonomyFieldValue taxCityFieldValue = item["city"] as TaxonomyFieldValue;
                Console.WriteLine($"ID: {item.Id} \tAbout: {item["about"]} \tCity: {taxCityFieldValue.Label}");
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
                                            <FieldRef Name='about'/>
                                            <Value Type='{FieldType.Text.ToString()}'>about default</Value>
                                        </Eq>
                                    </Where>
                                </Query>
                                <RowLimit>2</RowLimit>
                            </View>"
            });
            ctx.Load(items, its => its.Include(it => it["about"]));
            await ctx.ExecuteQueryAsync();

            if (items.Count > 0)
            {
                foreach (var item in items)
                {
                    item["about"] = "Update script";
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

            //Folder folder = ctx.Web.GetFolderByServerRelativeUrl(ctx.Web.ServerRelativeUrl + "/Document%20Test/Folder%201/Folder%202");
            var folder = myList.RootFolder;
            folder = folder.Folders.GetByUrl("Folder 1");
            folder = folder.Folders.GetByUrl("Folder 2");
            ctx.Load(folder, f => f.ServerRelativeUrl);
            await ctx.ExecuteQueryAsync();

            var items = myList.GetItems(new CamlQuery()
            {
                ViewXml = @$"<View>
                                <Query>
                                    <Where>
                                        <Includes>
                                            <FieldRef Name='cities'/>
                                            <Value Type='{FieldType.TaxonomyFieldTypeMulti.ToString()}'>Stockholm</Value>
                                        </Includes>
                                    </Where>
                                </Query>
                            </View>",
                FolderServerRelativeUrl = folder.ServerRelativeUrl
            });
            ctx.Load(items, its => its.Include(it => it.Id,
                                               it => it["about"],
                                               it => it["cities"]
                                               ));
            await ctx.ExecuteQueryAsync();

            foreach (var item in items)
            {
                TaxonomyFieldValueCollection taxCitiesFieldValues = item["cities"] as TaxonomyFieldValueCollection;
                string res = $"ID: {item.Id} \tAbout: {item["about"]} \tCities: ";
                foreach (var value in taxCitiesFieldValues)
                    res += "| " + value.Label + " |";
                Console.WriteLine(res);
            }
            Console.WriteLine("Finished!");
        }
        #endregion


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
                               u => u.Id);
                await ctx.ExecuteQueryAsync();

                Console.WriteLine($"Id: {user.Id} \nEmail: {user.Email} \nLoginName: {user.LoginName}");
            }
            else Console.WriteLine($"User With Logon Name '{logonName}' Not Found!");
        }
    }
}

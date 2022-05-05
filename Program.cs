using Microsoft.Extensions.Configuration;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using System;
using System.Collections.Generic;
using System.Linq;
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

                    // Create List CSOM Test
                    //await CreateList(ctx, "CSOM Test");

                    // Delete List CSOM Test
                    //await DeleteList(ctx, "CSOM Test");

                    // Create Term Set And 2 Terms
                    //await CreateTermSetAndTerms(ctx);

                    // Create 2 site fields: "about" type text and "city" type taxonomy
                    //await CreateTextField(ctx, "about");
                    //await CreateTaxonomyField(ctx, "city");

                    // Create Content Type "CSOM Test content type"
                    //await CreateContentType(ctx, "CSOM Test content type");

                    // Add Content Type To List CSOM Test
                    //await AddContentTypeToList(ctx);

                    // Add 2 Fields "about" And "city" To Content Type "CSOM Test content type"
                    //await AddFieldsToContentType(ctx);

                    // Set "CSOM Test content type" As Default Content Type In List "CSOM test"
                    //await SetDefaultContentType(ctx);

                    // Bind Taxonomy Field To Term Set
                    //await BindTaxonomyFieldToTermSet(ctx);

                    // Add 5 Items To List CSOM Test
                    //await AddItemsToList(ctx, 5);

                    // Set Default Value For 'about' site field
                    //await SetDefaultValueForAboutField(ctx);

                    // Add 2 Items With Default Value Of 'about' site field
                    //await AddItemsToList(ctx, 2, true);

                    // Set Default Value For 'city' site field
                    await SetDefaultValueForCityField(ctx);

                    // Add 2 Items With Default Value Of 'city' site field
                    await AddItemsToList(ctx, 2, true, true);
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
        private static async Task CreateList(ClientContext ctx, string title)
        {
            ListCreationInformation creationInfo = new ListCreationInformation();
            creationInfo.Title = title;
            creationInfo.TemplateType = (int)ListTemplateType.Announcements;
            List list = ctx.Web.Lists.Add(creationInfo);

            list.Description = $"This is {title} List, that was created from client side";
            list.OnQuickLaunch = true;
            list.Update();
            
            await ctx.ExecuteQueryAsync();

            Console.WriteLine($"Successfully Created {list.Title} List!");
        }

        private static async Task AddContentTypeToList(ClientContext ctx)
        {
            List myList = ctx.Web.Lists.GetByTitle("CSOM Test");
            ctx.Load(myList, list => list.ContentTypesEnabled);
            await ctx.ExecuteQueryAsync();

            if(!myList.ContentTypesEnabled)
            {
                myList.ContentTypesEnabled = true;
                myList.Update();
                await ctx.ExecuteQueryAsync();
            }

            ContentTypeCollection contentTypes = ctx.Web.ContentTypes;
            ctx.Load(contentTypes);
            await ctx.ExecuteQueryAsync();

            ContentType contentType = contentTypes.Where(c => c.Name.Equals("CSOM Test content type")).First();
            myList.ContentTypes.AddExistingContentType(contentType);
            myList.Update();

            await ctx.ExecuteQueryAsync();

            Console.WriteLine("Finished!");
        }

        private static async Task SetDefaultContentType(ClientContext ctx)
        {
            List myList = ctx.Web.Lists.GetByTitle("CSOM Test");
            ContentTypeCollection contentTypes = myList.ContentTypes;
            ctx.Load(contentTypes);
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

        private static async Task AddItemsToList(ClientContext ctx, int amount, bool isAboutDefault = false, bool isCityDefault = false)
        {
            for (int i = 0; i < amount; i++)
            {
                await AddItemToList(ctx, isAboutDefault ? null : (i+1).ToString(), isCityDefault);
            }
        }

        private static async Task AddItemToList(ClientContext ctx, string about = null, bool isCityDefault = false)
        {
            List myList = ctx.Web.Lists.GetByTitle("CSOM Test");

            ListItemCreationInformation creationInfo = new ListItemCreationInformation();
            ListItem newItem = myList.AddItem(creationInfo);
            if (about != null)
                newItem["about"] = $"Item {about}";

            if (!isCityDefault)
            {
                var field = ctx.Web.Fields.GetByTitle("city");

                var taxCityField = ctx.CastTo<TaxonomyField>(field);

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
            newItem.Update();
            await ctx.ExecuteQueryAsync();

            Console.WriteLine($"Created Item!");
        }

        private static async Task DisplayListColumns(ClientContext ctx)
        {
            List myList = ctx.Web.Lists.GetByTitle("CSOM Test");

            await ctx.ExecuteQueryAsync();
        }

        private static async Task DeleteList(ClientContext ctx, string title)
        {
            List list = ctx.Web.Lists.GetByTitle(title);
            ctx.Load(list, l => l.Title);
            list.DeleteObject();

            await ctx.ExecuteQueryAsync();

            Console.WriteLine($"Successfully Deleted {list.Title} List!");
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
        }
        #endregion


        #region Fields
        private static async Task CreateField(ClientContext ctx, string name, string type)
        {
            string schemaField = $"<Field ID='{Guid.NewGuid()}' Type='{type}' Name='{name}' StaticName='{name}' DisplayName='{name}' Group='59Tese'/>";
            Field field = ctx.Web.Fields.AddFieldAsXml(schemaField, true, AddFieldOptions.AddFieldInternalNameHint);

            ctx.Load(field, f => f.StaticName);

            await ctx.ExecuteQueryAsync();

            Console.WriteLine($"Successfully Created {field.StaticName} Field!");
        }

        private static async Task CreateTextField(ClientContext ctx, string name)
        {
            await CreateField(ctx, name, "Text");
        }

        private static async Task CreateTaxonomyField(ClientContext ctx, string name)
        {
            await CreateField(ctx, name, "TaxonomyFieldType");
        }

        private static async Task BindTaxonomyFieldToTermSet(ClientContext ctx)
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

            Field field = ctx.Web.Fields.GetByTitle("city");
            TaxonomyField taxCityField = ctx.CastTo<TaxonomyField>(field);
            taxCityField.SspId = termStore.Id;
            taxCityField.TermSetId = termSet.Id;
            taxCityField.TargetTemplate = String.Empty;
            taxCityField.AnchorId = Guid.Empty;
            taxCityField.UpdateAndPushChanges(true);

            await ctx.ExecuteQueryAsync();

            Console.WriteLine("Finished!");
        } 

        private static async Task SetDefaultValueForAboutField (ClientContext ctx)
        {
            Field field = ctx.Web.Fields.GetByTitle("about");
            field.DefaultValue = "about default";
            field.UpdateAndPushChanges(true);

            ctx.Load(field, f => f.Title, 
                            f => f.DefaultValue);

            await ctx.ExecuteQueryAsync();

            Console.WriteLine($"Successfully set '{field.Title}' site field default value to '{field.DefaultValue}'");
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

            Console.WriteLine($"Successfully set '{taxCityField.Title}' site field default value to '{taxCityField.DefaultValue}'");
        }
        #endregion


        #region Content Types
        private static async Task CreateContentType(ClientContext ctx, string name)
        {
            ContentTypeCollection contentTypes = ctx.Web.ContentTypes;
            ctx.Load(contentTypes);
            await ctx.ExecuteQueryAsync();

            ContentType temp = contentTypes.Where(c => c.Name.Equals(name)).FirstOrDefault();
            if (temp != null)
            {
                Console.WriteLine($"Content type name '{name}' has already exists!");
                return;
            }

            ContentTypeCreationInformation creationInfo = new ContentTypeCreationInformation();
            creationInfo.Name = name;
            creationInfo.Group = "59Tese Content Types";
            creationInfo.ParentContentType = contentTypes.Where(c => c.Name.Equals("Item")).First();

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
            ctx.Load(contentTypes);
            ctx.Load(fields);
            await ctx.ExecuteQueryAsync();

            ContentType contentType = contentTypes.Where(c => c.Name.Equals("CSOM Test content type")).First();
            var groupFields = fields.Where(f => f.Group.Equals("59Tese")).Select(f => f);

            foreach (Field field in groupFields)
            {
                FieldLinkCreationInformation creationInfo = new FieldLinkCreationInformation();
                creationInfo.Field = field;
                contentType.FieldLinks.Add(creationInfo);
            }
            contentType.Update(true);

            FieldLinkCollection ctFields = contentType.FieldLinks;
            ctx.Load(ctFields);
            await ctx.ExecuteQueryAsync();

            ctFields.Where(f => f.Name.Equals("Title")).First().Hidden = true;
            contentType.Update(true);

            await ctx.ExecuteQueryAsync();

            Console.WriteLine("Finished!");
        }
        #endregion
    }
}

using Microsoft.Extensions.Configuration;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using System;
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

                    // Create Term Set And 2 Terms
                    await CreateTermSetAndTerms(ctx);
                }

                Console.WriteLine($"Press Any Key To Stop!");
                Console.ReadKey();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        static ClientContext GetContext(ClientContextHelper clientContextHelper)
        {
            var builder = new ConfigurationBuilder().AddJsonFile($"appsettings.json", true, true);
            IConfiguration config = builder.Build();
            var info = config.GetSection("SharepointInfo").Get<SharepointInfo>();
            return clientContextHelper.GetContext(new Uri(info.SiteUrl), info.Username, info.Password);
        }

        private static async Task CreateList(ClientContext ctx, string title)
        {
            ListCreationInformation creationInfo = new ListCreationInformation();
            creationInfo.Title = title;
            creationInfo.TemplateType = (int)ListTemplateType.Announcements;
            List list = ctx.Web.Lists.Add(creationInfo);

            list.Description = $"This list is {title} List, that was created at client side";
            list.OnQuickLaunch = true;
            list.Update();
            
            await ctx.ExecuteQueryAsync();

            Console.WriteLine($"Successfully Created {list.Title} List");
        }

        private static async Task DeleteList(ClientContext ctx, string title)
        {
            List list = ctx.Web.Lists.GetByTitle(title);
            list.DeleteObject();

            await ctx.ExecuteQueryAsync();
        }

        private static async Task CreateTermSetAndTerms(ClientContext ctx)
        {
            TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(ctx);
            TermStore termStore = taxonomySession.GetDefaultSiteCollectionTermStore();

            // Create Group Test
            TermGroup termGroup = termStore.CreateGroup("Test", Guid.NewGuid());

            // Create Term Set
            TermSet termSet = termGroup.CreateTermSet("city-PhamMinhThuan", Guid.NewGuid(), 1316);

            // Create 2 Terms
            termSet.CreateTerm("Ho Chi Minh", 1399, Guid.NewGuid());

            termSet.CreateTerm("Stockholm", 1699, Guid.NewGuid());

            await ctx.ExecuteQueryAsync();
        }
    }
}

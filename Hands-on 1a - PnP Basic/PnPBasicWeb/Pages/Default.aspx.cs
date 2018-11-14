using System;
using Microsoft.SharePoint.Client;

namespace PnPBasicWeb
{
    public partial class Default : System.Web.UI.Page
    {
        protected void Page_PreInit(object sender, EventArgs e)
        {
            Uri redirectUrl;
            switch (SharePointContextProvider.CheckRedirectionStatus(Context, out redirectUrl))
            {
                case RedirectionStatus.Ok:
                    return;
                case RedirectionStatus.ShouldRedirect:
                    Response.Redirect(redirectUrl.AbsoluteUri, endResponse: true);
                    break;
                case RedirectionStatus.CanNotRedirect:
                    Response.Write("An error occurred while processing your request.");
                    Response.End();
                    break;
            }
        }

        protected void Page_Load(object sender, EventArgs e)
        {
            HyperLink1.NavigateUrl = SharePointContext.GetSPHostUrl(Context.Request).ToString();
        }

        protected void CSOMButton_Click(object sender, EventArgs e)
        {
            var spContext = SharePointContextProvider.Current.GetSharePointContext(Context);

            using (var clientContext = spContext.CreateUserClientContextForSPHost())
            {
                Web web = clientContext.Web;

                // Create Content Type on host web
                ContentTypeCreationInformation newCt = new ContentTypeCreationInformation();
                newCt.Name = "CSOM Item";
                newCt.Id = "0x010078874F9C61114245806D6F09BC0362F8";
                newCt.Group = "A Lab";
                ContentType myContentType = web.ContentTypes.Add(newCt);
                FieldCollection fields = web.Fields;
                clientContext.Load(fields);
                clientContext.ExecuteQuery();

                // Add field to content type
                Field field = fields.GetByInternalNameOrTitle("Categories");
                FieldLinkCreationInformation link = new FieldLinkCreationInformation();
                link.Field = field;
                myContentType.FieldLinks.Add(link);
                myContentType.Update(true);
                clientContext.ExecuteQuery();

                // Create list on host web
                var list = web.Lists.Add(new ListCreationInformation
                {
                    Title = "CSOM",
                    Url = "lists/csom",
                    TemplateType = (int)ListTemplateType.GenericList
                });
                clientContext.Load(list);
                clientContext.ExecuteQuery();

                // Add Content Type to list
                list.ContentTypesEnabled = true;
                list.Update();
                list.ContentTypes.AddExistingContentType(myContentType);
                clientContext.ExecuteQuery();

                // Add a view to list
                var view = list.Views.Add(new ViewCreationInformation
                {
                    Title = "New View",
                    ViewTypeKind = ViewType.Html,
                    RowLimit = 10,
                    ViewFields = new[] { "Title", "Categories" },
                    SetAsDefaultView = true,
                    Paged = false
                });
                clientContext.Load(view);
                clientContext.ExecuteQuery();

                // Set a property bag value
                var properties = clientContext.Web.AllProperties;
                clientContext.Load(properties);
                clientContext.ExecuteQuery();
                properties["ourwebkey"] = Guid.NewGuid().ToString();
                clientContext.Web.Update();
                clientContext.ExecuteQuery();

                // Make property bag value searchable
                // TBD
            }
        }
    }
}
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core;
using System.Net.Http;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using System;

namespace SharepointTemplate
{
    public static class SharePointFileMover
    {
        [FunctionName("SharePointFileMover")]
        public static void Run([HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)]HttpRequestMessage req, TraceWriter log)
        {
            log.Info("SharePoint File Mover Started");

            try
            {
                bool success = RunFileMover();
            }
            catch (Exception ex) {

                log.Info(ex.Message.ToString());
                log.Info(ex.StackTrace.ToString());
            }

            log.Info("SharePoint File Mover Completed");

        }

        private static bool RunFileMover()
        {
            string url = "https://chrishanna.sharepoint.com/sites/Chaxa/";
            string appId = ""; //replace it to your app id. 
            string appScret = ""; //replace it to your app secret. 
            var documentLibraryName = "Dropzone";

            AuthenticationManager manager = new AuthenticationManager();

            using (ClientContext context = manager.GetAppOnlyAuthenticatedContext(url, appId, appScret))
            {
                Web web = context.Web;
                context.Load(web);

                List list = web.Lists.GetByTitle(documentLibraryName);

                CamlQuery query = new CamlQuery();

                //string datelimit = DateTime.UtcNow.AddMinutes(-1).ToString("yyyy-MM-ddTHH:mm:ssZ");
                //query.ViewXml = string.Format(@"<View><Query><Where><Gt><FieldRef Name='Modified' /><Value Type='DateTime'>{0}</Value></Gt></Where></Query></View>",datelimit);

                ListItemCollection items = list.GetItems(query);
                context.Load(list.Fields);
                context.Load(list.RootFolder);
                context.Load(list.RootFolder.Folders);
                context.Load(list.RootFolder.Files);

                context.Load(list.RootFolder.Files,
                       files => files.Include(file => file.Name),
                       files => files.Include(file => file.ETag),
                       files => files.Include(file => file.TimeLastModified),
                       files => files.Include(file => file.Properties),
                       files => files.Include(file => file.ListItemAllFields["Company"]),
                       files => files.Include(file => file.ListItemAllFields["CompliantReport"]),
                       files => files.Include(file => file.ListItemAllFields["DocumentClass"]),
                       files => files.Include(file => file.ListItemAllFields["DocumentSubType"]),
                       files => files.Include(file => file.ListItemAllFields["DocumentType"]),
                       files => files.Include(file => file.ListItemAllFields["Format"]),
                       files => files.Include(file => file.ListItemAllFields["Language"]),
                       files => files.Include(file => file.ListItemAllFields["Month"]),
                       files => files.Include(file => file.ListItemAllFields["PlanYear"]),
                       files => files.Include(file => file.ListItemAllFields["ProductType"])
                );

                context.ExecuteQuery();

                //FolderCollection fcol = list.RootFolder.Folders;
                FileCollection filecol = list.RootFolder.Files;

                HandleFileMove(filecol, web);

                return true;

            }
        }

        public static void HandleFileMove(FileCollection fileCol, Web web)
        {
                    foreach (File file in fileCol)
                    {
                        string path = "";
                        
                        path = file.ListItemAllFields.FieldValues["Company"] != null ? path + "/" + file.ListItemAllFields.FieldValues["Company"].ToString() : path;
                        path = file.ListItemAllFields.FieldValues["PlanYear"] != null ? path + "/" + file.ListItemAllFields.FieldValues["PlanYear"].ToString() : path;
                        path = file.ListItemAllFields.FieldValues["DocumentClass"] != null ? path + "/" + file.ListItemAllFields.FieldValues["DocumentClass"].ToString() : path;
                        path = file.ListItemAllFields.FieldValues["DocumentType"] != null   ? path + "/" + file.ListItemAllFields.FieldValues["DocumentType"].ToString() : path;
                        path = file.ListItemAllFields.FieldValues["Month"] != null ? path + "/" + file.ListItemAllFields.FieldValues["Month"].ToString() : path;
                        path = file.ListItemAllFields.FieldValues["DocumentSubType"] != null ? path + "/" + file.ListItemAllFields.FieldValues["DocumentSubType"].ToString() : path;
                        path = file.ListItemAllFields.FieldValues["ProductType"] != null ? path + "/" + file.ListItemAllFields.FieldValues["ProductType"].ToString() : path;
                        path = file.ListItemAllFields.FieldValues["Language"] != null ? path + "/" + file.ListItemAllFields.FieldValues["Language"].ToString() : path;
                        path = file.ListItemAllFields.FieldValues["Format"] != null ? path + "/" + file.ListItemAllFields.FieldValues["Format"].ToString() : path;
                        path = file.ListItemAllFields.FieldValues["CompliantReport"] != null ? path + "/" + file.ListItemAllFields.FieldValues["CompliantReport"].ToString() : path;

                        if (path != "")
                        {
                            if (!CheckIfFolderExists(web, "/Document Repository" + path))
                            {
                                CreateFolder(web, "Document Repository", path);
                            }

                            file.MoveTo("https://chrishanna.sharepoint.com/sites/Chaxa/Document%20Repository" + path + "/" + file.Name, MoveOperations.Overwrite);

                            web.Context.ExecuteQuery();
                        }
                    }       
        }

        public static Folder CreateFolder(Web web, string listTitle, string fullFolderUrl)
        {
            if (string.IsNullOrEmpty(fullFolderUrl))
                throw new ArgumentNullException("fullFolderUrl");
            var list = web.Lists.GetByTitle(listTitle);
            return CreateFolderInternal(web, list.RootFolder, fullFolderUrl);
        }

        private static Folder CreateFolderInternal(Web web, Folder parentFolder, string fullFolderUrl)
        {
            var folderUrls = fullFolderUrl.Split(new char[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
            string folderUrl = folderUrls[0];
            var curFolder = parentFolder.Folders.Add(folderUrl);
            web.Context.Load(curFolder);
            web.Context.ExecuteQuery();

            if (folderUrls.Length > 1)
            {
                var subFolderUrl = string.Join("/", folderUrls, 1, folderUrls.Length - 1);
                return CreateFolderInternal(web, curFolder, subFolderUrl);
            }
            return curFolder;
        }

        public static bool CheckIfFolderExists(Web web, string fullFolderUrl)
        {
            try
            {
                if (string.IsNullOrEmpty(fullFolderUrl))
                    throw new ArgumentNullException("fullFolderUrl");

                if (!web.IsPropertyAvailable("ServerRelativeUrl"))
                {
                    web.Context.Load(web, w => w.ServerRelativeUrl);
                    web.Context.ExecuteQuery();
                }

                var folder = web.GetFolderByServerRelativeUrl(web.ServerRelativeUrl + fullFolderUrl);

                web.Context.Load(folder);
                web.Context.ExecuteQuery();
                return true;

            }
            catch (ServerException ex)
            {
                if (ex.ServerErrorTypeName == "System.IO.FileNotFoundException")
                {
                    return false;
                }

                throw;
            }
        }
    }
}

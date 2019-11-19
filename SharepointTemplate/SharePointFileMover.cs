using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core;
using System.Net.Http;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace SharepointTemplate
{

    public static class SharePointFileMover
    {
        [FunctionName("SharePointFileMover")]
        public static async Task RunAsync([HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)]HttpRequestMessage req, TraceWriter log)
        {
            log.Info("SharePoint File Mover Started");
            log.Info(await req.Content.ReadAsStringAsync());

            try
            {
                DropZoneInfo data = await req.Content.ReadAsAsync<DropZoneInfo>();
                
                bool success = RunFileMover(data);
            }
            catch (Exception ex)
            {

                log.Info(ex.Message.ToString());
                log.Info(ex.StackTrace.ToString());
            }

            log.Info("SharePoint File Mover Completed");
        }

        private static bool RunFileMover(DropZoneInfo folderInfo)
        {
            string url = "";
            string appId = ""; //replace it to your app id. 
            string appScret = ""; //replace it to your app secret. 
            string documentLibraryName = folderInfo.listTitle;

            AuthenticationManager manager = new AuthenticationManager();

            using (ClientContext context = manager.GetAppOnlyAuthenticatedContext(url, appId, appScret))
            {
                Web web = context.Web;
                context.Load(web);

                List list = web.Lists.GetByTitle(documentLibraryName);

                CamlQuery query = new CamlQuery();

                ListItemCollection items = list.GetItems(query);
                context.Load(list.Fields);
                context.Load(list.RootFolder);
                context.Load(list.RootFolder.Folders);

                context.ExecuteQuery();

                foreach (Folder f in list.RootFolder.Folders)
                {
                    if (f.Name == folderInfo.currentFolder.Substring(folderInfo.currentFolder.LastIndexOf('/') + 1))
                    {
                        context.Load(f.Files,
                            files => files.Include(file => file.Name),
                            files => files.Include(file => file.ETag),
                            files => files.Include(file => file.TimeLastModified),
                            files => files.Include(file => file.ListItemAllFields["Company"]),
                            files => files.Include(file => file.ListItemAllFields["CompliantReport"]),
                            files => files.Include(file => file.ListItemAllFields["DocumentClass"]),
                            files => files.Include(file => file.ListItemAllFields["DocumentType"]),
                            files => files.Include(file => file.ListItemAllFields["Format"]),
                            files => files.Include(file => file.ListItemAllFields["Language"]),
                            files => files.Include(file => file.ListItemAllFields["Month"]),
                            files => files.Include(file => file.ListItemAllFields["PlanYear"]),
                            files => files.Include(file => file.ListItemAllFields["ProductType"])
                        );

                        context.ExecuteQuery();

                        FileCollection filecol = f.Files;

                        List<FileDetails> pathfilecombo = GeneratePaths(filecol);

                        List<string> uniquePaths = pathfilecombo.Select(a => a.futurePath).Distinct().ToList();

                        CreatePaths(url, appId, appScret, uniquePaths, web);

                        HandleFileMove(pathfilecombo, web);
                    }
                }

                return true;

            }
        }

        public static List<FileDetails> GeneratePaths(FileCollection files)
        {
            List<FileDetails> filedetails = new List<FileDetails>();

            foreach (File file in files)
            {
                string path = "";
                FileDetails fd = new FileDetails();

                path = file.ListItemAllFields.FieldValues["Company"] != null ? path + "/" + file.ListItemAllFields.FieldValues["Company"].ToString() : path;
                path = file.ListItemAllFields.FieldValues["PlanYear"] != null ? path + "/" + file.ListItemAllFields.FieldValues["PlanYear"].ToString() : path;
                path = file.ListItemAllFields.FieldValues["DocumentClass"] != null ? path + "/" + file.ListItemAllFields.FieldValues["DocumentClass"].ToString() : path;
                path = file.ListItemAllFields.FieldValues["DocumentType"] != null ? path + "/" + file.ListItemAllFields.FieldValues["DocumentType"].ToString() : path;
                path = file.ListItemAllFields.FieldValues["Month"] != null ? path + "/" + file.ListItemAllFields.FieldValues["Month"].ToString() : path;
                path = file.ListItemAllFields.FieldValues["ProductType"] != null ? path + "/" + file.ListItemAllFields.FieldValues["ProductType"].ToString() : path;
                path = file.ListItemAllFields.FieldValues["Language"] != null ? path + "/" + file.ListItemAllFields.FieldValues["Language"].ToString() : path;
                path = file.ListItemAllFields.FieldValues["Format"] != null ? path + "/" + file.ListItemAllFields.FieldValues["Format"].ToString() : path;
                path = file.ListItemAllFields.FieldValues["CompliantReport"] != null ? path + "/" + file.ListItemAllFields.FieldValues["CompliantReport"].ToString() : path;

                fd.futurePath = path;
                fd.sharepointfile = file;
                filedetails.Add(fd);
            }

            return filedetails;
        }

        public static void CreatePaths(string url, string appId, string appScret, List<string> paths, Web web)
        {
            foreach (string p in paths)
            {
                if (!CheckIfFolderExists(web, "/Document Repository" + p))
                {
                    CreateFolder(web, "Document Repository", p);
                }
            }
        }

        public static void HandleFileMove(List<FileDetails> fileCol, Web web)
        {
            foreach (FileDetails file in fileCol)
            {
                if (file.futurePath != "")
                {
                    //Make sure to put URL here
                    file.sharepointfile.MoveTo("sites/XXXXXXX/Document%20Repository" + file.futurePath + "/" + file.sharepointfile.Name, MoveOperations.Overwrite);
                    web.Context.ExecuteQueryAsync();
                }
            }
        }

        public static Folder CreateFolder(Web web, string listTitle, string fullFolderUrl)
        {
            if (string.IsNullOrEmpty(fullFolderUrl))
                throw new ArgumentNullException("fullFolderUrl");
            List list = web.Lists.GetByTitle(listTitle);
            return CreateFolderInternal(web, list.RootFolder, fullFolderUrl);
        }

        private static Folder CreateFolderInternal(Web web, Folder parentFolder, string fullFolderUrl)
        {
            string[] folderUrls = fullFolderUrl.Split(new char[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
            string folderUrl = folderUrls[0];
            Folder curFolder = parentFolder.Folders.Add(folderUrl);
            web.Context.Load(curFolder);
            web.Context.ExecuteQuery();

            if (folderUrls.Length > 1)
            {
                string subFolderUrl = string.Join("/", folderUrls, 1, folderUrls.Length - 1);
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

                Folder folder = web.GetFolderByServerRelativeUrl(web.ServerRelativeUrl + fullFolderUrl);

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

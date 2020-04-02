using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Security.Cryptography.X509Certificates;
using Google.Apis.Drive.v3.Data;
using Google.Apis.Http;
using System.Net.Http;
using Newtonsoft.Json;
using System.Windows.Forms;

namespace embedInFile
{
    public static class DriveConnection
    {
        /// <summary>
        /// Does the authentication and creates driveService.
        /// <param name="Scopes">the scopes/permissions that the app has</param>
        /// <param name="credentialsPath">the path to the credentials file (json)</param>
        /// <param name="appName">the application name</param>
        /// </summary>
        public static DriveService startServerConnection(ref string[] Scopes, string credentialsPath, string appName, out ServiceAccountCredential credential)
        {
            //Service account certificate
            X509Certificate2 certificate =
                    new X509Certificate2(credentialsPath, "notasecret", X509KeyStorageFlags.Exportable);
            //Get credentials
            credential = new ServiceAccountCredential(
                    new ServiceAccountCredential.Initializer("embedinfile@embedinfile.iam.gserviceaccount.com")
                    {
                        Scopes = Scopes
                    }.FromCertificate(certificate)
                );
            return createDriveService(credential, appName);
        }

        public static DriveService startClientConnection(ref string[] Scopes, string credentialsPath, string appName)
        {

            UserCredential credential;

            using (var stream =
                new FileStream(credentialsPath, FileMode.Open, FileAccess.Read))
            {
                // The file token.json stores the user's access and refresh tokens, and is created
                // automatically when the authorization flow completes for the first time.
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
            }
            return createDriveService(credential, appName);
        }

        private static DriveService createDriveService(
            Google.Apis.Http.IConfigurableHttpClientInitializer credential, string appName)
        {
            // Create Drive API service.
            DriveService driveService = new DriveService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = appName,
            });
            return driveService;
        }

        /// <summary>
        /// Gets the id of a file/folder by its name.
        /// </summary>
        /// <param name="id"></param>
        /// <returns>the id of the first occurrence of that file name, 
        /// or "" if no file with that name was found</returns>
        public static string findIDByName(DriveService driveService, string name)
        {
            FilesResource.ListRequest listRequest = driveService.Files.List();
            listRequest.Q = $"mimeType = 'application/vnd.google-apps.folder' and name = '{name}'";
            listRequest.Fields = "nextPageToken, files(id, name)";
            //We want to get all of them
            IList<Google.Apis.Drive.v3.Data.File> files = listRequest.Execute().Files;
            if (files != null && files.Count > 0)
            {
                return files[0].Id;
            }
            else
            {
                Globals.ThisAddIn.sayWord($"No file with name {name} found");
                return "";
            }
        }

        /// <summary>
        /// Creates a folder using the service, inside the folder specified by parentID
        /// If the desired parent is root, then parentID should be ""
        /// </summary>
        /// <param name="folderName">the name that will be given to the folder</param>
        /// <param name="service">driveService used</param>
        /// <param name="parentID">the id of the parent folder, or "" if the parent is root</param>
        /// <returns>the ID of the folder that has been created</returns>
        public static string CreateFolder(DriveService service, string folderName, string parentID)
        {
            var fileMetadata = new Google.Apis.Drive.v3.Data.File()
            {
                Name = folderName,
                MimeType = "application/vnd.google-apps.folder",
            };
            //If the parent folder is different from the root one, specify parent
            if (parentID != "")
            {
                fileMetadata.Parents = new List<string>() { parentID };
            }
            var request = service.Files.Create(fileMetadata);
            request.Fields = "id";
            var file = request.Execute();
            return file.Id;
        }

        public static void listFiles(DriveService service)
        {
            FilesResource.ListRequest listRequest = service.Files.List();
            listRequest.PageSize = 10;
            listRequest.Fields = "nextPageToken, files(id, name)";
            IList<Google.Apis.Drive.v3.Data.File> files = listRequest.Execute()
                .Files;
            // List files.
            Globals.ThisAddIn.sayWord("Files:\r\n");
            if (files != null && files.Count > 0)
            {
                foreach (var file in files)
                {
                    Globals.ThisAddIn.sayWord($"{file.Name} ({file.Id})\r\n");
                }
            }
            else
            {
                Globals.ThisAddIn.sayWord("No files found.");
            }
        }

        //Structs to hold the valid values for permission types and roles
        public struct PermissionTypes
        {
            public const string USER = "user";
            public const string GROUP = "group";
            public const string DOMAIN = "domain";
            public const string ANYONE = "anyone";
        }

        public struct PermissionRoles
        {
            public const string OWNER = "owner";
            public const string ORGANIZER = "organizer";
            public const string FILE_ORGANIZER = "fileOrganizer";
            public const string WRITER = "writer";
            public const string COMMENTER = "commenter";
            public const string READER = "reader";
        }
        public static void changePermissions(DriveService service, string fileID, string type, string role)
        {
            changePermissions(service, fileID, type, role, "");
        }

        public static void changePermissions(
            DriveService service, string fileID, string type, string role, string emailAddress)
        {
            //Create permission, with email adress if necessary
            Permission p = new Permission();
            p.Type = type;
            p.Role = role;
            if (type == PermissionTypes.DOMAIN || type == PermissionTypes.ANYONE)
            {
                p.AllowFileDiscovery = true;
            }
            if (emailAddress != null && emailAddress != "")
            {
                p.EmailAddress = emailAddress;
            }
            try
            {
                service.Permissions.Create(p, fileID).Execute();
            } catch (Exception e)
            {
                Globals.ThisAddIn.sayWord("Exception creating permissions: " + e.ToString());
            }
            
        }
    
        public static string setSharedURL(DriveService service, string fileID)
        {
            changePermissions(service, fileID, PermissionTypes.ANYONE, PermissionRoles.READER);
            return "http://drive.google.com/file/d/" + fileID;
        }

        /// <summary>
        /// Returns a FileStream of the file to upload
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public static FileStream getFile(string path)
        {
            return new System.IO.FileStream(path, System.IO.FileMode.Open, System.IO.FileAccess.Read);
        }

        public static void deleteFile(DriveService service, string id)
        {
            try
            {
                var newid = id.Trim();
                service.Files.Delete(newid).Execute();
                MessageBox.Show("Deletion successful", "Deletion of file", MessageBoxButtons.OK);
            }
            catch (Exception e)
            {
                MessageBox.Show("Error while deleting: " + e.ToString(), "Error", MessageBoxButtons.OK);
            }
        }

        //Show permissions
        /*
            try
            {
                changePermissions(service, id, PermissionTypes.USER, 
                    PermissionRoles.FILE_ORGANIZER, "embedinfile@embedinfile.iam.gserviceaccount.com");
                
                PermissionsResource.ListRequest listRequest = service.Permissions.List(id.Trim());
                listRequest.Fields = "nextPageToken, permissions(id, type, role)";
                IList<Google.Apis.Drive.v3.Data.Permission> perms = listRequest.Execute().Permissions;
                // List files.
                Globals.ThisAddIn.sayWord("Files:\r\n");
                if (perms != null && perms.Count > 0)
                {
                    foreach (var perm in perms)
                    {
                        Globals.ThisAddIn.sayWord($"Id: {perm.Id}, type: {perm.Type}, " +
                            $"role: {perm.Role}\r\n");
                    }
                }
                else
                {
                    Globals.ThisAddIn.sayWord("No permissions found.");
                }
                MessageBox.Show("Permissions shown");
            }
            catch (Exception e)
            {
                Globals.ThisAddIn.sayWord(e.ToString() + "\r\n");
            }*/
    }
}

using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Http;
using Google.Apis.Services;
using Google.Apis.Upload;
using Google.Apis.Util.Store;
using Microsoft.Office.Tools.Word;
using Word = Microsoft.Office.Interop.Word;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace embedInFile
{
    class DriveEmbedding : IDriveConnection
    {
        private static string[] Scopes = { DriveService.Scope.DriveFile };
        private static string ApplicationName = "embedInFile";
        private const string ROOTNAME = "embedInFile";


        private bool started = false;
        private DriveService driveService;
        private string rootID;
        private ServiceAccountCredential credential;

        

        public DriveEmbedding()
        {

        }
        public void removeAll()
        {
            throw new NotImplementedException();
        }

        public bool removeLink(string path)
        {
            try
            {
                start();
                return DriveConnection.deleteFile(driveService, path);
            } catch (Exception e)
            {
                MessageBox.Show("An error happened: " + e.ToString(), "Error", MessageBoxButtons.OK);
                return false;
            }
        }
        
        public async Task uploadLink(Word.Range cc, string path, string documentName)
        {
            try
            {
                start();

                FileStream inputfile = DriveConnection.getFile(path); //File to upload
                await uploadLargeFileAsync(cc, inputfile, path, documentName);
            } catch (Exception e)
            {
                MessageBox.Show("An error happened: " + e.ToString(), "Error", MessageBoxButtons.OK);
            }
        }

        private void start()
        {
            if (!started)
            {
                //Will set this.driveService and authenticate
                driveService = DriveConnection.startServerConnection(ref Scopes,
                    "embedinfile-db8ffaf7c414.p12",
                    "embedInFile", out credential);

                //If root folder has not been created for this user, create it
                createRootFolder();
                started = true;

            }
        }

        /// <summary>
        /// Creates root folder for the app embedInFile if not present
        /// </summary>
        private void createRootFolder()
        {
            //First check if the root folder is present. If it is, retrieve its ID
            string id = DriveConnection.findIDByName(driveService, ROOTNAME);
            if (id == "")
            {
                rootID = DriveConnection.CreateFolder(driveService, ROOTNAME, ""); 
                //Set permissions to the folder
                DriveConnection.changePermissions(driveService, rootID,
                    DriveConnection.PermissionTypes.ANYONE, DriveConnection.PermissionRoles.READER);
            } else
            {
                rootID = id;
            }
        }

        private async Task uploadLargeFileAsync(Word.Range cc, FileStream stream, string fileName, string docName)
        {
            string name = docName + "_" + Path.GetFileName(fileName);
            string extension = Path.GetExtension(fileName);

            var fileMetadata = new Google.Apis.Drive.v3.Data.File()
            {
                Name = name
            };
            FilesResource.CreateMediaUpload request;
            request = driveService.Files.Create(
                fileMetadata, stream, "audio/mpeg");
            request.Fields = "id";
            request.ProgressChanged += (obj) => { string t = $"Uploading... {obj.BytesSent / stream.Length}%";
                cc.Text = t;
                //cc.PlaceholderText = t;
            };
            var result = await request.UploadAsync();
            UploadStatus finalResult = await handleUpload(request, result.Status);
            //var finalResult = result.Status;
            if (finalResult == UploadStatus.Completed)
            {
                var file = request.ResponseBody;
                object URL = DriveConnection.setSharedURL(driveService, file.Id);
                //cc.Text = (string) URL;
                cc.Hyperlinks.Add(cc, ref URL, System.Type.Missing, Path.GetFileName(name), Path.GetFileName(name));
                //cc.PlaceholderText = (string) URL;
                //cc.LockContents = true;
            } else if (finalResult == UploadStatus.Failed)
            {
                cc.Text = "Upload failed :(";
                //cc.PlaceholderText = "Upload failed :(";
                MessageBox.Show("Response: " + "\r\n", "Error", MessageBoxButtons.OK);
                //cc.LockContents = true;
            } else
            {
                Globals.ThisAddIn.sayWord("Recursion not happening");
            }
            //return setSharedURL(service, "ID DEL ARCHIVO SUBIDO");
        }

        private async Task<UploadStatus> handleUpload(FilesResource.CreateMediaUpload request, UploadStatus status)
        {
            switch (status)
            {
                case UploadStatus.Completed:
                    return UploadStatus.Completed;
                case UploadStatus.Failed:
                    return UploadStatus.Failed;
                default:
                    status = (await request.ResumeAsync()).Status;
                    return await handleUpload(request, status);
            }
        }

        public void listFiles()
        {
            try
            {
                start();
                DriveConnection.listFiles(driveService);
            } catch (Exception e)
            {
                MessageBox.Show("An error happened: " + e.ToString(), "Error", MessageBoxButtons.OK);
            }
        }
    }
}

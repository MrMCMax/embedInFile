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

        private string deployLocation;

        /// <summary>
        /// Constructor for the driveEmbedding object. Needs the path where credentials are located on deployment.
        /// </summary>
        /// <param name="deployLocation"></param>
        public DriveEmbedding(string deployLocation)
        {
            this.deployLocation = deployLocation;
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
        /// <summary>
        /// Uploads a file to drive and creates a link to it, displayed in cc.
        /// </summary>
        /// <param name="cc">The word content control to write the link to</param>
        /// <param name="absolutePath">The absolute path to the file to upload</param>
        /// <param name="name">The name of the file</param>
        /// <returns>true if the operation was successful, false otherwise</returns>
        public async Task<String> uploadLink(Word.Range cc, string absolutePath, string name)
        {
            try
            {
                start();

                FileStream inputfile = DriveConnection.getFile(absolutePath); //File to upload
                string result = await uploadLargeFileAsync(cc, inputfile, absolutePath, name);
                return result;
            } catch (Exception e)
            {
                MessageBox.Show("An error happened: " + e.ToString(), "Error", MessageBoxButtons.OK);
                return "";
            }
        }

        private void start()
        {
            if (!started)
            {
                //Will set this.driveService and authenticate
                driveService = DriveConnection.startServerConnection(ref Scopes,
                    deployLocation + "\\" + "embedinfile-db8ffaf7c414.p12",
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

        /// <summary>
        /// uploads a potentially large file to google drive and stores the link in cc.
        /// </summary>
        /// <param name="cc"></param>
        /// <param name="stream"></param>
        /// <param name="fileName"></param>
        /// <param name="docName"></param>
        /// <returns></returns>
        private async Task<String> uploadLargeFileAsync(Word.Range cc, FileStream stream,
            string absolutePath, string name)
        {
            //Create drive File object with the file name
            var fileMetadata = new Google.Apis.Drive.v3.Data.File()
            {
                Name = name
            };
            //Create the request for uploading a large audio file
            FilesResource.CreateMediaUpload request;
            request = driveService.Files.Create(
                fileMetadata, stream, "audio/mpeg");
            request.Fields = "id";
            //Set the progress change method
            request.ProgressChanged += (obj) => { string t = $"Uploading... {obj.BytesSent / stream.Length}%";
                cc.Text = t;
            };
            string URL = "";
            //Try to upload. Get result
            var result = await request.UploadAsync();
            //If the async method returned without completing the request, handleUpload will take care of it
            //and return the final result
            UploadStatus finalResult = await handleUpload(request, result.Status);
            if (finalResult == UploadStatus.Completed)
            {
                //The file was successfully uploaded. Retrieve the file
                var file = request.ResponseBody;
                //Set its permissions (anyone can read with URL) and retrieve its URL
                URL = DriveConnection.setSharedURL(driveService, file.Id);
                object link = URL;
                //create hyperlink in the content control
                cc.Hyperlinks.Add(cc, ref link, System.Type.Missing, name, name);
            }
            else if (finalResult == UploadStatus.Failed)
            {
                cc.Text = "Upload failed :(";
                MessageBox.Show("Response: " + "\r\n", "Error", MessageBoxButtons.OK);
            }
            else
            {
                Globals.ThisAddIn.sayWord("Recursion not happening");
            }
            return URL;
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

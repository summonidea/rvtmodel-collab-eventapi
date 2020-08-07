using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Http;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI.Events;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB.ExternalService;
using System.Windows.Forms;
using System.Threading;
using System.Diagnostics;
using MongoDB.Driver;
using MongoDB.Bson;

namespace cfarevit
{
    public class App : IExternalApplication
    {
        const string _test_project_filepath = "D:/Project1.rvt";
        BsonDocument current;
        public MongoClient client;
        public IMongoCollection<BsonDocument> collection;

        public Result OnShutdown(UIControlledApplication application)
        {            
            return Result.Succeeded;
        }

        void OnApplicationInitialized(object sender, ApplicationInitializedEventArgs e)
        {
            try
            {
                client = new MongoClient();
                var database = client.GetDatabase("my_database");
                collection = database.GetCollection<BsonDocument>("collabrequests");

                // Sender is an Application instance:
                Autodesk.Revit.ApplicationServices.Application app = sender as Autodesk.Revit.ApplicationServices.Application;
                UIApplication uiapp = new UIApplication(app);

                current = searchQueue();
                needsCollaboration(sender);

            }
            catch (Exception err)
            {
                Console.Write(err.ToString());
            }
            
        }

        void OnDocumentOpened(object sender, DocumentOpenedEventArgs e)
        {
            try
            {
                
                Document doc = e.Document;
                //string title = doc.Title;
                doc.SaveAsCloudModel(current["folder_id"].AsString, current["file_name"].AsString);
                Console.Write("cualquier cosa");
                if (doc.CanEnableCloudWorksharing())
                {
                    doc.EnableCloudWorksharing();

                    TransactWithCentralOptions transact = new TransactWithCentralOptions();
                    SynchronizeWithCentralOptions synch = new SynchronizeWithCentralOptions();
                    synch.Comment = "Autosaved by the API at " + DateTime.Now;
                    RelinquishOptions relinquishOptions = new RelinquishOptions(true);
                    relinquishOptions.CheckedOutElements = true;
                    synch.SetRelinquishOptions(relinquishOptions);
                                   
                    doc.SynchronizeWithCentral(transact, synch);
                }
            }
            catch (Exception err)
            {
                Console.Write(err.ToString());
            }
        }

        void OnDocumentSaving(object sender, DocumentSavingEventArgs e)
        {
            try
            {
                Console.WriteLine("savingDoc");
            }
            catch (Exception err)
            {
                Console.Write(err.ToString());
            }
        }

        void needsCollaboration(Object sender)
        {
            if (current != null)
            {
                Autodesk.Revit.ApplicationServices.Application app = sender as Autodesk.Revit.ApplicationServices.Application;
                UIApplication uiapp = new UIApplication(app);

                var filter = Builders<BsonDocument>.Filter.Eq("_id", current["_id"]);
                var update = Builders<BsonDocument>.Update.Set("is_running", true);
                collection.UpdateOne(filter, update);

                uiapp.OpenAndActivateDocument(_test_project_filepath);

                RevitCommandId closeDoc = RevitCommandId.LookupPostableCommandId(PostableCommand.Close);
                uiapp.PostCommand(closeDoc);
            }
            else
            {
                closeRevit();
            }
        }

        void OnDocumentClosed(object sender, DocumentClosedEventArgs e)
        {
            try
            {
                var filter = Builders<BsonDocument>.Filter.Eq("_id", current["_id"]);
                var update = Builders<BsonDocument>.Update.Set("is_collaborated", true).Set("is_running", false);
                collection.UpdateOne(filter, update);
        
                current = searchQueue();
                needsCollaboration(sender);
            }
            catch (Exception err)
            {
                Console.Write(err.ToString());
            }
        }

        void OnSynchronizedWithCentral(object sender, DocumentSynchronizedWithCentralEventArgs e)
        {
            try
            {
                Console.WriteLine("Synchronized with central");
            }
            catch (Exception err)
            {
                Console.Write(err.ToString());
            }
        }

        private BsonDocument searchQueue()
        {
            var filter = Builders<BsonDocument>.Filter.Eq("is_collaborated", false);
            var document = collection.Find(filter).FirstOrDefault();
            return document;
        }

        private void closeRevit()
        {
            Process[] procs = null;
            try
            {
                procs = Process.GetProcessesByName("revit");

                Process mspaintproc = procs[0];

                if (!mspaintproc.HasExited)
                {
                    mspaintproc.Kill();
                }
            }
            finally
            {
                if (procs != null)
                {
                    foreach (Process p in procs)
                    {
                        p.Dispose();
                    }
                }
            }
        }

        public Result OnStartup(UIControlledApplication application)
        {
            application.ControlledApplication.ApplicationInitialized += OnApplicationInitialized;
            application.ControlledApplication.DocumentOpened += OnDocumentOpened;
            application.ControlledApplication.DocumentSynchronizedWithCentral += OnSynchronizedWithCentral;
            application.ControlledApplication.DocumentClosed += OnDocumentClosed;
            return Result.Succeeded;
        }
    }

    //[Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    //public class ECommand : IExternalCommand
    //{
    //    static string str;
    //    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    //    {
    //        Console.WriteLine("dfsdf");
            
    //        //UIApplication uiapp = commandData.Application;
    //        //str = uiapp.;
           
    //        TaskDialog.Show("Revit", "adfadfa");
    //        IExternalEventHandler myHandler = new EEvent();
    //        ExternalEvent myEvent = ExternalEvent.Create(myHandler);

    //        return Autodesk.Revit.UI.Result.Succeeded;
    //    }
        
    //};

    //    public class EEvent : IExternalEventHandler
    //    {
    //        public void Execute(UIApplication uiapp)
    //        {

    //            string Path = "D:/Project1.rvt";
    //            Document doc = uiapp.ActiveUIDocument.Document;

    //            UIDocument uidoc = new UIDocument(doc);
    //            uiapp.OpenAndActivateDocument(Path);

    //            //throw new NotImplementedException();
    //        }

    //        public string GetName()
    //        {
    //            throw new NotImplementedException();
    //        }
    //    }
}

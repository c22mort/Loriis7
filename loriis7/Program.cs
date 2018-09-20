//=============================================================================================
//=     LOg Rotator for IIS7
//=============================================================================================
//=
//=     ©A.Patrick 2014
//=     
//=============================================================================================
using Ionic.Zip;
using Microsoft.Web.Administration;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace loriis7
{
    public class Program
    {
        public static string configFileName = "loriis7.config.xml";
        // List of Log Folders
        private static List<string> logDirList = new List<string>();


        // Number of Log Files Zipped, Deleted and Archvied
        private static Int32 logsZipped = 0;
        private static Int32 logsDeleted = 0;
        private static Int32 logsArchived = 0;

        // Size of Files Deleted
        private static long sizeDeleted = 0;
        // Space Saved by Zipping
        private static long zipSaved = 0;
        // Size of Files Archvied
        private static long sizeArchived = 0;
        // Config Info
        public static LoriisConfig config;

        static void Main(string[] args)
        {
            // Add Handler For Embedded Resource
            AppDomain.CurrentDomain.AssemblyResolve += new ResolveEventHandler(CurrentDomain_AssemblyResolve);

            // Get Version 
            Version version = Assembly.GetEntryAssembly().GetName().Version;

            // Show Splash
            Console.WriteLine("LOg Rotator IIS7, " + version.Major + "." + version.Minor + "." + version.Build + ", © A.Patrick 2013");
            Console.WriteLine("");

            // Read Our Config File or create with defaults if first run
            readConfig();
            
            // Get Log Dirs
            getLogDirs();

            Parallel.For(0, logDirList.Count, i =>
            {
                // Check if Directory Exists First
                if (Directory.Exists(logDirList[i]))
                {
                    string[] logFileList = Directory.GetFiles(logDirList[i], "*", SearchOption.AllDirectories);
                    foreach (String sLogFile in logFileList)
                    {
                        if (Path.GetExtension(sLogFile).ToLower() == ".zip")
                        {
                            processZipFile(sLogFile);
                        }
                        else
                        {
                            processLogFile(sLogFile);
                        }
                    }
                }
                else { Console.WriteLine(logDirList[i] + " NOT FOUND!"); }
            });

            zipSaved = zipSaved / 1024;
            Console.WriteLine("Logs Files Zipped   : " + logsZipped.ToString() + ", Space Saved -> " + zipSaved.ToString() + "Kb");
            sizeDeleted = sizeDeleted / 1024;
            Console.WriteLine("Logs Files Deleted  : " + logsDeleted.ToString() + ", Space Saved -> " + sizeDeleted.ToString() + "Kb");
            sizeArchived = sizeArchived / 1024;
            Console.WriteLine("Logs Files Archived : " + logsArchived.ToString() + ", Size        -> " + sizeArchived.ToString() + "Kb");
        }

        //=====================================================================================
        //=     get Log Dirs
        //=====================================================================================
        //=     Caclulate Log file directories from site objects
        //=====================================================================================
        static void getLogDirs() 
        {
            // Create a Server Manager
            ServerManager sm = new ServerManager();

            // Finding Log File Directries
            Console.WriteLine("Searching For Log File Directores...");

            try
            {
                SiteCollection sites = sm.Sites;
                // Loop Through Each Site
                foreach (Site s in sm.Sites)
                {
                    // Calculate Log File Directory
                    string logDir = Path.Combine(Environment.ExpandEnvironmentVariables(s.LogFile.Directory), "W3SVC" + s.Id);
                    if (!logDirList.Contains(logDir))
                    {
                        // Add to List if it doesn't already
                        logDirList.Add(logDir);
                        Console.WriteLine("    " + logDir);
                    }
                }
            }
            catch (Exception)
            {
                Console.WriteLine("Couldn't get Site list...are you sure IIS is installed!");
            }

            // Dispose of Server Manager
            sm.Dispose();
        }


        //=====================================================================================
        //=     Process Log File
        //=====================================================================================
        //=     Checks Last Write Time of log File and zips it if older than X Days
        //=====================================================================================
        static void processLogFile(string logFileName)
        {
            
            // Get FileInfo Sctructure of Log File
            FileInfo lfi = new FileInfo(logFileName);

            // Just Return if it's less than Specified Days Old
            if (lfi.LastWriteTime.AddDays(config.DaysBeforeZip) > DateTime.Now) { return; }

            // Otherwise Zip it Up
            string zipFileName = Path.Combine(Path.GetDirectoryName(logFileName),Path.GetFileNameWithoutExtension(logFileName) + ".zip");
            
            // Delete Zip File if it already exists
            if (File.Exists(zipFileName)) { File.Delete(zipFileName); }

            try
            {
                // Now we can proceed with Creating Zip File
                using (ZipFile zip = new ZipFile())
                {
                    // Set Compression
                    zip.CompressionLevel = Ionic.Zlib.CompressionLevel.BestCompression;
                    // Add Our Log File to the zipFile
                    zip.AddFile(logFileName);
                    // Save Zip File 
                    zip.Save(zipFileName);
                    logsZipped += 1;

                    // Get File info for Zip File
                    FileInfo zfi = new FileInfo(zipFileName);
                    zipSaved = zipSaved + (lfi.Length - zfi.Length);

                    // See if we have an Archive Path
                    if (config.ArchivePath != "")
                    {
                        //Does it Exist
                        if (Directory.Exists(config.ArchivePath))
                        {
                            // Now we Can Copy The Zip File to The Archive Path
                            try
                            {
                                // Work out Destination Folder
                                string sDestPath = Path.Combine(config.ArchivePath, System.Environment.MachineName);
                                // Create it if it doesn't exist
                                if (!Directory.Exists(sDestPath)) { Directory.CreateDirectory(sDestPath); }
                                // And Copy File
                                File.Copy(zipFileName, Path.Combine(sDestPath, Path.GetFileName(zipFileName)));
                                logsArchived += 1;
                                sizeArchived += zfi.Length;
                            }
                            catch (Exception)
                            {
                                Console.WriteLine("File not Copied : " + zipFileName);
                            }
                        }
                    }
                }

                // And Then Delete it
                File.Delete(logFileName);

            }
            catch (Exception) { }

        }


        //=====================================================================================
        //=     Process Zip File
        //=====================================================================================
        //=     Checks Last Write Time of zip File and Deletes it if older than X Days
        //=====================================================================================
        static void processZipFile(string zipFileName)
        {
            // Get FileInfo Sctructure of Log File
            FileInfo fi = new FileInfo(zipFileName);
            // Just Return if it's less than Specified Days Old
            if (fi.LastWriteTime.AddDays(config.DaysBeforeDeletion) > DateTime.Now) { return; }

            // Otherwise Delete it
            try
            {
                File.Delete(zipFileName);
                logsDeleted += 1;
                sizeDeleted += fi.Length;
            }
            catch (Exception) {}
        }


        //=====================================================================================
        //=     Read Config
        //=====================================================================================
        //=     Reads XML Config File, creates new one with default settings if
        //=     it doesn't exist or there are problems.
        //=====================================================================================
        public static void readConfig()
        {
            // Create a new Serializer
            XmlSerializer serializer = new XmlSerializer(typeof(LoriisConfig));

            try 
        	{
                // Open File
                FileStream fs = new FileStream(configFileName, FileMode.Open);

                // Serialize Config
                config = (LoriisConfig)serializer.Deserialize(fs);
                
                // Close File Stream
                fs.Close();
        	}
	        catch (Exception)
	        {
                // Tell User there was a problem
                Console.WriteLine("There was a problem with the config file, lorriss7.config.xml");
                Console.WriteLine("Creating a new one using defaults!");
                Console.WriteLine("");
                
                // There has been a problem so Create a new Config
                config = new LoriisConfig();
                
                // Set Default Properties
                config.DaysBeforeZip = 7;
                config.DaysBeforeDeletion = 180;
                config.ArchivePath = "";

                // Save Config File
                TextWriter writer = new StreamWriter(configFileName);
                serializer.Serialize(writer, config);
            }
            // Show Config
            Console.WriteLine("Days Before Zipping  = " + config.DaysBeforeZip);
            Console.WriteLine("Days Before Deletion = " + config.DaysBeforeDeletion);
            Console.WriteLine("Archive Destination  = " + config.ArchivePath);
            Console.WriteLine("");
        }

        // Configuration Class
        // Read/Write Using XML Serialiser
        public class LoriisConfig
        {
            public Int16 DaysBeforeZip;
            public Int16 DaysBeforeDeletion;
            public string ArchivePath;
        }


        // Used For Embedded Resources
        static System.Reflection.Assembly CurrentDomain_AssemblyResolve(object sender, ResolveEventArgs args)
        {

            string dllName = args.Name.Contains(",") ? args.Name.Substring(0, args.Name.IndexOf(',')) : args.Name.Replace(".dll", "");

            dllName = dllName.Replace(".", "_");

            if (dllName.EndsWith("_resources")) return null;

            System.Resources.ResourceManager rm = new System.Resources.ResourceManager(typeof(Program).Namespace + ".Properties.Resources", System.Reflection.Assembly.GetExecutingAssembly());

            byte[] bytes = (byte[])rm.GetObject(dllName);

            return System.Reflection.Assembly.Load(bytes);

        }
    
    }
}

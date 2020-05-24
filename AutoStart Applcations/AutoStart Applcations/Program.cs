using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.IO;
using System.Diagnostics;
using System.Threading;
using System.Windows.Forms;
using F = IWshRuntimeLibrary ;
using System.Reflection;

namespace AutoStart_Applcations
{
    class Program
    {
        #region Variables
        /// <summary>
        /// Variable of a Thread
        /// </summary>
        public static Thread Background;

        //0 in the array for each string is the path and 1 is the application name
        /// <summary>
        /// Contains Applications absolute path inside the string array
        /// </summary>
        public static List<string[]> ApplicationPaths = new List<string[]>();

        public static string StartMenu = $@"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp";

        /// <summary>
        /// Tells latest thread what app is loaded as of the moment it started
        /// </summary>
        public static int CurrentlyLoadedApp = 0;
        #endregion

        #region Main Important Methods...
        /// <summary>
        /// Entry Point of the applcation
        /// </summary>
        /// <param name="args"></param>
        static void Main(string[] args)
        {
            //Startup Function
            Startup();

            //Message that we will ask the user apon launch of this application
            var msg = MessageBox.Show("Do you wish to have your applcations auto start?", "Applcation Auto Start Process!", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (msg == DialogResult.Yes)
            {
                //Gets Applcation Names and Path Files from the Location Text File
                GetItems();

                //Loops each applcation and checks if they are started if not starts them automatically
                foreach (var app in ApplicationPaths)
                {
                    //Creates an Instance of the method [Main Thread] and every time it cycles to a new applcation.
                    //it starts the function in another thread then checks if the app is started or not
                    Background = new Thread(MainThread);

                    //Starts current instance of background
                    Background.Start();

                    //Sleeps for two seconds before starting the next instance 
                    Thread.Sleep(2000);

                    //Increments the value for the currently loaded app so it will be no longer 0 so it can auto match its path and name for the app
                    CurrentlyLoadedApp++;
                }
            }
            else if (msg == DialogResult.No) Environment.Exit(0);
        }

        /// <summary>
        /// Place at which every applcation is handled
        /// </summary>
        static void MainThread()
        {
            //Sets the Loaded App as _LoadedApp
            int _LoadedApp = CurrentlyLoadedApp;

            //Updates that variable into another to prevent the public int from affecting the currently loaded application for this thread
            //the currently loaded app is now the variable of [LoadedApp]
            int LoadedApp = _LoadedApp;
            
            //Starts the Program since the program is not open yet
            if(ApplicationPaths[LoadedApp][0] != "" && CheckRunningApp(ApplicationPaths[LoadedApp][1]) == false){
                try
                {
                    ProcessStartInfo PS = new ProcessStartInfo(ApplicationPaths[LoadedApp][0]);
                    Console.WriteLine("Program is [NOT] Open!");
                    Process.Start(PS);
                }
                catch(Exception Ex){ }
            }
            //The Program is open so its gonna do something later ig
            else if (ApplicationPaths[LoadedApp][0] != "" && CheckRunningApp(ApplicationPaths[LoadedApp][1]) == true){
                Console.WriteLine("Program is Open!");
            }
        }

        /// <summary>
        /// Populates the items at the top with an items name and location
        /// </summary>
        static void GetItems()
        {
            string Location = $@"{Environment.CurrentDirectory}\ItemLocation.txt";

            //Loads Locations from this file
            if (File.Exists(Location))
            {
                try
                {
                    #region Stream Reader...
                    string[] t;
                    string test = "";
                    //This is what is used to read the text file for the path files 
                    using (StreamReader sr = new StreamReader(Location))
                    {
                        while (sr.EndOfStream == false)
                        {
                            test += sr.ReadLine();
                        }
                        t = test.Split(',');
                        sr.Close();
                        //MessageBox.Show($"Total Items: {t.Count().ToString()}");
                    }
                    #endregion

                    #region Item Sorting...
                    int Count = 0;
                    foreach (var i in t)
                    {
                        if (i.Contains(@"\") == true && File.Exists(i))
                        {
                            string name = i.Remove(0, i.LastIndexOf(@"\"));
                            name = name.Replace(@"\", "");
                            name = name.Replace(".exe", "");
                            string[] item = { i, name };
                            ApplicationPaths.Add(item);
                            //MessageBox.Show($"Name: {ApplicationPaths[Count][1]}\nPath: {ApplicationPaths[Count][0]}");
                        }
                        else MessageBox.Show($"The Path {i} does not exist!");
                        Count++;
                    }
                    #endregion
                }
                catch (Exception Ex) { }
            }
            else MessageBox.Show("Missing file with item locations!", "ItemLocation.txt Missing!", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        /// <summary>
        /// Asks the user if they want the program to be saved in the startup menu
        /// </summary>
        static void Startup()
        {
            //This means there is [NOT] a shortcut of this application inside of the start menu
            if(CheckLocation() == false)
            {
                var MSG = MessageBox.Show("Would you like to auto set up this program to lauch as soon as you login?", "Auto Setup Question Dialog", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (MSG == DialogResult.Yes) CreateShortcut("AutoStart", $@"{StartMenu}", Assembly.GetExecutingAssembly().Location);
            }
        }
        #endregion

        //Extra stuff either cleaned up functions or one for a one time use method
        #region Extra...
        /// <summary>
        /// Checks if the program is running and returns a true or false depending on the result
        /// </summary>
        /// <param name="Applcation">name of the application to check</param>
        /// <returns></returns>
        static bool CheckRunningApp(string Applcation)
        {
            if (Process.GetProcessesByName(Applcation).Length > 0) return true;
            else return false;
        }

        /// <summary>
        /// Handy Tool To show process names
        /// </summary>
        static void ShowProcesses()
        {
            foreach (var p in Process.GetProcesses())
            {
                Console.WriteLine(p.ProcessName);
            }
        }
        
        /// <summary>
        /// Returns true or false based on if the application startup is located in the start menu folder
        /// </summary>
        /// <returns></returns>
        static bool CheckLocation()
        {
            if (File.Exists($@"{StartMenu}\AutoStart.lnk")) return true;
            else return false;
        }


        static void CreateShortcut(string shortcutName, string shortcutPath, string targetFileLocation)
        {
            string shortcutLocation = Path.Combine(shortcutPath, shortcutName + ".lnk");
            F.WshShell shell = new F.WshShell();
            F.IWshShortcut shortcut = (F.IWshShortcut)shell.CreateShortcut(shortcutLocation);

            shortcut.Description = "Auto Launches Programs for us";
            shortcut.TargetPath = targetFileLocation;
            shortcut.Save();
        }
        #endregion
    }
}

using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.ComponentModel.Design;
using System.Xml;
using EnvDTE;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Shell;
using Microsoft.Win32;

namespace Community.VSBrowserSwitcher
{
	/// <summary>
	/// This is the class that implements the package exposed by this assembly.
	///
	/// The minimum requirement for a class to be considered a valid package for Visual Studio
	/// is to implement the IVsPackage interface and register itself with the shell.
	/// This package uses the helper classes defined inside the Managed Package Framework (MPF)
	/// to do it: it derives from the Package class that provides the implementation of the 
	/// IVsPackage interface and uses the registration attributes defined in the framework to 
	/// register itself and its components with the shell.
	/// </summary>
	// This attribute tells the PkgDef creation utility (CreatePkgDef.exe) that this class is
	// a package.
	[PackageRegistration(UseManagedResourcesOnly = true)]
	// This attribute is used to register the informations needed to show the this package
	// in the Help/About dialog of Visual Studio.
	[InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]
	// This attribute is needed to let the shell know that this package exposes some menus.
	[ProvideMenuResource("Menus.ctmenu", 1)]
	[Guid(GuidList.guidCommunity_VSBrowserSwitcherPkgString)]
	[ProvideAutoLoad(UIContextGuids.SolutionExists)]
	[ProvideAutoLoad(UIContextGuids.NoSolution)]
	[ProvideAutoLoad(UIContextGuids.EmptySolution)]
	public sealed class Community_VSBrowserSwitcherPackage : Package
	{
		private FileSystemWatcher _watcher = new FileSystemWatcher();

		/// <summary>
		/// Default constructor of the package.
		/// Inside this method you can place any initialization code that does not require 
		/// any Visual Studio service because at this point the package object is created but 
		/// not sited yet inside Visual Studio environment. The place to do all the other 
		/// initialization is the Initialize method.
		/// </summary>
		public Community_VSBrowserSwitcherPackage()
		{
			//Trace.WriteLine(string.Format(CultureInfo.CurrentCulture, "Entering constructor for: {0}", this.ToString()));
		}

		/////////////////////////////////////////////////////////////////////////////
		// Overriden Package Implementation
		#region Package Members

		/// <summary>
		/// Initialization of the package; this method is called right after the package is sited, so this is the place
		/// where you can put all the initilaization code that rely on services provided by VisualStudio.
		/// </summary>
		protected override void Initialize()
		{
			//Trace.WriteLine(string.Format(CultureInfo.CurrentCulture, "Entering Initialize() of: {0}", this.ToString()));
			base.Initialize();

			// Add our command handlers for menu (commands must exist in the .vsct file)
			var mcs = GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
			if (null != mcs)
			{
				var menuCommandID = new CommandID(GuidList.guidCommunity_VSBrowserSwitcherCmdSet, PkgCmdIDList.cmdidFirefox);
				var menuItem = new MenuCommand(MenuItemCallback, menuCommandID);
				menuItem.Enabled = BrowserExists(Strings.FIREFOX_FILE_NAME);
				menuItem.Checked = IsCurrentDefault(Strings.FIREFOX_FILE_NAME);
				mcs.AddCommand(menuItem);

				menuCommandID = new CommandID(GuidList.guidCommunity_VSBrowserSwitcherCmdSet, PkgCmdIDList.cmdidInternetExplorer);
				menuItem = new MenuCommand(MenuItemCallback, menuCommandID);
				menuItem.Enabled = BrowserExists(Strings.IE_FILE_NAME);
				menuItem.Checked = IsCurrentDefault(Strings.IE_FILE_NAME);
				mcs.AddCommand(menuItem);

				menuCommandID = new CommandID(GuidList.guidCommunity_VSBrowserSwitcherCmdSet, PkgCmdIDList.cmdidChrome);
				menuItem = new MenuCommand(MenuItemCallback, menuCommandID);
				menuItem.Enabled = BrowserExists(Strings.CHROME_FILE_NAME);
				menuItem.Checked = IsCurrentDefault(Strings.CHROME_FILE_NAME);
				mcs.AddCommand(menuItem);

				menuCommandID = new CommandID(GuidList.guidCommunity_VSBrowserSwitcherCmdSet, PkgCmdIDList.cmdidSafari);
				menuItem = new MenuCommand(MenuItemCallback, menuCommandID);
				menuItem.Enabled = BrowserExists(Strings.SAFARI_FILE_NAME);
				menuItem.Checked = IsCurrentDefault(Strings.SAFARI_FILE_NAME);
				mcs.AddCommand(menuItem);

				menuCommandID = new CommandID(GuidList.guidCommunity_VSBrowserSwitcherCmdSet, PkgCmdIDList.cmdidOpera);
				menuItem = new MenuCommand(MenuItemCallback, menuCommandID);
				menuItem.Enabled = BrowserExists("opera.exe");
				menuItem.Checked = IsCurrentDefault(Strings.OPERA_FILE_NAME);
				mcs.AddCommand(menuItem);

				menuCommandID = new CommandID(GuidList.guidCommunity_VSBrowserSwitcherCmdSet, PkgCmdIDList.cmdidSpartan);
				menuItem = new MenuCommand(MenuItemCallback, menuCommandID);
				menuItem.Enabled = BrowserExists(Strings.SPARTAN_FILE_NAME);
				menuItem.Checked = IsCurrentDefault(Strings.SPARTAN_FILE_NAME);
				mcs.AddCommand(menuItem);
			}

            _watcher.Filter = "*.xml";
            _watcher.Changed += _watcher_Changed;
			_watcher.Created += _watcher_Changed;

            EnableWatcher();
		}

		protected override void Dispose(bool disposing)
		{
			_watcher.EnableRaisingEvents = false;
			_watcher.Dispose();
			base.Dispose(disposing);
		}

		#endregion

        void EnableWatcher()
        {
            //watch the file for changes from VS
            var watcherPath = GetBrowserFilePath().Replace("browsers.xml", string.Empty);

            if (string.IsNullOrEmpty(watcherPath))
                return;

            _watcher.Path = watcherPath;

            _watcher.EnableRaisingEvents = true;
        }

		void _watcher_Changed(object sender, FileSystemEventArgs e)
		{
			if (e.Name == "browsers.xml")
			{
				ParseNewBrowser(e.FullPath);
			}
		}

		/// <summary>
		/// This function is the callback used to execute a command when the a menu item is clicked.
		/// See the Initialize method to see how the menu item is associated to this function using
		/// the OleMenuCommandService service and the MenuCommand class.
		/// </summary>
		private void MenuItemCallback(object sender, EventArgs e)
		{
			var cmd = sender as MenuCommand;
			var browserFile = GetBrowserFilePath();

			if (!string.IsNullOrEmpty(browserFile))
			{
				_watcher.EnableRaisingEvents = false;
				UpdateBrowserFile(browserFile, cmd.CommandID.ID);
				UpdateBrowserRegistry(browserFile);
				EnableWatcher();
			}

			UncheckAllButtons();

			cmd.Checked = true;

			// Show a Message Box to prove we were here
			//IVsUIShell uiShell = (IVsUIShell)GetService(typeof(SVsUIShell));
			//Guid clsid = Guid.Empty;
			//int result;
			//Microsoft.VisualStudio.ErrorHandler.ThrowOnFailure(uiShell.ShowMessageBox(
			//           0,
			//           ref clsid,
			//           "VS Browser Switcher",
			//           "Done",
			//           string.Empty,
			//           0,
			//           OLEMSGBUTTON.OLEMSGBUTTON_OK,
			//           OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST,
			//           OLEMSGICON.OLEMSGICON_INFO,
			//           0,        // false
			//           out result));
		}

		void UncheckAllButtons()
		{
			var mcs = GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
			if (null != mcs)
			{
				var command = mcs.FindCommand(new CommandID(GuidList.guidCommunity_VSBrowserSwitcherCmdSet, PkgCmdIDList.cmdidFirefox));
				command.Checked = false;
				command = mcs.FindCommand(new CommandID(GuidList.guidCommunity_VSBrowserSwitcherCmdSet, PkgCmdIDList.cmdidInternetExplorer));
				command.Checked = false;
				command = mcs.FindCommand(new CommandID(GuidList.guidCommunity_VSBrowserSwitcherCmdSet, PkgCmdIDList.cmdidChrome));
				command.Checked = false;
				command = mcs.FindCommand(new CommandID(GuidList.guidCommunity_VSBrowserSwitcherCmdSet, PkgCmdIDList.cmdidSafari));
				command.Checked = false;
				command = mcs.FindCommand(new CommandID(GuidList.guidCommunity_VSBrowserSwitcherCmdSet, PkgCmdIDList.cmdidOpera));
				command.Checked = false;
				command = mcs.FindCommand(new CommandID(GuidList.guidCommunity_VSBrowserSwitcherCmdSet, PkgCmdIDList.cmdidSpartan));
				command.Checked = false;
			}
		}

		private string GetBrowserFilePath()
		{
			var dte = (DTE)GetService(typeof(DTE));
			var root = @"HKEY_CURRENT_USER\" + dte.RegistryRoot;
			var path = Path.Combine(root, @"WebBrowser\ConfigTimestamp");
			var result = Registry.GetValue(path, "CacheFilePath", null);

			return result != null ? result.ToString() : string.Empty;
		}

		private void UpdateBrowserFile(string browserFilePath, int commandId)
		{
			string name;
			string path;
			string service;

			switch (commandId)
			{
				case PkgCmdIDList.cmdidFirefox:
					name = Strings.FIREFOX_FRIENDLY_NAME;
					path = Strings.FIREFOX_FILE_NAME;
					service = Strings.FIREFOX_SERVICE_NAME;
					break;
				case PkgCmdIDList.cmdidInternetExplorer:
					name = Strings.IE_FRIENDLY_NAME;
					path = Strings.IE_FILE_NAME;
					service = Strings.IE_SERVICE_NAME;
					break;
				case PkgCmdIDList.cmdidChrome:
					name = Strings.CHROME_FRIENDLY_NAME;
					path = Strings.CHROME_FILE_NAME;
					service = Strings.CHROME_SERVICE_NAME;
					break;
				case PkgCmdIDList.cmdidSafari:
					name = Strings.SAFARI_FRIENDLY_NAME;
					path = Strings.SAFARI_FILE_NAME;
					service = Strings.SAFARI_SERVICE_NAME;
					break;
				case PkgCmdIDList.cmdidOpera:
					name = Strings.OPERA_FRIENDLY_NAME;
					path = Strings.OPERA_FILE_NAME;
					service = Strings.OPERA_SERVICE_NAME;
					break;
				case PkgCmdIDList.cmdidSpartan:
					name = Strings.SPARTAN_FRIENDLY_NAME;
					path = Strings.SPARTAN_FILE_NAME;
					service = Strings.SPARTAN_SERVICE_NAME;
					break;
				default:
					return;
			}

			//fixup path
			var key = Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\" + path, string.Empty, string.Empty);

			if (key != null && !string.IsNullOrEmpty(key.ToString()))
			{
				path = key.ToString();
				if (!path.StartsWith("\""))
					path = "\"" + path;

				if (!path.EndsWith("\""))
					path += "\"";

				var doc = new XmlDocument();
				try
				{
					doc.Load(browserFilePath);
				}
				catch (FileNotFoundException)
				{
					doc.LoadXml(@"<?xml version=""1.0""?><BrowserInfo><Browser><Name>Firefox</Name><Path>""C:\Program Files (x86)\Mozilla Firefox\firefox.exe""</Path><Resolution>0</Resolution><IsDefault>True</IsDefault><DDE><Service>FIREFOX</Service><TopicOpenURL>WWW_OpenURL</TopicOpenURL><ItemOpenURL>%s,,0xffffffff,3,,,</ItemOpenURL><TopicActivate>WWW_Activate</TopicActivate><ItemActivate>0xffffffff</ItemActivate></DDE></Browser></BrowserInfo>");
				}
				doc.SelectSingleNode(@"//BrowserInfo/Browser[IsDefault/text()='True']/Name").InnerText = name;
				doc.SelectSingleNode(@"//BrowserInfo/Browser[IsDefault/text()='True']/Path").InnerText = path;
				var el = doc.SelectSingleNode(@"//BrowserInfo/Browser[IsDefault/text()='True']/DDE/Service");
				if (el != null)
				{
					el.InnerText = service;
				}
				doc.Save(browserFilePath);
			}
		}

		void UpdateBrowserRegistry(string browserFilePath)
		{
			var file = new FileInfo(browserFilePath);

			var dte = (DTE)GetService(typeof(DTE));
			var root = @"HKEY_CURRENT_USER\" + dte.RegistryRoot;
			var path = Path.Combine(root, @"WebBrowser\ConfigTimestamp");
			Registry.SetValue(path, "CacheFileDateLastMod", file.LastWriteTimeUtc.ToFileTime(), RegistryValueKind.QWord);
			Registry.SetValue(path, "CacheFileSizeBytes", file.Length, RegistryValueKind.DWord);
			Registry.SetValue(path, "LastConfigurationTimestamp", DateTime.Now.ToUniversalTime().ToFileTime(), RegistryValueKind.QWord);

		}

		bool BrowserExists(string browserName)
		{
			var key = Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\" + browserName, string.Empty, string.Empty);

			if (key != null)
			{
				return key.ToString() != string.Empty;
			}
			return false;
		}

		bool IsCurrentDefault(string browserName)
		{
			var browserFile = GetBrowserFilePath();

            if (string.IsNullOrEmpty(browserFile))
                return false;

			var doc = new XmlDocument();
			try
			{
				doc.Load(browserFile);
			}
			catch
			{
				return false;
			}
			var current = doc.SelectSingleNode(@"//BrowserInfo/Browser[IsDefault/text()='True']/Path").InnerText;

			return current.ToLower().Contains(browserName.ToLower());
		}

		void ParseNewBrowser(string path)
		{
			var doc = new XmlDocument();

			try
			{
				doc.Load(new FileStream(path, FileMode.Open, FileAccess.Read));
			}
			catch
			{
				return;
			}

			var browser = doc.SelectSingleNode(@"//BrowserInfo/Browser[IsDefault/text()='True']/Name").InnerText;

			UncheckAllButtons();
			var mcs = GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
			MenuCommand command;
			switch (browser)
			{
				case Strings.FIREFOX_FRIENDLY_NAME:
					command = mcs.FindCommand(new CommandID(GuidList.guidCommunity_VSBrowserSwitcherCmdSet, PkgCmdIDList.cmdidFirefox));
					command.Checked = true;
					break;
				case Strings.IE_FRIENDLY_NAME:
					command = mcs.FindCommand(new CommandID(GuidList.guidCommunity_VSBrowserSwitcherCmdSet, PkgCmdIDList.cmdidInternetExplorer));
					command.Checked = true;
					break;
				case Strings.CHROME_FRIENDLY_NAME:
					command = mcs.FindCommand(new CommandID(GuidList.guidCommunity_VSBrowserSwitcherCmdSet, PkgCmdIDList.cmdidChrome));
					command.Checked = true;
					break;
				case Strings.SAFARI_FRIENDLY_NAME:
					command = mcs.FindCommand(new CommandID(GuidList.guidCommunity_VSBrowserSwitcherCmdSet, PkgCmdIDList.cmdidSafari));
					command.Checked = true;
					break;
				case Strings.OPERA_FRIENDLY_NAME:
					command = mcs.FindCommand(new CommandID(GuidList.guidCommunity_VSBrowserSwitcherCmdSet, PkgCmdIDList.cmdidOpera));
					command.Checked = true;
					break;
				case Strings.SPARTAN_FRIENDLY_NAME:
					command = mcs.FindCommand(new CommandID(GuidList.guidCommunity_VSBrowserSwitcherCmdSet, PkgCmdIDList.cmdidSpartan));
					command.Checked = true;
					break;
			}
		}
	}
}
// MD Doc Link - dock pane + Tool Add-ins toggle
// Author: MD (Mayank Dholariya)
using System;
using System.IO;
using System.Windows.Forms;
using Autodesk.Navisworks.Api.Plugins;
using DocLink.Core;
using NavisApp = Autodesk.Navisworks.Api.Application;

namespace MDDocLink
{
    internal static class Ids
    {
        internal const string Developer = "MDDL";
        internal const string PaneName = "MDDocLink.Pane";
        internal const string PaneId = PaneName + "." + Developer;
        internal const string DisplayName = Names.Product;
        internal const string ToolTip = Names.Product + " - link documents to model items from Excel | " + Names.Author;
        internal const string Version = Names.Version;
    }

    internal static class Diagnostics
    {
        private static readonly object Gate = new object();
        public static string Folder
        {
            get { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "MDDocLink"); }
        }
        public static string LogPath { get { return Path.Combine(Folder, "mddoclink.log"); } }
        public static void Log(string message)
        {
            try
            {
                Directory.CreateDirectory(Folder);
                lock (Gate)
                    File.AppendAllText(LogPath, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "  " + message + Environment.NewLine);
            }
            catch { }
        }
    }

    [Plugin(Ids.PaneName, Ids.Developer, DisplayName = Ids.DisplayName, ToolTip = Ids.ToolTip)]
    [DockPanePlugin(420, 720, AutoScroll = false, FixedSize = false, MinimumWidth = 320, MinimumHeight = 420)]
    public sealed class PanePlugin : DockPanePlugin
    {
        public override Control CreateControlPane()
        {
            Diagnostics.Log("CreateControlPane md-doclink-" + Ids.Version + ": " + typeof(PanePlugin).Assembly.Location);
            var control = new DocLinkPaneControl();
            control.Dock = DockStyle.Fill;
            try { control.CreateControl(); } catch (Exception ex) { Diagnostics.Log("CreateControl: " + ex); }
            return control;
        }

        public override void DestroyControlPane(Control pane)
        {
            if (pane != null) pane.Dispose();
        }
    }

    [Plugin("MDDocLink.Show", Ids.Developer, DisplayName = Ids.DisplayName, ToolTip = Ids.ToolTip,
        ExtendedToolTip = "Show or hide MD Doc Link: read an Excel list and add document hyperlinks (PDF, folder, URL) to matching model items. " + Names.Author)]
    [AddInPlugin(AddInLocation.AddIn, Icon = "Images\\MDDocLink_16.png", LargeIcon = "Images\\MDDocLink_32.png",
        CanToggle = true, LoadForCanExecute = true)]
    public sealed class ShowCommand : AddInPlugin
    {
        public override CommandState CanExecute()
        {
            var state = new CommandState(true);
            try
            {
                PluginRecord record = NavisApp.Plugins.FindPlugin(Ids.PaneId);
                var pane = record != null ? record.LoadedPlugin as DockPanePlugin : null;
                state.IsChecked = pane != null && pane.Visible;
            }
            catch { state.IsChecked = false; }
            return state;
        }

        public override int Execute(params string[] parameters)
        {
            if (NavisApp.IsAutomated) return 0;
            PluginRecord record = NavisApp.Plugins.FindPlugin(Ids.PaneId);
            if (record == null)
            {
                MessageBox.Show("MD Doc Link panel was not found. Please reinstall.", Ids.DisplayName,
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return 0;
            }
            try
            {
                if (record.LoadedPlugin == null) record.LoadPlugin();
            }
            catch (Exception ex)
            {
                Diagnostics.Log("LoadPlugin failed: " + ex);
                MessageBox.Show("MD Doc Link panel failed to load.\r\n\r\n" + ex.Message +
                    "\r\n\r\nLog: %APPDATA%\\MDDocLink\\mddoclink.log", Ids.DisplayName,
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return 0;
            }
            var pane = record.LoadedPlugin as DockPanePlugin;
            if (pane == null) return 0;
            pane.Visible = !pane.Visible;
            if (pane.Visible) pane.ActivatePane();
            return 0;
        }
    }
}

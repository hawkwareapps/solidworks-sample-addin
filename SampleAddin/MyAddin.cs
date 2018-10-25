using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorks.Interop.swpublished;

namespace SampleAddin
{
    /// <summary>
    /// SolidWorks Sample addin.  Create a new unique GUID for your addin.
    /// 
    /// Go to Properties -> Application -> Assembly Information, and check Make assembly COM-Visible. 
    /// Go to Properties -> Build, and set the platform target to x86 (for 32bit) or x64 (for 64bit).
    /// Go to Properties -> Build, and check Register for COM interop.  This will register the assembly and allow you to debug the addin.  Make sure you run Visual Studio as Administrator.
    /// Go to Properties -> Debug, and select Start external program and select the sldworks.exe in the SolidWorks installation directory.
    /// 
    /// </summary>
    [Guid("B2D0B8D5-3B07-4905-A946-3054AEFB267A"), ComVisible(true)]
    public class MyAddin : ISwAddin
    {
        private ISldWorks _swApp;
        private int _cookie;
        private ICommandManager _cmdMgr;
        private int _mainCmdGrpId = 15;
        private int _menuItemId1 = 1;

        #region Connect to SolidWorks
        /// <summary>
        /// Called when addin is loaded into SolidWorks
        /// </summary>
        /// <param name="ThisSW"></param>
        /// <param name="Cookie"></param>
        /// <returns></returns>
        public bool ConnectToSW(object ThisSW, int Cookie)
        {
            _swApp = (ISldWorks)ThisSW;
            _cookie = Cookie;
            _swApp.SetAddinCallbackInfo2(0, this, Cookie);
            _cmdMgr = _swApp.GetCommandManager(_cookie);

            SetupUI();

            return true;
        }

        /// <summary>
        /// Called when addin is removed from SolidWorks.
        /// </summary>
        /// <returns></returns>
        public bool DisconnectFromSW()
        {
            // tear down ui
            TearDownUI();
            // destroy reference to solidworks
            _swApp = null;
            // destory reference to command manager
            _cmdMgr = null;
            // garbage collect
            GC.Collect();

            return true;
        }
        #endregion

        #region Setup UI

        private void SetupUI()
        {
            int errors = 0;
            // gather up the icons
            string folder = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetAssembly(this.GetType()).Location);
            // create command group
            ICommandGroup commandGroup = _cmdMgr.CreateCommandGroup2(_mainCmdGrpId, "Custom Addin", "My custom add-in", "", -1, true, ref errors);
            // set the icon lists
            commandGroup.IconList = new string[] { System.IO.Path.Combine(folder, "Icons", "icons.bmp") };
            commandGroup.MainIconList = new string[] { System.IO.Path.Combine(folder, "Icons", "icons.bmp") };

            // add the commands to the command group
            // this is just an example for demo purposes***
            List<int> commandItems = new List<int>();
            for (int i = 0; i < 10; i++) {
                // you wouldn't add your commands like this.
                // you would obviously create them one at a time
                // for example, in this demo, we're adding 10 commands, but they all call the same method ExucuteAction... that would be a dumb addin
                commandItems.Add(commandGroup.AddCommandItem2("Action " + i, -1, "Action " + i, "Action " + i, i, "ExecuteAction", "CanExecuteAction", _menuItemId1 + i, (int)swCommandItemType_e.swMenuItem | (int)swCommandItemType_e.swToolbarItem));
            }

            // activate command group
            commandGroup.HasToolbar = true;
            commandGroup.HasMenu = true;
            commandGroup.Activate();

            // create the command tab
            foreach (int docType in new int[] { (int)swDocumentTypes_e.swDocASSEMBLY, (int)swDocumentTypes_e.swDocPART, (int)swDocumentTypes_e.swDocDRAWING }) {
                ICommandTab commandTab = _cmdMgr.GetCommandTab(docType, "Custom Addin");

                if(commandTab != null) {
                    // recreate the command tab
                    _cmdMgr.RemoveCommandTab((CommandTab)commandTab);
                    commandTab = null;
                }

                // add the command tab
                commandTab = _cmdMgr.AddCommandTab(docType, "Custom Addin");
                // create a command box
                ICommandTabBox commandTabBox = commandTab.AddCommandTabBox();
                // add commands to command tab
                List<int> cmds = new List<int>();
                List<int> textTypes = new List<int>();
                foreach (var cmdItem in commandItems) {
                    cmds.Add(commandGroup.CommandID[cmdItem]);
                    textTypes.Add((int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow);
                }
                bool result = commandTabBox.AddCommands(cmds.ToArray(), textTypes.ToArray());

                commandTab.AddSeparator((CommandTabBox)commandTabBox, cmds.First());
            }
            
        }

        private void TearDownUI()
        {
            _cmdMgr.RemoveCommandGroup(_mainCmdGrpId);
        }

        #endregion

        #region Callbacks

        /// <summary>
        /// For every command added above, you need to have a callback method.  This is where you do your work.
        /// </summary>
        public void ExecuteAction()
        {
            _swApp.SendMsgToUser2("Addin loaded!", (int)swMessageBoxIcon_e.swMbInformation, (int)swMessageBoxBtn_e.swMbOk);
        }

        /// <summary>
        /// For every command add above, you can (optional) have a callback method that let's SolidWorks know whether or not to enable the menu item.
        /// For example, if the command should only run in the active document is an drawing, you could check the active document type here.
        /// </summary>
        /// <returns></returns>
        public int CanExecuteAction()
        {
            // return 1 if action can execute
            return 1;
        }

        #endregion

        #region COM Registration
        /// <summary>
        /// This is some boiler plate stuff to register your addin.  Shouldn't really have to change anything here.
        /// </summary>
        /// <param name="t"></param>
        [ComRegisterFunction]
        public static void RegisterCom(Type t)
        {
            Microsoft.Win32.RegistryKey hklm = Microsoft.Win32.Registry.LocalMachine;
            Microsoft.Win32.RegistryKey hkcu = Microsoft.Win32.Registry.CurrentUser;

            string key = "SOFTWARE\\SolidWorks\\Addins\\{" + t.GUID.ToString() + "}";

            Microsoft.Win32.RegistryKey addinkey = hklm.CreateSubKey(key);
            addinkey.SetValue(null, 0);
            addinkey.SetValue("Description", "My Custom Addin");
            addinkey.SetValue("Title", "Custom Addin");

            key = "Software\\SolidWorks\\AddInsStartup\\{" + t.GUID.ToString() + "}";
            addinkey = hkcu.CreateSubKey(key);
            addinkey.SetValue(null, 1);
        }

        /// <summary>
        /// This is some boiler plate stuff to unregister your addin.  Shouldn't really have to change anything here.
        /// </summary>
        /// <param name="t"></param>
        [ComUnregisterFunction]
        public static void UnregisterCom(Type t)
        {
            Microsoft.Win32.RegistryKey hklm = Microsoft.Win32.Registry.LocalMachine;
            Microsoft.Win32.RegistryKey hkcu = Microsoft.Win32.Registry.CurrentUser;

            string key = "SOFTWARE\\SolidWorks\\Addins\\{" + t.GUID.ToString() + "}";

            hklm.DeleteSubKey(key);

            key = "Software\\SolidWorks\\AddInsStartup\\{" + t.GUID.ToString() + "}";

            hkcu.DeleteSubKey(key);
        }
        #endregion
    }
}

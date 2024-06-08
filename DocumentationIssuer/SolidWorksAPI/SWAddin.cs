using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorks.Interop.swpublished;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SolidWorksAPI
{
    [ComVisible(true)]
    public class SWAddin : ISwAddin
    {
        private SldWorks _swApp;
        private int _addinId;
        private ICommandManager _cmdMgr;
        private const int CMDGroupId = 0;
        private const int documentationIssuerId = 1;
        private const int OFGeneratorId = 2;
        private const int PDFMergerId = 3;
        private const int CSVCreatorId = 4;


        public bool ConnectToSW(object ThisSW, int Cookie)
        {
            _swApp = (SldWorks)ThisSW;
            _addinId = Cookie;
            _swApp.SetAddinCallbackInfo2(0, this, _addinId);

            _cmdMgr = _swApp.GetCommandManager(_addinId);

            try
            {
                AddCommandManager();
                _swApp.ActiveDocChangeNotify += OnActiveDocChange;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro durante a inicialização do add-in: " + ex.Message);
            }

            return true;
        }

        public bool DisconnectFromSW()
        {
            RemoveCommandManager();
            return true;
        }

        [ComRegisterFunction()]
        private static void ComRegister(Type t)
        {
            var keyPath = string.Format(@"SOFTWARE\\Solidworks\\AddIns\\{0:b}", t.GUID);

            using (var rk = Microsoft.Win32.Registry.LocalMachine.CreateSubKey(keyPath))
            {
                rk.SetValue(null, 1);

                rk.SetValue("Title", "Documentation Issuer");
                rk.SetValue("Description", "Addin to issue files to the factory");
                rk.SetValue("Author", "LeonardoMelo");
            }

        }

        [ComUnregisterFunction()]
        private static void ComUnregister(Type t)
        {
            var keyPath = string.Format(@"SOFTWARE\\Solidworks\\AddIns\\{0:b}", t.GUID);

            Microsoft.Win32.Registry.LocalMachine.DeleteSubKeyTree(keyPath);
        }

        private int OnActiveDocChange()
        {
            RemoveCommandManager();
            AddCommandManager();
            //_cmdMgr.HotKeyManager.RebuildMenu();
            return 0;
        }

        private void AddCommandManager()
        {
            int errors = 0;
            bool ignorePrevious = false;

            ICommandGroup cmdGroup = _cmdMgr.CreateCommandGroup2(CMDGroupId, "Documentation Issuer", "Commands for my add-in", "", -1, ignorePrevious, ref errors);
            cmdGroup.IconList = @"C:\Path\To\Your\Icons.bmp";
            cmdGroup.MainIconList = @"C:\Path\To\Your\IconsLarge.bmp";

            cmdGroup.AddCommandItem2("Emitir Documentação", -1, "Tooltip for Command 1", "Emitir Documentação", 0, 
                "DocumentationIssuerCallback", "EnableDocumentationIssuer", documentationIssuerId, (int)swCommandItemType_e.swMenuItem | (int)swCommandItemType_e.swToolbarItem);
            cmdGroup.AddCommandItem2("Gerar OF", -1, "Gerar Ordem de Fabricação", "Gerar OF", 1,
                "OFGeneratorCallback", "EnableOFGenerator", OFGeneratorId, (int)swCommandItemType_e.swMenuItem | (int)swCommandItemType_e.swToolbarItem);
            cmdGroup.AddCommandItem2("Mesclar PDFs", -1, "Juntar os PDFs ", "Mesclar PDFs", 2,
                "PDFMergerCallback", "EnablePDFMerger", PDFMergerId, (int)swCommandItemType_e.swMenuItem | (int)swCommandItemType_e.swToolbarItem);
            cmdGroup.AddCommandItem2("Criar CSV", -1, "Tooltip for Command 4", "Criar CSV", 3,
                "CSVCreatorCallback", "EnableCSVCreator", CSVCreatorId, (int)swCommandItemType_e.swMenuItem | (int)swCommandItemType_e.swToolbarItem);

            cmdGroup.HasToolbar = true;
            cmdGroup.HasMenu = true;
            cmdGroup.Activate();

            CommandTab cmdTab = _cmdMgr.GetCommandTab(_addinId, "My Commands");
            if (cmdTab != null)
            {
                _cmdMgr.RemoveCommandTab(cmdTab);
            }
            cmdTab = _cmdMgr.AddCommandTab((int)swDocumentTypes_e.swDocPART, "My Commands");

            // Verificar se a aba já existe
            if (cmdTab != null)
            {
                // Criar uma Command Tab Box
                CommandTabBox cmdBox = cmdTab.AddCommandTabBox();

                int[] cmdIDs = new int[4];
                int[] TextType = new int[4];

                cmdIDs[0] = cmdGroup.get_CommandID(documentationIssuerId);
                TextType[0] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal;

                cmdIDs[1] = cmdGroup.get_CommandID(OFGeneratorId);
                TextType[1] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal;

                cmdIDs[2] = cmdGroup.get_CommandID(PDFMergerId);
                TextType[2] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal;

                cmdIDs[3] = cmdGroup.get_CommandID(CSVCreatorId);
                TextType[3] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal;

                cmdBox.AddCommands(cmdIDs, TextType);
            }
        }
        private void RemoveCommandManager()
        {
            _cmdMgr.RemoveCommandGroup(CMDGroupId);

            CommandTab cmdTab = _cmdMgr.GetCommandTab(_addinId, "My Commands");
            if (cmdTab != null)
            {
                _cmdMgr.RemoveCommandTab(cmdTab);
            }
        }

        public void DocumentationIssuerCallback()
        {
            MessageBox.Show("Command 1 executed!");
        }

        public void OFGeneratorCallback()
        {
            MessageBox.Show("Command 2 executed!");
        }

        public void PDFMergerCallback()
        {
            MessageBox.Show("Command 3 executed!");
        }

        public void CSVCreatorCallback()
        {
            MessageBox.Show("Command 4 executed!");
        }

        public int EnableDocumentationIssuer()
        {
            ModelDoc2 activeDoc = _swApp.ActiveDoc as ModelDoc2;
            if (activeDoc == null)
            {
                return 0;
            }

            swDocumentTypes_e docType = (swDocumentTypes_e)activeDoc.GetType();
            if (docType == swDocumentTypes_e.swDocPART || docType == swDocumentTypes_e.swDocASSEMBLY)
            {
                return 1; 
            }

            return 0; 
        }

        public int EnableOFGenerator()
        {
            ModelDoc2 activeDoc = _swApp.ActiveDoc as ModelDoc2;
            if (activeDoc == null)
            {
                return 0;
            }

            swDocumentTypes_e docType = (swDocumentTypes_e)activeDoc.GetType();
            if (docType == swDocumentTypes_e.swDocPART || docType == swDocumentTypes_e.swDocASSEMBLY)
            {
                return 1;
            }

            return 0;
        }
        
        public int EnablePDFMerger()
        {
            ModelDoc2 activeDoc = _swApp.ActiveDoc as ModelDoc2;
            if (activeDoc == null)
            {
                return 0;
            }

            swDocumentTypes_e docType = (swDocumentTypes_e)activeDoc.GetType();
            if (docType == swDocumentTypes_e.swDocPART || docType == swDocumentTypes_e.swDocASSEMBLY)
            {
                return 1;
            }

            return 0;
        }
        
        public int EnableCSVCreator()
        {
            ModelDoc2 activeDoc = _swApp.ActiveDoc as ModelDoc2;
            if (activeDoc == null)
            {
                return 0;
            }

            swDocumentTypes_e docType = (swDocumentTypes_e)activeDoc.GetType();
            if (docType == swDocumentTypes_e.swDocPART || docType == swDocumentTypes_e.swDocASSEMBLY)
            {
                return 1;
            }

            return 0;
        }
    }
}

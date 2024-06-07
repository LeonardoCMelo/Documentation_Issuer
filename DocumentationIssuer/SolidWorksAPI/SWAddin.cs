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
    public class SWAddin : SwAddin
    {
        public bool ConnectToSW(object ThisSW, int Cookie)
        {
            MessageBox.Show("Deu certo!");
            return true;
        }

        public bool DisconnectFromSW()
        {
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
    }
}

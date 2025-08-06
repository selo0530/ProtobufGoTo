using Microsoft.VisualStudio.TextManager.Interop;
using stdole;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace ProtobufGoTo
{
    [Guid(ProtobufGoToLanguageService.LanguageServiceString)]
    public class ProtobufGoToLanguageService
    {
        public const string LanguageServiceString = "f970e673-a3ad-4159-8447-a6a5344cfe8f";

        public int GetLanguageName(out string bstrName)
        {
            bstrName = "Protocol Buffers";
            return 0;
        }

        public int GetFileExtensions(out string pbstrExtensions)
        {
            pbstrExtensions = ".proto";
            return 0;
        }

        public int GetColorizer(IVsTextLines pBuffer, out IVsColorizer ppColorizer)
        {
            ppColorizer = null;
            return 1;
        }
    }
}

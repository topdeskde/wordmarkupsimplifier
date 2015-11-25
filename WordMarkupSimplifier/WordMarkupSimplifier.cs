using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenXmlPowerTools;
using DocumentFormat.OpenXml.Packaging;

namespace WordMarkupSimplifier
{
    class WordMarkupSimplifier
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Word Markup Simplifier 1.0.0");

            if (args.Length < 1)
            {
                Console.WriteLine("Usage: WordMarkupSimplifier <some word document>");
                return;
            }
            Console.WriteLine("Simplifying " + args[0]);

            try
            {
                using (WordprocessingDocument doc = WordprocessingDocument.Open(args[0], true))
                {
                    SimplifyMarkupSettings settings = new SimplifyMarkupSettings
                    {
                        RemoveComments = true,
                        RemoveContentControls = true,
                        RemoveEndAndFootNotes = true,
                        RemoveFieldCodes = false,
                        RemoveLastRenderedPageBreak = true,
                        RemovePermissions = true,
                        RemoveProof = true,
                        RemoveRsidInfo = true,
                        RemoveSmartTags = true,
                        RemoveSoftHyphens = true,
                        ReplaceTabsWithSpaces = true,
                    };
                    MarkupSimplifier.SimplifyMarkup(doc, settings);
                }
            }
            catch (OpenXmlPackageException e)
            {
                Console.WriteLine("Error: " + args[0] + " is not a valid Word document.");
                Console.WriteLine("Exception: " + e.Message);
            }
        }
    }
}

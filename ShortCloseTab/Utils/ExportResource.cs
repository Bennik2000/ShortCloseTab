using System;
using System.IO;
using System.Reflection;

namespace ShortCloseTab.Utils
{
    static class ExportResource
    {
        public static void Export(string resourceName, string outputPath, Assembly assembly)
        {
            if (File.Exists(outputPath))
            {
                File.Delete(outputPath);
            }

            using(Stream inputStream = assembly.GetManifestResourceStream(resourceName))
            using (Stream outputStream = File.Create(outputPath))
            {
                if (inputStream != null)
                {
                    byte[] buffer = new byte[8192];

                    int bytesRead;
                    while ((bytesRead = inputStream.Read(buffer, 0, buffer.Length)) > 0)
                    {
                        outputStream.Write(buffer, 0, bytesRead);
                    }
                }
                else
                {
                    throw new InvalidOperationException("Ressource not found");
                }
            }
        }
    }
}

using CoverageKiller2.Logging;
using System;
using System.IO;
using System.Xml.Serialization;

namespace CoverageKiller2.Pipeline.Config
{
    public class ProcessorConfigLoader
    {
        public ProcessorConfig ProcessorConfig { get; private set; }
        public Tracer Tracer { get; } = new Tracer(typeof(ProcessorConfig));

        public bool LoadConfig(string xmlFilePath)
        {
            if (string.IsNullOrEmpty(xmlFilePath))
            {
                Tracer.Log("Error: XML file path is null or empty.");
                return false;
            }

            if (!File.Exists(xmlFilePath))
            {
                Tracer.Log($"Error: File not found at {xmlFilePath}");
                return false;
            }

            FileStream fileStream = null;
            try
            {
                XmlSerializer serializer = new XmlSerializer(typeof(ProcessorConfig));
                fileStream = new FileStream(xmlFilePath, FileMode.Open, FileAccess.Read);
                ProcessorConfig = (ProcessorConfig)serializer.Deserialize(fileStream);
                return true;
            }
            catch (Exception ex)
            {
                Tracer.Log($"Error loading XML: {ex.Message}");
                return false;
            }
            finally
            {
                if (fileStream != null)
                {
                    fileStream.Close();
                    fileStream.Dispose();
                }
            }
        }
    }
}

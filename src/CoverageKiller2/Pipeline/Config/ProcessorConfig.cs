using System.Collections.Generic;
using System.Xml.Serialization;

namespace CoverageKiller2.Pipeline.Config
{
    [XmlRoot("ProcessorConfig")]
    public class ProcessorConfig
    {
        [XmlAttribute("Name")]
        public string Name { get; set; }

        [XmlAttribute("Description")]
        public string Description { get; set; }

        [XmlAttribute("SourceTemplate")]
        public string SourceTemplate { get; set; }

        [XmlElement("Pipeline")]
        public Pipeline Pipeline { get; set; }
    }

    public class Pipeline
    {
        [XmlElement("Steps")]
        public Steps Steps { get; set; }
    }

    public class Steps
    {
        [XmlElement("Step")]
        public List<Step> StepList { get; set; } = new List<Step>();
    }

    public class Step
    {
        [XmlAttribute("Name")]
        public string Name { get; set; }

        [XmlAttribute("Namespace")]
        public string Namespace { get; set; }
    }
}

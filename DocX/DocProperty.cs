using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace Novacode
{
    /// <summary>
    ///     Represents a field of type document property. This field displays the value stored in a custom property.
    /// </summary>
    public class DocProperty : DocXElement
    {
        private readonly Regex _extractName = new Regex(@"DOCPROPERTY  (?<name>.*)  ");
        private readonly string _name;

        /// <summary>
        ///     The custom property to display.
        /// </summary>
        public string Name { get { return _name; } }

        internal DocProperty(DocX document, XElement xml)
            : base(document, xml)
        {
            string instr = Xml.Attribute(XName.Get("instr", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")).Value;
            _name = _extractName.Match(instr.Trim()).Groups["name"].Value;
        }
    }
}
using System;
using System.Globalization;

namespace Novacode
{
    public class CustomProperty
    {
        private string _name;
        private object _value;
        private string _type;

        /// <summary>
        /// The name of this CustomProperty.
        /// </summary>
        public string Name { get { return _name;} }

        /// <summary>
        /// The value of this CustomProperty.
        /// </summary>
        public object Value { get { return _value; } }

        internal string Type { get { return _type; } }

        internal CustomProperty(string name, string type, string value)
        {
            object realValue;
            switch (type)
            {
                case "lpwstr": 
                {
                    realValue = value;
                    break;
                }

                case "i4":
                {
                    realValue = int.Parse(value, CultureInfo.InvariantCulture);
                    break;
                }

                case "r8":
                {
                    realValue = Double.Parse(value, CultureInfo.InvariantCulture);
                    break;
                }

                case "filetime":
                {
                    realValue = DateTime.Parse(value, CultureInfo.InvariantCulture);
                    break;
                }

                case "bool":
                {
                    realValue = bool.Parse(value);
                    break;
                }

                default: throw new Exception();
            }

            _name = name;
            _type = type;
            _value = realValue;
        }

        private CustomProperty(string name, string type, object value)
        {
            _name = name;
            _type = type;
            _value = value;
        }

        /// <summary>
        /// Create a new CustomProperty to hold a string.
        /// </summary>
        /// <param name="name">The name of this CustomProperty.</param>
        /// <param name="value">The value of this CustomProperty.</param>
        public CustomProperty(string name, string value) : this(name, "lpwstr", value as object) { }


        /// <summary>
        /// Create a new CustomProperty to hold an int.
        /// </summary>
        /// <param name="name">The name of this CustomProperty.</param>
        /// <param name="value">The value of this CustomProperty.</param>
        public CustomProperty(string name, int value) : this(name, "i4", value) { }


        /// <summary>
        /// Create a new CustomProperty to hold a double.
        /// </summary>
        /// <param name="name">The name of this CustomProperty.</param>
        /// <param name="value">The value of this CustomProperty.</param>
        public CustomProperty(string name, double value) : this(name, "r8", value) { }


        /// <summary>
        /// Create a new CustomProperty to hold a DateTime.
        /// </summary>
        /// <param name="name">The name of this CustomProperty.</param>
        /// <param name="value">The value of this CustomProperty.</param>
        public CustomProperty(string name, DateTime value) : this(name, "filetime", value.ToUniversalTime()) { }

        /// <summary>
        /// Create a new CustomProperty to hold a bool.
        /// </summary>
        /// <param name="name">The name of this CustomProperty.</param>
        /// <param name="value">The value of this CustomProperty.</param>
        public CustomProperty(string name, bool value) : this(name, "bool", value) { }
    }
}

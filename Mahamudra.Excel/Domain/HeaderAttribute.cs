using System;

namespace Mahamudra.Excel.Domain
{
    /// <summary>
    /// Attribute to define Excel column header properties for a class property.
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, Inherited = true, AllowMultiple = false)]
    public sealed class HeaderAttribute : Attribute
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="HeaderAttribute"/> class.
        /// </summary>
        /// <param name="caption">The column header text.</param>
        /// <param name="readOnly">Whether the column is read-only.</param>
        /// <exception cref="ArgumentException">Thrown when caption is null or whitespace.</exception>
        public HeaderAttribute(string caption, bool readOnly = false)
        {
            if (string.IsNullOrWhiteSpace(caption))
                throw new ArgumentException("Caption cannot be null or whitespace.", nameof(caption));

            Caption = caption;
            ReadOnly = readOnly;
        }

        /// <summary>
        /// Gets the column header text.
        /// </summary>
        public string Caption { get; }

        /// <summary>
        /// Gets or sets the property name (set internally during reflection).
        /// </summary>
        public string Name { get; internal set; } = string.Empty;

        /// <summary>
        /// Gets whether the column is read-only.
        /// </summary>
        public bool ReadOnly { get; }

        /// <summary>
        /// Gets or sets the property type (set internally during reflection).
        /// </summary>
        public Type? Type { get; internal set; }

        /// <summary>
        /// Gets or sets the column order. Lower values appear first.
        /// </summary>
        public short Order { get; set; } = 0;

        /// <summary>
        /// Gets or sets the cell formatting style.
        /// </summary>
        public XCellStyle Style { get; set; } = XCellStyle.None;
    }
}

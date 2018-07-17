using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DynamicsCRMCustomizationToolForExcel.Model
{
    public class Publisher
    {
        /// <summary>
        /// Gets or sets the organization id.
        /// </summary>
        /// <value>
        /// The organization id.
        /// </value>
        public Guid OrganizationId { get; set; }

        /// <summary>
        /// Gets or sets the name of the unique.
        /// </summary>
        /// <value>
        /// The name of the unique.
        /// </value>
        public string UniqueName { get; set; }

        /// <summary>
        /// Gets or sets the publisher id.
        /// </summary>
        /// <value>
        /// The publisher id.
        /// </value>
        public Guid PublisherId { get; set; }

        /// <summary>
        /// Gets or sets the customization prefix.
        /// </summary>
        /// <value>
        /// The customization prefix.
        /// </value>
        public string CustomizationPrefix { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether this instance is read only.
        /// </summary>
        /// <value>
        /// <c>true</c> if this instance is read only; otherwise, <c>false</c>.
        /// </value>
        public bool IsReadOnly { get; set; }

        /// <summary>
        /// Gets or sets the name of the friendly.
        /// </summary>
        /// <value>
        /// The name of the friendly.
        /// </value>
        public string FriendlyName { get; set; }
    }
}

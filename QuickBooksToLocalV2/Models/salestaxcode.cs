//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated from a template.
//
//     Manual changes to this file may cause unexpected behavior in your application.
//     Manual changes to this file will be overwritten if the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace QuickBooksToLocalV2.Models
{
    using System;
    using System.Collections.Generic;
    
    public partial class salestaxcode
    {
        public string salestaxcodesID { get; set; }
        public string salestaxcodesName { get; set; }
        public Nullable<bool> IsActive { get; set; }
        public string ClassRef_FullName { get; set; }
        public string ClassRef_ListID { get; set; }
        public string ItemDesc { get; set; }
        public Nullable<double> TaxRate { get; set; }
        public string TaxVendorRef_FullName { get; set; }
        public string TaxVendorRef_ListID { get; set; }
        public string CustomFields { get; set; }
        public string EditSequence { get; set; }
        public Nullable<System.DateTime> TimeModified { get; set; }
        public Nullable<System.DateTime> TimeCreated { get; set; }
    }
}

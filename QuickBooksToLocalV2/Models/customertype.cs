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
    
    public partial class customertype
    {
        public string ID { get; set; }
        public string Name { get; set; }
        public string FullName { get; set; }
        public string ParentName { get; set; }
        public string ParentId { get; set; }
        public Nullable<bool> IsActive { get; set; }
        public Nullable<System.DateTime> TimeCreated { get; set; }
        public Nullable<System.DateTime> TimeModified { get; set; }
        public string EditSequence { get; set; }
        public Nullable<int> Sublevel { get; set; }
    }
}

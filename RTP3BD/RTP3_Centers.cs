//------------------------------------------------------------------------------
// <auto-generated>
//     Этот код создан по шаблону.
//
//     Изменения, вносимые в этот файл вручную, могут привести к непредвиденной работе приложения.
//     Изменения, вносимые в этот файл вручную, будут перезаписаны при повторном создании кода.
// </auto-generated>
//------------------------------------------------------------------------------

namespace RTP3BD
{
    using System;
    using System.Collections.Generic;
    
    public partial class RTP3_Centers
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public RTP3_Centers()
        {
            this.RTP3_Sections = new HashSet<RTP3_Sections>();
        }
    
        public string Name { get; set; }
        public int GUID { get; set; }
        public int OwnerGUID { get; set; }
        public Nullable<System.DateTime> DBB { get; set; }
        public System.Guid ID_BD { get; set; }
    
        public virtual RTP3_GUID RTP3_GUID { get; set; }
        public virtual RTP3_RES RTP3_RES { get; set; }
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<RTP3_Sections> RTP3_Sections { get; set; }
    }
}

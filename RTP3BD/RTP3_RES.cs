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
    
    public partial class RTP3_RES
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public RTP3_RES()
        {
            this.RTP3_Centers = new HashSet<RTP3_Centers>();
        }
    
        public string Name { get; set; }
        public int GUID { get; set; }
        public int OwnerGUID { get; set; }
        public System.Guid ID_BD { get; set; }
    
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<RTP3_Centers> RTP3_Centers { get; set; }
        public virtual RTP3_GUID RTP3_GUID { get; set; }
        public virtual RTP3_PES RTP3_PES { get; set; }
    }
}

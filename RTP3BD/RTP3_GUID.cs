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
    
    public partial class RTP3_GUID
    {
        public int GUID { get; set; }
        public Nullable<short> RELATION_ID { get; set; }
        public System.Guid ID_BD { get; set; }
    
        public virtual RTP3_AOEnergo RTP3_AOEnergo { get; set; }
        public virtual RTP3_BD RTP3_BD { get; set; }
        public virtual RTP3_Centers RTP3_Centers { get; set; }
        public virtual RTP3_Coors RTP3_Coors { get; set; }
        public virtual RTP3_PES RTP3_PES { get; set; }
        public virtual RTP3_RES RTP3_RES { get; set; }
        public virtual RTP3_Sections RTP3_Sections { get; set; }
    }
}

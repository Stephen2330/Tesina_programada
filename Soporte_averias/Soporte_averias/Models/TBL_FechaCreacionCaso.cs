//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated from a template.
//
//     Manual changes to this file may cause unexpected behavior in your application.
//     Manual changes to this file will be overwritten if the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace Soporte_averias.Models
{
    using System;
    using System.Collections.Generic;
	using System.ComponentModel.DataAnnotations;

    public partial class TBL_FechaCreacionCaso
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public TBL_FechaCreacionCaso()
        {
            this.TBL_Caso = new HashSet<TBL_Caso>();
        }

        public int TN_IdFechaCreacionCaso { get; set; }


        [Required(ErrorMessage ="La fecha de creaci�n de caso es obligatoria")]
        [DisplayFormat(DataFormatString = "{0:dd-MM-yyyy}")]
		public Nullable<System.DateTime> TD_FechaCreacionCaso { get; set; }

		[Required(ErrorMessage = "La descripci�n es obligatoria")]
		public string TC_Descripcion { get; set; }
    
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<TBL_Caso> TBL_Caso { get; set; }
    }
}

using System.ComponentModel.DataAnnotations;

namespace SlideMergerAPINew.Models
{
    public class SlideMergeRequest
    {
        [Required]
        public string Mba { get; set; } = string.Empty;

        [Required]
        public string Theme { get; set; } = string.Empty;

        [Required]
        public string TituloAula { get; set; } = string.Empty;

        [Required]
        public string NomeProfessor { get; set; } = string.Empty;

        [Required]
        public string LinkedinPerfil { get; set; } = string.Empty;
    }
}

namespace SlideMergerAPINew.Models
{
    public class SlideMergeResponse
    {
        public bool Success { get; set; }
        public string Message { get; set; } = string.Empty;
        public string? DownloadUrl { get; set; }
        public string? FileName { get; set; }
    }
}

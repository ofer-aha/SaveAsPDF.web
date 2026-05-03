public class AttachmentDto
{
    public string? Name { get; set; }
    public string? ContentType { get; set; }
    public long Size { get; set; }
    public bool IsInline { get; set; }
    public string? Base64 { get; set; }
}

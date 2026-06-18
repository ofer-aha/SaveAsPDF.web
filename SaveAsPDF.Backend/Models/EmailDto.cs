public class EmailDto
{
    public string? Subject { get; set; }
    public string? From { get; set; }
    public List<string>? To { get; set; }
    public List<string>? Cc { get; set; }
    public DateTime? SentDate { get; set; }
    public DateTime? ReceivedDate { get; set; }
    public string? InternetMessageId { get; set; }
    public string? BodyHtml { get; set; }
}

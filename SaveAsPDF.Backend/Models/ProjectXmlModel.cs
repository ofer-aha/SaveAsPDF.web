using System.Xml.Serialization;

[XmlRoot("ProjectModel")]
public class ProjectXmlModel
{
    public int Id { get; set; }
    public string? ProjectNumber { get; set; }
    public string? ProjectName { get; set; }
    public bool NoteToProjectLeader { get; set; }
    public string? DefaultSaveFolder { get; set; }
    public string? ProjectNotes { get; set; }
    public string? LastSavePath { get; set; }
    public string? ProjectDate { get; set; }
}

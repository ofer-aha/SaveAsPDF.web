using System.Xml.Serialization;

[XmlType("EmployeeModel")]   // file uses <EmployeeModel> as the item element
public class EmployeeXmlModel
{
    public int Id { get; set; }
    public string? FirstName { get; set; }
    public string? LastName { get; set; }
    public string? EmailAddress { get; set; }
    public string? PhoneNumber { get; set; }
    public bool IsLeader { get; set; }
}

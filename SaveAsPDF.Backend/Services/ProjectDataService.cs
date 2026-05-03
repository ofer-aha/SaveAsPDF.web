using System.Xml.Serialization;

public class ProjectDataService
{
    // Filenames are kept exactly as the legacy app wrote them (note the typo "Emploeeys").
    private const string ProjectXmlName   = ".SaveAsPDF_Project.xml";
    private const string EmployeesXmlName = ".SaveAsPDF_Emploeeys.xml";
    private const string SaveAsPdfDir     = ".SaveAsPDF";

    public (ProjectXmlModel? Project, List<EmployeeXmlModel> Employees) Load(string projectFolder)
    {
        var dir = Path.Combine(projectFolder, SaveAsPdfDir);

        ProjectXmlModel? project = null;
        var employees = new List<EmployeeXmlModel>();

        var projectFile = Path.Combine(dir, ProjectXmlName);
        if (File.Exists(projectFile))
        {
            try
            {
                var serializer = new XmlSerializer(typeof(ProjectXmlModel));
                using var stream = File.OpenRead(projectFile);
                project = serializer.Deserialize(stream) as ProjectXmlModel;
            }
            catch { /* malformed XML - ignore */ }
        }

        var employeesFile = Path.Combine(dir, EmployeesXmlName);
        if (File.Exists(employeesFile))
        {
            try
            {
                var serializer = new XmlSerializer(
                    typeof(List<EmployeeXmlModel>),
                    new XmlRootAttribute("ArrayOfEmployeeModel"));
                using var stream = File.OpenRead(employeesFile);
                employees = serializer.Deserialize(stream) as List<EmployeeXmlModel>
                            ?? new List<EmployeeXmlModel>();
            }
            catch { /* malformed XML - ignore */ }
        }

        return (project, employees);
    }
}

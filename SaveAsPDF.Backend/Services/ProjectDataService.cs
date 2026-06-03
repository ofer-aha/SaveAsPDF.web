using System.Xml.Serialization;

public class ProjectDataService
{
    // Filenames are kept exactly as the legacy app wrote them (note the typo "Emploeeys").
    private const string ProjectXmlName   = ".SaveAsPDF_Project.xml";
    private const string EmployeesXmlName = ".SaveAsPDF_Emploeeys.xml";
    private const string SaveAsPdfDir     = ".SaveAsPDF";

    // Cache serializers — the employees one uses XmlRootAttribute which generates a temp
    // assembly each time if not reused, causing memory leaks and occasional failures.
    private static readonly XmlSerializer _projectSerializer =
        new XmlSerializer(typeof(ProjectXmlModel));
    private static readonly XmlSerializer _employeesSerializer =
        new XmlSerializer(typeof(List<EmployeeXmlModel>), new XmlRootAttribute("ArrayOfEmployeeModel"));

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
                using var stream = File.OpenRead(projectFile);
                project = _projectSerializer.Deserialize(stream) as ProjectXmlModel;
            }
            catch { /* malformed XML - ignore */ }
        }

        var employeesFile = Path.Combine(dir, EmployeesXmlName);
        if (File.Exists(employeesFile))
        {
            try
            {
                using var stream = File.OpenRead(employeesFile);
                employees = _employeesSerializer.Deserialize(stream) as List<EmployeeXmlModel>
                            ?? new List<EmployeeXmlModel>();
            }
            catch { /* malformed XML - ignore */ }
        }

        return (project, employees);
    }

    // Creates the .SaveAsPDF folder and empty XML files only if they don't already exist.
    // All exceptions are suppressed — this is a best-effort silent operation.
    public void EnsureInitialized(string projectFolder, string projectNumber)
    {
        try
        {
            var dir = Path.Combine(projectFolder, SaveAsPdfDir);
            if (Directory.Exists(dir)) return;

            Directory.CreateDirectory(dir);
            SetHidden(dir);

            var projectFile   = Path.Combine(dir, ProjectXmlName);
            var employeesFile = Path.Combine(dir, EmployeesXmlName);

            if (!File.Exists(projectFile))
                WriteXml(projectFile, _projectSerializer,
                    new ProjectXmlModel { ProjectNumber = projectNumber });

            if (!File.Exists(employeesFile))
                WriteXml(employeesFile, _employeesSerializer, new List<EmployeeXmlModel>());
        }
        catch { }
    }

    // Writes (creates or overwrites) the project and employees XML files.
    public void Save(string projectFolder, ProjectXmlModel project, List<EmployeeXmlModel> employees)
    {
        var dir = Path.Combine(projectFolder, SaveAsPdfDir);
        if (!Directory.Exists(dir))
        {
            Directory.CreateDirectory(dir);
            SetHidden(dir);
        }
        WriteFiles(dir, project, employees);
    }

    private void WriteFiles(string dir, ProjectXmlModel project, List<EmployeeXmlModel> employees)
    {
        var projectPath   = Path.Combine(dir, ProjectXmlName);
        var employeesPath = Path.Combine(dir, EmployeesXmlName);

        WriteXml(projectPath, _projectSerializer, project);
        WriteXml(employeesPath, _employeesSerializer, employees);
    }

    private static void WriteXml(string path, XmlSerializer serializer, object obj)
    {
        // Strip Hidden/ReadOnly before overwriting so existing meta files can be updated.
        if (File.Exists(path))
            try { File.SetAttributes(path, FileAttributes.Normal); } catch { }

        using var stream = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None);
        serializer.Serialize(stream, obj);
        stream.Close();
        SetHidden(path);
    }

    private static void SetHidden(string path)
    {
        try { File.SetAttributes(path, File.GetAttributes(path) | FileAttributes.Hidden); } catch { }
    }
}

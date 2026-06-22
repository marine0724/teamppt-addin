using System.IO;

namespace TeampptAddin
{
    public interface IAccessPolicy
    {
        bool IsAdmin { get; }
        bool CanIngest { get; }
    }

    public class LocalFileAccessPolicy : IAccessPolicy
    {
        private readonly string _adminPath;
        public LocalFileAccessPolicy(string adminPath = null)
        {
            _adminPath = adminPath ?? AdminCredentials.DefaultPath;
        }
        public bool IsAdmin => File.Exists(_adminPath);
        public bool CanIngest => IsAdmin;
    }
}

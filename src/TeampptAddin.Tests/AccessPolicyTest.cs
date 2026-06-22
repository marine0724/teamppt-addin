using System.IO;
using Xunit;

namespace TeampptAddin.Tests
{
    public class AccessPolicyTest
    {
        [Fact]
        public void IsAdmin_True_When_File_Exists()
        {
            var tmp = Path.GetTempFileName();
            try
            {
                var p = new LocalFileAccessPolicy(tmp);
                Assert.True(p.IsAdmin);
                Assert.True(p.CanIngest);
            }
            finally { File.Delete(tmp); }
        }

        [Fact]
        public void IsAdmin_False_When_File_Missing()
        {
            var p = new LocalFileAccessPolicy(Path.Combine(Path.GetTempPath(), "no_such_admin_" + System.Guid.NewGuid() + ".json"));
            Assert.False(p.IsAdmin);
            Assert.False(p.CanIngest);
        }
    }
}

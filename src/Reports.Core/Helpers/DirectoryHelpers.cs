namespace Reports.Core.Helpers
{
    public static class DirectoryHelper
    {
        public static void CheckCreatePath(string path)
        {
            var dir = new DirectoryInfo(path);
            if (!dir.Exists)
            {
                Directory.CreateDirectory(path);
            }
        }

        public static void DeleteDirectory(string path)
        {
            if (Directory.Exists(path) && CheckFilePath(path))
            {
                try
                {
                    Directory.Delete(path, true);
                }
                catch (Exception)
                {
                    //ignore simultaneous deletes
                }
            }
        }

        public static bool CheckFilePath(string filePath)
        {
            if (string.IsNullOrEmpty(filePath))
                return false;

            filePath = new FileInfo(filePath).DirectoryName;

            return !string.IsNullOrEmpty(filePath)
                &&
                (
                    //filePath.StartsWith(AppSettings.ReportDirectory.TrimEndPath())
                    //||
                    //filePath.StartsWith(AppSettings.ImportDirectory.TrimEndPath())
                    //||
                    //filePath.StartsWith(AppSettings.ResourceDirectory.TrimEndPath())
                    //||
                    //filePath.StartsWith(AppSettings.SubmissionDirectory.TrimEndPath())
                    //||
                    filePath.StartsWith(AppDomain.CurrentDomain.BaseDirectory + "Content\\Downloads")
                 );
        }
    }


    public static class StringExtensions
    {

        public static string TrimEndPath(this string input)
        {
            if (!string.IsNullOrEmpty(input))
            {
                return input.Trim().TrimEnd('/', '\\');
            }

            return input;
        }


    }
}

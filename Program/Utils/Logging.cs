using System.Linq;
using System.Windows.Forms;
using NLog;
using NLog.Targets;

namespace Program.Utils;

public class Logging
{
    private static readonly Logger logger = LogManager.GetCurrentClassLogger();

    public static void LoggingTest()
    {
        // Check Logging File   
        var fileTarget = LogManager.Configuration.AllTargets.FirstOrDefault(t => t is FileTarget) as FileTarget;
        var filePath = fileTarget == null
            ? string.Empty
            : fileTarget.FileName.Render(new LogEventInfo { Level = LogLevel.Info });
        MessageBox.Show($"filePath: {filePath}");
    }
}
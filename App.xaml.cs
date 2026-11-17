using System.Windows;
using System.Windows.Media.Animation;

namespace SarpNote
{
    public partial class App : System.Windows.Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            Timeline.DesiredFrameRateProperty.OverrideMetadata(
                typeof(Timeline),
                new FrameworkPropertyMetadata { DefaultValue = 200 }
            );
            base.OnStartup(e);
        }
    }
}


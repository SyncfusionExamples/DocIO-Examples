using Foundation;
using Microsoft.Maui;
using Microsoft.Maui.Hosting;

namespace Create_Word_document
{
	[Register("AppDelegate")]
	public class AppDelegate : MauiUIApplicationDelegate
	{
		protected override MauiApp CreateMauiApp() => MauiProgram.CreateMauiApp();
	}
}
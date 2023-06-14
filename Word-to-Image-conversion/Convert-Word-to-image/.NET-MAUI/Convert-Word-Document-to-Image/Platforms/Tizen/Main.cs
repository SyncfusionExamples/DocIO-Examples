using System;
using Microsoft.Maui;
using Microsoft.Maui.Hosting;

namespace Convert_Word_Document_to_Image;

class Program : MauiApplication
{
	protected override MauiApp CreateMauiApp() => MauiProgram.CreateMauiApp();

	static void Main(string[] args)
	{
		var app = new Program();
		app.Run(args);
	}
}

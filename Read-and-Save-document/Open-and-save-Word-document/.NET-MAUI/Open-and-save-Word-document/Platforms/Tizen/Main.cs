using System;
using Microsoft.Maui;
using Microsoft.Maui.Hosting;

namespace Open_and_save_Word_document;

class Program : MauiApplication
{
	protected override MauiApp CreateMauiApp() => MauiProgram.CreateMauiApp();

	static void Main(string[] args)
	{
		var app = new Program();
		app.Run(args);
	}
}

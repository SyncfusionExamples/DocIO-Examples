Copy necessary fonts to Linux containers
----------------------------------------

The fonts present in the location(in Docker container) “/usr/local/share/fonts/” is used for conversion. By default, there will be limited number of fonts available in the container.

You should copy necessary fonts to this location “/usr/local/share/fonts/” before conversion.

How to run this example
-----------------------

*   Download this project to a location in your disk.
*   Create a folder "Fonts" parallelly to .sln file.
*   Download necessary Microsoft compatible fonts and paste into "Fonts" folder.
*   Open the solution file using Visual Studio.
*   Rebuild the solution to install the required NuGet packages.
*   Run the application.
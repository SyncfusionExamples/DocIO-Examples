Copy necessary fonts to Linux containers
----------------------------------------

In Word to PDF conversion, Essential DocIO uses the fonts which are installed in the corresponding production machine to measure and draw the text. If the font is not available in the production environment, then the alternate font will be used to measure and draw text based on the environment. And so, it is mandatory to install all the fonts used in the Word document in machine to achieve proper preservation.

How to run the examples
-----------------------

*   Download this project to a location in your disk.
*   Download necessary Microsoft compatible fonts and placed it in the fonts folder of this project.
*   Open the solution file using Visual Studio.
*   Rebuild the solution to install the required NuGet packages.
*   Run the application.
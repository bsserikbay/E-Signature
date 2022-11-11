using ESignature;

Console.WriteLine($"ESignature v.{System.Diagnostics.Process.GetCurrentProcess().MainModule.FileVersionInfo.ProductVersion}");
Console.Write("Enter path to your folder and press enter: ");

var targetDirectory = Console.ReadLine();
targetDirectory = targetDirectory.Trim('"');

if (targetDirectory != null && Directory.Exists(targetDirectory))
{
  
        var watch = System.Diagnostics.Stopwatch.StartNew();
        Console.WriteLine("Processing files ... ");

        var inspector = new SignatureInspector();
        var signatures = inspector.InspectSignatures(targetDirectory);
        var excelBuilder = new ExcelBuilder();
        excelBuilder.DisplayInExcel(signatures);

        watch.Stop();
        Console.WriteLine($"Execution Time: {watch.ElapsedMilliseconds} ms");
        Console.ReadLine();
    
   
}
else
{
    Console.WriteLine($"Check the path {targetDirectory}... Exiting ...");
    Console.ReadLine();
}

using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ESignature
{
    public class ExcelBuilder
    {
        public void DisplayInExcel(List<Signature> signatures)
        {
            var orderedSignatures = signatures.OrderBy(x => x.Signed);
            var excelApp = new Application
            {
                Visible = true
            };
            excelApp.Workbooks.Add();
            _Worksheet workSheet = (Worksheet)excelApp.ActiveSheet;

            workSheet.Cells[1, "A"] = "Is signed?";
            workSheet.Cells[1, "B"] = "File name";
            workSheet.Cells[1, "C"] = "Issued by";
            workSheet.Cells[1, "D"] = "Valid from";
            workSheet.Cells[1, "E"] = "Valid to";
            workSheet.Cells[1, "F"] = "Signature algorithm";
            workSheet.Cells[1, "G"] = "Subject";

            var row = 1;
            foreach (var signature in orderedSignatures)
            {
                row++;
                workSheet.Cells[row, "A"] = signature.Signed;
                workSheet.Cells[row, "B"] = signature.FileName;
                workSheet.Cells[row, "C"] = signature.Issuer;
                workSheet.Cells[row, "D"] = signature.NotBefore;
                workSheet.Cells[row, "E"] = signature.NotAfter;
                workSheet.Cells[row, "F"] = signature.AlgorithmName;
                workSheet.Cells[row, "G"] = signature.Subject;
            }

            workSheet.Columns.AutoFit();

            CalculateSignedFiles(signatures);
        }

        public void CalculateSignedFiles(List<Signature> signatures)
        {
            var signedFiles = new List<string>();
            var unsignedFiles = new List<string>();
            for (int i = 0; i < signatures.Count; i++)
            {
                var signature = signatures[i];
                if (signature.Signed == true)
                {
                    signedFiles.Add(signature.FileName);
                }
                else
                {
                    unsignedFiles.Add(signature.FileName);
                }
            }
            Console.WriteLine("...............");
            Console.WriteLine("Job completed");
            Console.WriteLine($"Finished at: {DateTime.Now}");
            Console.WriteLine($"Processed files: {signedFiles.Distinct().Count() + unsignedFiles.Distinct().Count()}");
            Console.WriteLine($"  -signed: {signedFiles.Distinct().Count()}");
            Console.WriteLine($"  -unsigned: {unsignedFiles.Distinct().Count()}");
        }
    }
}

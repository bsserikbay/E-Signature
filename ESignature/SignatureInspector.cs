using iText.Kernel.Pdf;
using iText.Signatures;
using Org.BouncyCastle.X509;
using System.Text.RegularExpressions;

namespace ESignature
{
    public class SignatureInspector
    {
        public List<Signature> InspectSignatures(string directory)
        {
            string[] fileEntries = Directory.GetFiles(directory);
            List<Signature> signatures = new();

            foreach (var fileName in fileEntries)
            {
                if (System.IO.Path.GetExtension(fileName).ToUpper().Contains("PDF"))
                {
                    try
                    {
                        PdfDocument pdfDoc = new PdfDocument(new PdfReader(fileName));
                        SignatureUtil signUtil = new SignatureUtil(pdfDoc);


                        try
                        {
                            IList<string> names = signUtil.GetSignatureNames();


                            if (names.Count == 0)
                            {
                                signatures.Add(new Signature
                                {
                                    FileName = System.IO.Path.GetFileName(fileName),
                                    Signed = false
                                });
                            }
                            else
                            {
                                foreach (string name in names)
                                {


                                    PdfPKCS7 pkcs7 = signUtil.ReadSignatureData(name);
                                    X509Certificate[] certs = pkcs7.GetSignCertificateChain();

                                    for (int i = 0; i < certs.Length; i++)
                                    {
                                        X509Certificate cert = (X509Certificate)certs[i];

                                        if (i == 0) // we are retrieving only root certificates
                                        {
                                            signatures.Add(new Signature
                                            {
                                                FileName = System.IO.Path.GetFileName(fileName),
                                                SignatureName = i.ToString(),
                                                NotBefore = cert.NotBefore.ToUniversalTime().ToString("yyyy-MM-dd"),
                                                NotAfter = cert.NotAfter.ToUniversalTime().ToString("yyyy-MM-dd"),
                                                AlgorithmName = cert.SigAlgName,
                                                Version = cert.Version,
                                                Subject = Regex.Replace(cert.SubjectDN.ToString(), @"(IIN[0-9]{12})", "IIN***********"),
                                                Issuer = cert.IssuerDN.ToString(),
                                                Signed = true
                                            });
                                        }
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {

                            Console.WriteLine("Error: "+ ex.Message + "Corrupt file: " + fileName);
                        }
                       
                    }
                    catch (Exception ex)
                    {

                        Console.WriteLine("Error message: " + ex.Message + "Corrupt file: " + fileName);
                    }
                }

            }
            return signatures;
        }
    }
}

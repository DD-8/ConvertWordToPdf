using Oracle.ManagedDataAccess.Client;
using System.Diagnostics;

class Program
{
    static void Main(string[] args)
    {
        Console.WriteLine("Word dosyalarını PDF'e dönüştürme programına hoşgeldiniz.");
        bool isProgramRunning = true;
        while (isProgramRunning)
        {
            Console.Write("Test modu için 0'a, normal mod için 1'e basınız: ");
            string mode = Console.ReadLine();
            while (mode != "0" && mode != "1")
            {
                Console.Write("Yanlış tuşlama yaptınız. Lütfen tekrar deneyin: ");
                mode = Console.ReadLine();
            }

            Console.Write(@"Word dosyalarının bulunduğu klasör yolunu giriniz(Örnek Yol: C:\Users\user\Desktop\Files\Words): ");
            string folderPath = @Console.ReadLine();
            folderPath = IsDirectoryExists(folderPath);

            Console.Write(@"PDF dosyalarının kaydedileceği klasör yolunu giriniz(Örnek Yol: C:\Users\user\Desktop\Files\PDFs): ");
            string outputFolderPath = Console.ReadLine();
            outputFolderPath = IsDirectoryExists(outputFolderPath);

            Console.Write(@"Libre Office yolunu giriniz(Örnek Yol: C:\Program Files\LibreOffice\program\soffice.exe): ");
            string libreOfficePath = Console.ReadLine();
            libreOfficePath = IsFileExists(libreOfficePath);

            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();
            string[] wordFiles = Directory.GetFiles(folderPath, "*.doc*", SearchOption.AllDirectories);
            if (mode == "0")
            {
                TestMode(wordFiles, outputFolderPath, libreOfficePath);
            }
            else
            {
                string connectionString = "ConnectionString";
                RealMode(wordFiles, connectionString, outputFolderPath, libreOfficePath);
            }

            stopwatch.Stop();
            Console.WriteLine("İşlem tamamlandı. Toplam İşlenen Dosya Sayısı: " + wordFiles.Length + " Toplam süre: " + stopwatch.Elapsed.TotalSeconds.ToString("F2") + " saniye.");
            Console.Write("Yeniden işlem yapmak için lütfen 1'e basınız. Çıkmak için bunun dışında herhangi bir tuşa basabilirsiniz.");
            string input = Console.ReadLine();
            if (input == "1")
                isProgramRunning = true;
            else
                isProgramRunning = false;
        }
    }

    static string IsDirectoryExists(string path)
    {
        while (!Directory.Exists(path))
        {
            Console.Write("Belirtilen klasör yolu bulunamadı. Lütfen geçerli bir klasör yolu giriniz: ");
            path = Console.ReadLine();
        }
        return path;
    }

    static string IsFileExists(string path)
    {
        while (!File.Exists(path))
        {
            Console.Write("Belirtilen dosya bulunamadı. Lütfen geçerli bir dosya yolu giriniz: ");
            path = Console.ReadLine();
        }
        return path;
    }

    static void TestMode(string[] wordFiles, string outputFolderPath, string libreOfficePath)
    {
        foreach (string wordFile in wordFiles)
        {
            ConvertToPdf(wordFile, outputFolderPath, libreOfficePath);
        }
    }

    static string ConvertToPdf(string wordFile, string outputFolderPath, string libreOfficePath)
    {
        string pdfFile = Path.ChangeExtension(wordFile, ".pdf");
        pdfFile = Path.Combine(outputFolderPath, Path.GetFileName(pdfFile));

        ProcessStartInfo startInfo = new ProcessStartInfo();
        startInfo.FileName = libreOfficePath;
        startInfo.Arguments = $"--headless --convert-to pdf \"{wordFile}\" --outdir \"{outputFolderPath}\"";

        Process process = new Process();
        process.StartInfo = startInfo;
        process.Start();
        process.WaitForExit();
        return pdfFile;
    }

    static void RealMode(string[] wordFiles, string connectionString, string outputFolderPath, string libreOfficePath)
    {
        foreach (string wordFile in wordFiles)
        {
            string sampleNumber = Path.GetFileNameWithoutExtension(wordFile).Split('_')[0];
            if (IsThereARecord(connectionString, sampleNumber))
            {
                string pdfPath = ConvertToPdf(wordFile, outputFolderPath, libreOfficePath);
                UpdateTable(pdfPath, connectionString);
                DeletePdfFile(pdfPath);
            }
        }
    }

    static void DeletePdfFile(string pdfFilePath)
    {
        File.Delete(pdfFilePath);
    }

    static bool IsThereARecord(string connectionString, string fileName)
    {
        int count = 0;
        using (OracleConnection connection = new OracleConnection(connectionString))
        {
            connection.Open();
            using (OracleCommand command = connection.CreateCommand())
            {
                command.Parameters.Add("SAMPLE", OracleDbType.Varchar2).Value = fileName;
                command.CommandText = "SELECT COUNT(*) FROM TABLE WHERE SAMPLE=:SAMPLE AND PDF IS NULL";
                count = Convert.ToInt32(command.ExecuteScalar());
            }
            connection.Close();
        }
        if (count == 0)
            return false;
        else
            return true;
    }

    static void UpdateTable(string pdfFile, string connectionString)
    {
        byte[] pdfData = File.ReadAllBytes(pdfFile);
        string sampleNumber = Path.GetFileNameWithoutExtension(pdfFile).Split('_')[0];
        using (OracleConnection connection = new OracleConnection(connectionString))
        {
            connection.Open();
            using (OracleCommand command = connection.CreateCommand())
            {
                command.Parameters.Add("PDF", OracleDbType.Blob).Value = pdfData;
                command.Parameters.Add("SAMPLE", OracleDbType.Varchar2).Value = sampleNumber;
                command.CommandText = "UPDATE TABLE SET PDF=:PDF WHERE SAMPLE=:SAMPLE";
                int rowsUpdated = command.ExecuteNonQuery();
                Console.WriteLine($"{sampleNumber} - PDF güncellendi:");
            }
            connection.Close();
        }
    }
}


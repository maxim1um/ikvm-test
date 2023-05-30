namespace ikvm_test
{
    internal class Program
    {
        static void Main(string[] args)
        {
            java.util.Map data = new java.util.HashMap();
            data.put("name", "sayi");

            com.deepoove.poi.XWPFTemplate template = com.deepoove.poi.XWPFTemplate.compile("test.docx").render(data);

            var outputPath = "output.docx";

            if (File.Exists(outputPath))
            {
                if (IsFileLocked(outputPath))
                {
                    Console.WriteLine("Test Failed");
                    return;
                }

                File.Delete(outputPath);
            }

            template.writeAndClose(new java.io.FileOutputStream(outputPath));

            Console.WriteLine("Test Done");
        }

        public static bool IsFileLocked(string filePath)
        {
            try
            {
                using (FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                {
                    stream.Close();
                }
            }
            catch (IOException)
            {
                return true;
            }

            return false;
        }

    }
}
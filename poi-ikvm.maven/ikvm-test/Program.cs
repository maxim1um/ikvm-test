namespace ikvm_test
{
    using org.apache.poi.xwpf.usermodel;
    internal class Program
    {
        static void Main(string[] args)
        {
            org.apache.logging.log4j.Logger log4j2 = org.apache.logging.log4j.LogManager.getLogger();

            var outputPath = "output.docx";

            XWPFDocument doc = new XWPFDocument();
            XWPFParagraph p1 = doc.createParagraph();
            XWPFRun r1 = p1.createRun();
            r1.setText("The quick brown fox");

            if (File.Exists(outputPath))
            {
                if (IsFileLocked(outputPath))
                {
                    Console.WriteLine("Test Failed");
                    return;
                }

                File.Delete(outputPath);
            }

            java.io.FileOutputStream os = new java.io.FileOutputStream(outputPath);
            doc.write(os);
            doc.close();

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
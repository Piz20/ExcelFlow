public class GenerationRequest
{
    public string filePath { get; set; } = string.Empty;
    public string templatePath { get; set; } = string.Empty;
    public string outputDir { get; set; } = string.Empty;
    public string sheetName { get; set; } = "Analyse";
    public int startIndex { get; set; } = 0;
    public int count { get; set; } = 200;

    public int currentIndex { get; set; }

}

namespace Budget_Forms_Core
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string currentYear = DateTime.Now.ToString("yyyy");
            string workbookFileName = $"{currentYear}.xlsx";
            var result = WorkbookWrite.CreateBlank(workbookFileName);
            //instead of looping through every week with new popups, add 
            if (result.Success)
            {
                Console.WriteLine(result.Message);
            }
            else
            {
                Console.WriteLine($"Failed to create workbook: {result.Message}");
            }
        }
    }
}

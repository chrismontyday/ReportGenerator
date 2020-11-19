namespace UzonePageObject
{
    public class TeamMember
    {
        public string Name { get; set; }
        public string Email { get; set; }
        public string Event { get; set; }
        public string TeamName { get; set; }
        public string SubTeamName { get; set; }
        public string PhotoPath { get; set; }
        public int Id { get; set; }
        public string PhotoFilePath { get; set; }
        public string MapFilePath { get; set; }

        public static string AddDummyPhoto()
        {
            return @"C:\Code\testautomation-ui-uzone-page-object-model\Test\Testdata\shawn (2).jpg";
            //return null;
        }


        public static string AddDummyMap()
        {
            //return @"C:\Code\testautomation-ui-uzone-page-object-model\Test\Testdata\Screenshotsaria ying2020-10-30 14-49-43north building 1st floor.jpg";
            return null;
        }
    }
}

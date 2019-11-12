namespace HAA.BusinessEntities.Input
{
    public class UpdateSpkrConfigModel
    {
        //ByRef ProjNumb As String, ByRef SelectedSpkrConfigur As String _
        //, ByRef ConfigTypeDolby As String, ByRef Seatss As Object, ByRef SpkrInfo As Object

        public string ProjNumb { get; set; }

        public string SelectedSpkrConfigur { get; set; }

        public string ConfigTypeDolby { get; set; }

        public object Seatss { get; set; }

        public object SpkrInfo { get; set; }
    }
}
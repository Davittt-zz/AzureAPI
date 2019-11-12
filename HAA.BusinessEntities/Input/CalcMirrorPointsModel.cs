using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HAA.BusinessEntities.Input
{
    public class CalcMirrorPointsModel
    {
        //ByRef MainLP As Object, ByRef WallArr As Object, ByRef SpkrImage As Object _
        //, ByRef SpkrConfig As String, ByRef TopWidth As Object, ByRef SpeakerCnt As Integer, ByRef SpeakerInfo As Object

        public object MainLP { get; set; }

        public object WallArr { get; set; }

        public object SpkrImage { get; set; }

        public string SpkrConfig { get; set; }

        public object TopWidth { get; set; }

        public int SpeakerCnt { get; set; }

        public object SpeakerInfo { get; set; }
    }
}
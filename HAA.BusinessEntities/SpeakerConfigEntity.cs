namespace HAA.BusinessEntities
{
    public class SpeakerConfigEntity
    {
        public int? DolbyIndex { get; set; }

        public string DolbyName { get; set; }

        public string ConfigBrand { get; set; }

        public int? SubIndex { get; set; }

        public string UnitsV { get; set; }

        public bool? OptStandardDolby { get; set; }

        public bool? OptImax { get; set; }

        public bool? OptTrinnov { get; set; }
    }
}
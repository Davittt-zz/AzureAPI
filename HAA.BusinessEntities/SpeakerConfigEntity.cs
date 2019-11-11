namespace HAA.BusinessEntities
{
    public class SpeakerConfigEntity
    {
        public int Index { get; set; }
        public string DrawRefIndex { get; set; }
        public string Brand { get; set; }
        public string Model { get; set; }
        public string Version { get; set; }
        public string RetailPrice { get; set; }
        public decimal? Height { get; set; }
        public decimal? Width { get; set; }
        public decimal? Depth { get; set; }
        public decimal? AcOffsetX { get; set; }
        public decimal? AcOffsetY { get; set; }
        public decimal? AcOffsetZ { get; set; }
        public decimal? WooferOffsetX { get; set; }
        public decimal? WooferOffsetY { get; set; }
        public decimal? WooferOffsetZ { get; set; }
        public string TopImage { get; set; }
        public string SideImage { get; set; }
        public string FrontImage { get; set; }
        public string BackImage { get; set; }
        public string FreqResponse { get; set; }
        public string FreqResponsePlot { get; set; }
        public string OffFreqResponsePlot { get; set; }
        public string HorCoverage { get; set; }
        public string VerCoverage { get; set; }
        public string NominalImpedance { get; set; }
        public string Sensitivity { get; set; }
        public string MaxSoundLevel { get; set; }
        public string Crossover { get; set; }
        public string Mounting { get; set; }
        public string Type { get; set; }
        public string Slope { get; set; }
        public decimal? X { get; set; }
        public decimal? Y { get; set; }
        public decimal? Z { get; set; }
        public string Status { get; set; }
        public string Designation { get; set; }
        public string Wall { get; set; }
        public string Family { get; set; }

        public ProjectEntity Project { get; set; }

        public string ProjectNumber { get; set; }
    }
}
namespace HAA.DataModel
{
    public class SpeakerConfig
    {
        public int Index { get; set; }
        public int? DrawRefIndex { get; set; }
        public string Brand { get; set; }
        public string Model { get; set; }
        public string Version { get; set; }
        public string RetailPrice { get; set; }
        public double? Height { get; set; }
        public double? Width { get; set; }
        public double? Depth { get; set; }
        public double? AcOffsetX { get; set; }
        public double? AcOffsetY { get; set; }
        public double? AcOffsetZ { get; set; }
        public double? WooferOffsetX { get; set; }
        public double? WooferOffsetY { get; set; }
        public double? WooferOffsetZ { get; set; }
        public string TopImage { get; set; }
        public string SideImage { get; set; }
        public string FrontImage { get; set; }
        public string BackImage { get; set; }
        public string FreqResponse { get; set; }
        public string FreqResponsePlot { get; set; }
        public string OffFreqResponsePlot { get; set; }
        public double? HorCoverage { get; set; }
        public double? VerCoverage { get; set; }
        public double? NominalImpedance { get; set; }
        public double? Sensitivity { get; set; }
        public double? MaxSoundLevel { get; set; }
        public double? Crossover { get; set; }
        public string Mounting { get; set; }
        public string Type { get; set; }
        public string Slope { get; set; }
        public double? X { get; set; }
        public double? Y { get; set; }
        public double? Z { get; set; }
        public int? Status { get; set; }
        public string Designation { get; set; }
        public string Wall { get; set; }
        public string Family { get; set; }
        public Project Project { get; set; }

        public int ProjectId { get; set; }
    }
}

namespace ElectricalToolSuite.FindAndReplace
{
    public class MatchingParameterDto
    {
        public string Name { get; set; }
        public string Value { get; set; }

        public MatchingParameterDto(string name, string value)
        {
            Name = name;
            Value = value;
        }
    }
}

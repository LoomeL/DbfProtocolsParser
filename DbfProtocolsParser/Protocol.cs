namespace DbfProtocolsParser;

public struct Protocol
{
    public string protocol_number { get; set; }
    public string type { get; set; }
    public string factory_number { get; set; }
    public string manufacturer { get; set; }
    public string high_year { get; set; }
    public string user { get; set; }
    public string verification_conditions { get; set; }
    public string date_of_check { get; set; }
    public string verifier { get; set; }
    public string label_number { get; set; }
}
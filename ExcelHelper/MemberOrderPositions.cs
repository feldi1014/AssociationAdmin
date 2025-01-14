public class MemberOrderPositions : IEquatable<MemberOrderPositions>
{
    public MemberOrderPositions(string bestellDatum, string mitgliedsNummer) 
    {
        Bestelldatum = bestellDatum;    
        Mitgliedsnummer = mitgliedsNummer;
    }


    public string Bestelldatum { get; }
    public string Mitgliedsnummer { get; }
    public string Vorname { get; set; } = string.Empty;
    public string Nachname { get; set; } = string.Empty;
    public int BestellPositionen { get; set; }

    public bool Equals(MemberOrderPositions? other)
    {
        if (other == null) return false;
        if (ReferenceEquals(this, other)) return true;
        return Bestelldatum.Equals(other.Bestelldatum) && Mitgliedsnummer.Equals(other.Mitgliedsnummer);
    }

    public override string ToString()
    {
        return $"{Bestelldatum} {Mitgliedsnummer} {Vorname} {Nachname} ({BestellPositionen})";
    }
}
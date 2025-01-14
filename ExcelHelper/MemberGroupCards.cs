
public class MemberGroupCards : IEquatable<MemberGroupCards>
{
    public MemberGroupCards(string groupMember, string card)

    {
        GroupMember = groupMember;
        Card = card;
        Count = 1;
    }
    public string GroupMember { get; }
    public string Card { get; }

    public int Count { get; set; }

    public bool Equals(MemberGroupCards? other)
    {
        if (other == null) return false;
        if (other == this) return true;
        return GroupMember == other.GroupMember && Card == other.Card;
    }

    public override string ToString()
    {
        return $"{GroupMember} {Card} count: {Count}";
    }
}

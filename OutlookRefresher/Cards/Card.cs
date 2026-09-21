using System;

namespace OutlookRefresher.Cards
{
    public sealed class Card : IEquatable<Card>, IComparable<Card>
    {
        public Rank Rank { get; }
        public Suit Suit { get; }

        public Card(Rank rank, Suit suit)
        {
            Rank = rank;
            Suit = suit;
        }

        public override string ToString()
        {
            return $"{Rank} of {Suit}";
        }

        public bool Equals(Card other)
        {
            if (other == null) return false;
            return Rank == other.Rank && Suit == other.Suit;
        }

        public override bool Equals(object obj)
        {
            return Equals(obj as Card);
        }

        public override int GetHashCode()
        {
            return ((int)Rank * 397) ^ (int)Suit;
        }

        public int CompareTo(Card other)
        {
            if (other == null) return 1;
            return ((int)Rank).CompareTo((int)other.Rank);
        }

        public static bool operator ==(Card left, Card right)
        {
            if (ReferenceEquals(left, right)) return true;
            if (left is null || right is null) return false;
            return left.Equals(right);
        }

        public static bool operator !=(Card left, Card right)
        {
            return !(left == right);
        }
    }
}

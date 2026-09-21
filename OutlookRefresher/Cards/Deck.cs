using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace OutlookRefresher.Cards
{
    public sealed class Deck : IEnumerable<Card>
    {
        private readonly List<Card> _cards;
        private readonly Random _random;

        public Deck()
            : this(new Random())
        {
        }

        public Deck(Random random)
        {
            _random = random ?? throw new ArgumentNullException(nameof(random));
            _cards = new List<Card>(52);
            Reset();
        }

        public int Count => _cards.Count;

        public void Reset()
        {
            _cards.Clear();
            foreach (Suit suit in Enum.GetValues(typeof(Suit)))
            {
                foreach (Rank rank in Enum.GetValues(typeof(Rank)))
                {
                    _cards.Add(new Card(rank, suit));
                }
            }
        }

        public void Shuffle()
        {
            for (int i = _cards.Count - 1; i > 0; i--)
            {
                int j = _random.Next(i + 1);
                Card temp = _cards[i];
                _cards[i] = _cards[j];
                _cards[j] = temp;
            }
        }

        public Card Draw()
        {
            if (_cards.Count == 0)
                throw new InvalidOperationException("Cannot draw from an empty deck.");

            int lastIndex = _cards.Count - 1;
            Card card = _cards[lastIndex];
            _cards.RemoveAt(lastIndex);
            return card;
        }

        public IList<Card> Draw(int count)
        {
            if (count < 0)
                throw new ArgumentOutOfRangeException(nameof(count));
            if (count > _cards.Count)
                throw new InvalidOperationException("Not enough cards remaining in the deck.");

            var drawn = new List<Card>(count);
            for (int i = 0; i < count; i++)
            {
                drawn.Add(Draw());
            }
            return drawn;
        }

        public IEnumerator<Card> GetEnumerator()
        {
            return _cards.ToList().GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}

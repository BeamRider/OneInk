using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace LibOneInk
{
    public struct OneInkColor
    {
        public byte r;
        public byte g;
        public byte b;
        public byte a;
    }

    public struct OneInkPoint
    {
        public float x;
        public float y;
        public float pressure;
    }

    public class OneInkStrokeGroup
    {
        public List<OneInkStroke> Strokes { get; } = new List<OneInkStroke>();
    }

    public class OneInkStroke
    {
        public OneInkPointsCollection Points { get; }
        public OneInkColor Color { get; set; } = new OneInkColor { r = 0, g = 0, b = 0, a = 0xff };

        public OneInkStroke()
        {
            Points = new OneInkPointsCollection();
        }

        public OneInkStroke(int sz)
        {
            Points = new OneInkPointsCollection(sz);
        }
    }

    public class OneInkPointsCollection : ICollection<OneInkPoint>
    {
        private readonly ArrayList _points;

        public OneInkPointsCollection()
        {
            _points = new ArrayList();
        }

        public OneInkPointsCollection(int sz)
        {
            _points = new ArrayList(sz);
        }

        public int Count => _points.Count;

        public bool IsReadOnly => _points.IsReadOnly;

        public void Add(OneInkPoint item) => _points.Add(item);

        public void Clear() => _points.Clear();

        public bool Contains(OneInkPoint item) => _points.Contains(item);

        public void CopyTo(OneInkPoint[] array, int arrayIndex) => _points.CopyTo(array, arrayIndex);

        public IEnumerator<OneInkPoint> GetEnumerator() => _points.Cast<OneInkPoint>().GetEnumerator();

        public bool Remove(OneInkPoint item)
        {
            int idx = _points.IndexOf(item);
            if (idx < 0)
                return false;
            _points.RemoveAt(idx);
            return true;
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _points.GetEnumerator();
        }
    }
}

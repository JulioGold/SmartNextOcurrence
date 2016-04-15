using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;
using System;

namespace SmartNextOcurrence
{
    internal class TextSelection
    {
        private readonly IWpfTextView _view;
        private Tuple<ITrackingPoint, ITrackingPoint> _selection;
        private int _start;
        private int _end;
        private int _lastPosition;

        public int Start { get { return _selection.Item1.GetPosition(_view.TextSnapshot); } }

        public int End { get { return _selection.Item2.GetPosition(_view.TextSnapshot); } }

        public TextSelection(Tuple<ITrackingPoint, ITrackingPoint> selection, IWpfTextView view)
        {
            _selection = selection;
            _view = view;
            _start = _selection.Item1.GetPosition(_view.TextSnapshot);
            _end = _selection.Item2.GetPosition(_view.TextSnapshot);
            _lastPosition = _end;
        }

        public void Move(int position)
        {
            if (_lastPosition > position)
            {
                if (position > _start)
                {
                    _end = position;
                }
                else if (position < _start)
                {
                    _start = position;
                }
            }
            else if (_lastPosition < position)
            {
                if (position > _start && _end > _lastPosition)
                {
                    _start = position;
                }
                else if (position > _start)
                {
                    _end = position;
                }
                else if (position < _start)
                {
                    _start = position;
                }
            }

            _selection = new Tuple<ITrackingPoint, ITrackingPoint>
                (
                    _view.TextSnapshot.CreateTrackingPoint(new SnapshotPoint(_view.TextSnapshot, _start), PointTrackingMode.Positive),
                    _view.TextSnapshot.CreateTrackingPoint(new SnapshotPoint(_view.TextSnapshot, _end), PointTrackingMode.Positive)
                );

            _lastPosition = position;
        }
    }
}

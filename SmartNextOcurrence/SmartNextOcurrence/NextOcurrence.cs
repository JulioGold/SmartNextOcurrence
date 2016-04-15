using System;
using System.Windows.Controls;
using System.Windows.Media;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Text.Formatting;
using System.Windows.Input;
using System.Windows;
using System.Text;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Windows.Shapes;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio;

namespace SmartNextOcurrence
{
    internal sealed class NextOcurrence
    {
        private readonly IOleCommandTarget _nextTarget; // Usado para propagar os eventos nos demais locais onde conter um cursor
        private readonly IAdornmentLayer _layer;
        private readonly IWpfTextView _view;
        private List<ITrackingPoint> _trackPointList = new List<ITrackingPoint>();
        private List<Tuple<ITrackingPoint, ITrackingPoint>> _selectedTrackPointList = new List<Tuple<ITrackingPoint, ITrackingPoint>>();
        private int _lastIndex = 0;

        public bool Editing { get; set; } = false;

        // Indicates if selecting
        public bool Selecting { get; set; } = false;

        public NextOcurrence(IWpfTextView view, IOleCommandTarget nextTarget)
        {
            if (view == null)
            {
                throw new ArgumentNullException("view");
            }

            _layer = view.GetAdornmentLayer("NextOcurrence");

            _view = view;

            _nextTarget = nextTarget;

            // Quando qualquer coisa mudar na tela, dispara o evento
            _view.LayoutChanged += this.OnLayoutChanged;
        }

        void OnLayoutChanged(object sender, TextViewLayoutChangedEventArgs e)
        {
            RedrawScreen();
        }

        private void DrawSinglePoint(ITrackingPoint trackingPoint)
        {
            if (trackingPoint.GetPosition(_view.TextSnapshot) >= _view.TextSnapshot.Length)
            {
                return;
            }

            SnapshotSpan span = new SnapshotSpan(trackingPoint.GetPoint(_view.TextSnapshot), 1);

            SolidColorBrush brush = Brushes.Black;

            Geometry geometry = _view.TextViewLines.GetLineMarkerGeometry(span);

            GeometryDrawing drawing = new GeometryDrawing(brush, null, geometry);

            if (drawing.Bounds.IsEmpty)
            {
                return;
            }

            Rectangle rectangle = new Rectangle()
            {
                Fill = brush,
                Width = drawing.Bounds.Width / 6,
                //Height = drawing.Bounds.Height - 4,
                Height = drawing.Bounds.Height - 2,
                Margin = new System.Windows.Thickness(0, 2, 0, 0),
            };

            Canvas.SetLeft(rectangle, geometry.Bounds.Left);
            Canvas.SetTop(rectangle, geometry.Bounds.Top);

            _layer.AddAdornment(AdornmentPositioningBehavior.TextRelative, span, "NextOcurrence", rectangle, null);
        }

        private void DrawSingleSelection(Tuple<ITrackingPoint, ITrackingPoint> trackingPoint)
        {
            SnapshotSpan span = new SnapshotSpan(_view.TextSnapshot, Span.FromBounds(trackingPoint.Item1.GetPoint(_view.TextSnapshot), trackingPoint.Item2.GetPoint(_view.TextSnapshot)));

            Geometry geometry = _view.TextViewLines.GetMarkerGeometry(span);

            if (geometry != null)
            {
                GeometryDrawing drawing = new GeometryDrawing(new SolidColorBrush(Color.FromRgb(0xa5, 0xcf, 0xf3)), new Pen(), geometry);
                drawing.Freeze();

                DrawingImage drawingImage = new DrawingImage(drawing);
                drawingImage.Freeze();

                Image image = new Image
                {
                    Source = drawingImage,
                };

                // Align the image with the top of the bounds of the text geometry
                Canvas.SetLeft(image, geometry.Bounds.Left);
                Canvas.SetTop(image, geometry.Bounds.Top);

                // Adiciona o adornment, mas ao fazer o RedrawScreen ele é perdido
                _layer.AddAdornment(AdornmentPositioningBehavior.TextRelative, span, null, image, null);
            }
        }

        internal void RedrawScreen()
        {
            _layer.RemoveAllAdornments();

            // Draw the selections
            foreach (var item in _selectedTrackPointList)
            {
                DrawSingleSelection(item);
            }

            // Draw the cursors
            foreach (ITrackingPoint trackingPoint in _trackPointList)
            {
                DrawSinglePoint(trackingPoint);
            }
        }

        /// <summary>
        /// When hit Ctrl+D
        /// </summary>
        internal void SelectNextOcurrence()
        {
            string selectedText = _view.Selection.StreamSelectionSpan.GetText();

            // Coloca o cursor na própria palavra já selecionada
            _lastIndex = (_lastIndex == 0) ? _view.Selection.ActivePoint.Position.Position - selectedText.Length : _lastIndex;

            // Expressão para buscar todas as palavras que forem iguais ao texto selecionado
            Regex todoLineRegex = new Regex(selectedText + @"\b");

            // Executa a expressão buscando as palavras
            MatchCollection matches = todoLineRegex.Matches(_view.TextViewLines.FormattedSpan.GetText());

            // Se encontrou algo igual ao selecionado
            if (matches.Count > 0)
            {
                // Último match que foi encontrado
                Match ultimoMatch = matches[matches.Count - 1];

                // Próximo match depois do último selecionado
                Match match = todoLineRegex.Match(_view.TextViewLines.FormattedSpan.GetText(), _lastIndex);

                // Se selecionou a próxima e é depois da última
                if (match.Success && (_trackPointList.Count < matches.Count))
                {
                    if (match.Index == ultimoMatch.Index)
                    {
                        _lastIndex = 1;
                    }
                    else
                    {
                        _lastIndex = match.Index + match.Length;
                    }

                    SnapshotSpan span = new SnapshotSpan(_view.TextSnapshot, Span.FromBounds(match.Index, match.Index + match.Length));

                    Geometry geometry = _view.TextViewLines.GetMarkerGeometry(span);

                    if (geometry != null)
                    {
                        Brush brush = new SolidColorBrush(Color.FromArgb(0x20, 0x00, 0x00, 0xff));
                        brush.Freeze();

                        SolidColorBrush penBrush = new SolidColorBrush(Colors.Red);
                        penBrush.Freeze();

                        Pen pen = new Pen(penBrush, 0.5);
                        pen.Freeze();

                        GeometryDrawing drawing = new GeometryDrawing(brush, pen, geometry);
                        drawing.Freeze();

                        DrawingImage drawingImage = new DrawingImage(drawing);
                        drawingImage.Freeze();

                        Image image = new Image
                        {
                            Source = drawingImage,
                        };

                        // Align the image with the top of the bounds of the text geometry
                        Canvas.SetLeft(image, geometry.Bounds.Left);
                        Canvas.SetTop(image, geometry.Bounds.Top);

                        // Adiciona o adornment, mas ao fazer o RedrawScreen ele é perdido
                        _layer.AddAdornment(AdornmentPositioningBehavior.TextRelative, span, null, image, null);

                        // Cria o novo ponto do cursor, neste caso estou colocando ele na frente da palavra, ou seja na direita da palavra
                        ITrackingPoint cursorTrackingPoint = _view.TextSnapshot.CreateTrackingPoint(new SnapshotPoint(_view.TextSnapshot, match.Index + match.Length), PointTrackingMode.Positive);

                        _trackPointList.Add(cursorTrackingPoint);

                        // Adiciono uma seleção
                        _selectedTrackPointList.Add(
                            new Tuple<ITrackingPoint, ITrackingPoint>
                            (
                                _view.TextSnapshot.CreateTrackingPoint(new SnapshotPoint(_view.TextSnapshot, match.Index), PointTrackingMode.Positive),
                                _view.TextSnapshot.CreateTrackingPoint(new SnapshotPoint(_view.TextSnapshot, match.Index + match.Length), PointTrackingMode.Positive)
                            )
                        );

                        RedrawScreen();
                    }
                }

                Selecting = true;

                // Diz que está editando
                Editing = true;
            }
        }

        internal void SelectPreviousCharacter()
        {
            _view.Selection.Clear();
            _view.Caret.IsHidden = true;

            // Faz com que o "cursor" vá para o início da palavra, pois aqui é SelectPreviousWord, ou seja a palavra anterior
            for (int i = 0; i < _trackPointList.Count; i++)
            {
                ITrackingPoint item = _trackPointList[i];

                int start = (item.GetPosition(_view.TextSnapshot) == 0) ? 0 : (item.GetPosition(_view.TextSnapshot) - 1);
                int end = item.GetPosition(_view.TextSnapshot);

                // Se o tamanho da minha lista de seleção é menor que a lista de trackingpoint quer dizer que esta é a primeira ver e que não tinha nada selecionado ainda
                if (_selectedTrackPointList.Count < _trackPointList.Count)
                {
                    // Adiciono uma seleção
                    _selectedTrackPointList.Add(
                        new Tuple<ITrackingPoint, ITrackingPoint>
                        (
                            _view.TextSnapshot.CreateTrackingPoint(new SnapshotPoint(_view.TextSnapshot, start), PointTrackingMode.Positive),
                            _view.TextSnapshot.CreateTrackingPoint(new SnapshotPoint(_view.TextSnapshot, end), PointTrackingMode.Positive)
                        )
                    );
                }
                else // Se já tem algo selecionado
                {
                    // Pega o item da seleção mas este item tem o mesmo índice do item de tracking point
                    Tuple<ITrackingPoint, ITrackingPoint> selectedItem = _selectedTrackPointList[i];

                    // Crio uma tupla mas atualizo apenas o início da seleção, o final da seleção permanece o mesmo
                    _selectedTrackPointList[i] =
                        new Tuple<ITrackingPoint, ITrackingPoint>
                        (
                            _view.TextSnapshot.CreateTrackingPoint(new SnapshotPoint(_view.TextSnapshot, start), PointTrackingMode.Positive),
                            selectedItem.Item2
                        );
                }

                int newCursorPosition = (_trackPointList[i].GetPosition(_view.TextSnapshot) == 0) ? 0 : (_trackPointList[i].GetPosition(_view.TextSnapshot) - 1);

                _trackPointList[i] = _view.TextSnapshot.CreateTrackingPoint(new SnapshotPoint(_view.TextSnapshot, newCursorPosition), PointTrackingMode.Positive);
            }

            Selecting = true;

            RedrawScreen();

            // Diz que está editando
            Editing = true;
        }

        internal void SelectNextCharacter()
        {
            // TODO: Implement the behavior here

            _view.Selection.Clear();
            _view.Caret.IsHidden = true;

            //// Faz com que o "cursor" vá para o início da palavra, pois aqui é SelectPreviousWord, ou seja a palavra anterior
            for (int i = 0; i < _trackPointList.Count; i++)
            {
                ITrackingPoint item = _trackPointList[i];

                int textLength = _view.TextViewLines.FormattedSpan.GetText().Length;
                int start = item.GetPosition(_view.TextSnapshot); 
                int end = ((item.GetPosition(_view.TextSnapshot) + 1) > textLength) ? textLength : (item.GetPosition(_view.TextSnapshot) + 1);

                // Se o tamanho da minha lista de seleção é menor que a lista de trackingpoint quer dizer que esta é a primeira ver e que não tinha nada selecionado ainda
                if (_selectedTrackPointList.Count < _trackPointList.Count)
                {
                    // Adiciono uma seleção
                    _selectedTrackPointList.Add(
                        new Tuple<ITrackingPoint, ITrackingPoint>
                        (
                            _view.TextSnapshot.CreateTrackingPoint(new SnapshotPoint(_view.TextSnapshot, start), PointTrackingMode.Positive),
                            _view.TextSnapshot.CreateTrackingPoint(new SnapshotPoint(_view.TextSnapshot, end), PointTrackingMode.Positive)
                        )
                    );
                }
                else // Se já tem algo selecionado
                {
                    // Pega o item da seleção mas este item tem o mesmo índice do item de tracking point
                    Tuple<ITrackingPoint, ITrackingPoint> selectedItem = _selectedTrackPointList[i];

                    // Crio uma tupla mas atualizo apenas o início da seleção, o final da seleção permanece o mesmo
                    _selectedTrackPointList[i] =
                        new Tuple<ITrackingPoint, ITrackingPoint>
                        (
                            selectedItem.Item1,
                            _view.TextSnapshot.CreateTrackingPoint(new SnapshotPoint(_view.TextSnapshot, end), PointTrackingMode.Positive)
                        );
                }

                int newCursorPosition = ((item.GetPosition(_view.TextSnapshot) + 1) > textLength) ? textLength : (item.GetPosition(_view.TextSnapshot) + 1);

                _trackPointList[i] = _view.TextSnapshot.CreateTrackingPoint(new SnapshotPoint(_view.TextSnapshot, newCursorPosition), PointTrackingMode.Positive);
            }

            Selecting = true;

            RedrawScreen();

            // Diz que está editando
            Editing = true;
        }

        internal void SelectPreviousWord()
        {
            _view.Selection.Clear();
            _view.Caret.IsHidden = true;

            // Faz com que o "cursor" vá para o início da palavra, pois aqui é SelectPreviousWord, ou seja a palavra anterior
            for (int i = 0; i < _trackPointList.Count; i++)
            {
                ITrackingPoint item = _trackPointList[i];

                WordPosition previousWord = WordPosition.PreviousWord(_view.TextViewLines.FormattedSpan.GetText(), item.GetPosition(_view.TextSnapshot));

                // Pega o item da seleção mas este item tem o mesmo índice do item de tracking point
                Tuple<ITrackingPoint, ITrackingPoint> selectedItem = _selectedTrackPointList[i];

                int startPosition = previousWord.Start < selectedItem.Item1.GetPosition(_view.TextSnapshot) ? previousWord.Start : selectedItem.Item1.GetPosition(_view.TextSnapshot);
                int endPosition = previousWord.End < selectedItem.Item2.GetPosition(_view.TextSnapshot) ? previousWord.End : selectedItem.Item2.GetPosition(_view.TextSnapshot);

                // Crio uma tupla mas atualizo apenas o início da seleção, o final da seleção permanece o mesmo
                _selectedTrackPointList[i] =
                    new Tuple<ITrackingPoint, ITrackingPoint>
                    (
                        _view.TextSnapshot.CreateTrackingPoint(new SnapshotPoint(_view.TextSnapshot, startPosition), PointTrackingMode.Positive),
                        _view.TextSnapshot.CreateTrackingPoint(new SnapshotPoint(_view.TextSnapshot, endPosition), PointTrackingMode.Positive)
                    );

                int newCursorPosition = ((_trackPointList[i].GetPosition(_view.TextSnapshot) - previousWord.Word.Length) >= 0) && (previousWord.Word.Length > 0) ? (_trackPointList[i].GetPosition(_view.TextSnapshot) - previousWord.Word.Length) : 0;

                _trackPointList[i] = _view.TextSnapshot.CreateTrackingPoint(new SnapshotPoint(_view.TextSnapshot, newCursorPosition), PointTrackingMode.Positive);
            }

            Selecting = true;

            RedrawScreen();

            // Diz que está editando
            Editing = true;
        }

        internal void SelectNextWord()
        {
            _view.Selection.Clear();
            _view.Caret.IsHidden = true;

            // Faz com que o "cursor" vá para o final da palavra, pois aqui é SelectNextWord, ou seja a próxima palavra
            for (int i = 0; i < _trackPointList.Count; i++)
            {
                ITrackingPoint item = _trackPointList[i];

                WordPosition nextWord = WordPosition.NextWord(_view.TextViewLines.FormattedSpan.GetText(), item.GetPosition(_view.TextSnapshot));

                // Pega o item da seleção mas este item tem o mesmo índice do item de tracking point
                Tuple<ITrackingPoint, ITrackingPoint> selectedItem = _selectedTrackPointList[i];

                int startPosition = nextWord.End < selectedItem.Item2.GetPosition(_view.TextSnapshot) ? nextWord.End : selectedItem.Item1.GetPosition(_view.TextSnapshot);
                int endPosition = nextWord.End > selectedItem.Item2.GetPosition(_view.TextSnapshot) ? nextWord.End : selectedItem.Item2.GetPosition(_view.TextSnapshot);

                // Crio uma tupla mas atualizo apenas o início da seleção, o final da seleção permanece o mesmo
                _selectedTrackPointList[i] =
                    new Tuple<ITrackingPoint, ITrackingPoint>
                    (
                        _view.TextSnapshot.CreateTrackingPoint(new SnapshotPoint(_view.TextSnapshot, startPosition), PointTrackingMode.Positive),
                        //selectedItem.Item1,
                        _view.TextSnapshot.CreateTrackingPoint(new SnapshotPoint(_view.TextSnapshot, endPosition), PointTrackingMode.Positive)
                    //_view.TextSnapshot.CreateTrackingPoint(new SnapshotPoint(_view.TextSnapshot, nextWord.End), PointTrackingMode.Positive)
                    );

                int newCursorPosition = ((_trackPointList[i].GetPosition(_view.TextSnapshot) + nextWord.Word.Length) < _view.TextViewLines.FormattedSpan.GetText().Length) && (nextWord.Word.Length > 0) ?
                    (_trackPointList[i].GetPosition(_view.TextSnapshot) + nextWord.Word.Length) :
                    _view.TextViewLines.FormattedSpan.GetText().Length - 1;

                _trackPointList[i] = _view.TextSnapshot.CreateTrackingPoint(new SnapshotPoint(_view.TextSnapshot, newCursorPosition), PointTrackingMode.Positive);
            }

            Selecting = true;

            RedrawScreen();

            // Diz que está editando
            Editing = true;
        }

        private void DeleteSelection()
        {
            foreach (Tuple<ITrackingPoint, ITrackingPoint> item in _selectedTrackPointList)
            {
                // Delete the selected text
                ITextEdit edit = _view.TextBuffer.CreateEdit();
                ITextSnapshot snapshot = edit.Snapshot;
                int caracteresToDelete = item.Item2.GetPosition(snapshot) - item.Item1.GetPosition(snapshot);
                edit.Delete(item.Item1.GetPosition(snapshot), caracteresToDelete);
                edit.Apply();
            }
        }

        internal void CopySelection()
        {
            StringBuilder sb = new StringBuilder();

            foreach (Tuple<ITrackingPoint, ITrackingPoint> item in _selectedTrackPointList)
            {
                int quant = item.Item2.GetPosition(_view.TextSnapshot) - item.Item1.GetPosition(_view.TextSnapshot);
                sb.Append(_view.TextViewLines.FormattedSpan.GetText().Substring(item.Item1.GetPosition(_view.TextSnapshot), quant));
            }

            string content = sb.ToString();

            if (!String.IsNullOrEmpty(content))
            {
                Clipboard.SetText(content);
            }
        }

        internal int SyncedOperation(ref Guid pguidCmdGroup, uint nCmdID, uint nCmdexecopt, IntPtr pvaIn, IntPtr pvaOut)
        {
            int result = 0;

            if (_trackPointList.Count > 0)
            {
                // Se tem algo selecionado
                if (Selecting)
                {
                    switch (nCmdID)
                    {
                        case ((uint)VSConstants.VSStd2KCmdID.TYPECHAR):
                        case ((uint)VSConstants.VSStd2KCmdID.BACKSPACE):
                        case ((uint)VSConstants.VSStd2KCmdID.TAB):
                        case ((uint)VSConstants.VSStd97CmdID.Delete):
                        case ((uint)VSConstants.VSStd2KCmdID.RETURN):
                        case ((uint)VSConstants.VSStd2KCmdID.BACKTAB):
                        case ((uint)VSConstants.VSStd97CmdID.Paste): /* Ctrl+V */

                            DeleteSelection();

                            break;

                        default:
                            break;
                    }
                }

                //char typedChar = char.MinValue;
                //typedChar = (char)(ushort)Marshal.GetObjectForNativeVariant(pvaIn);

                _view.Selection.Clear();

                ITextCaret caret = _view.Caret;

                List<ITrackingPoint> tempTrackList = _trackPointList;

                _trackPointList = new List<ITrackingPoint>();

                SnapshotPoint snapshotPoint = tempTrackList[0].GetPoint(_view.TextSnapshot);

                //m_dte.UndoContext.Open("Multi-point edit");

                for (int i = 0; i < tempTrackList.Count; i++)
                {
                    snapshotPoint = tempTrackList[i].GetPoint(_view.TextSnapshot);

                    caret.MoveTo(snapshotPoint);

                    // Propaga o evento para os demais locais
                    result = _nextTarget.Exec(ref pguidCmdGroup, nCmdID, nCmdexecopt, pvaIn, pvaOut);

                    AddTrackingPoint(_view.Caret.Position);
                }

                //m_dte.UndoContext.Close();

                RedrawScreen();
            }

            return result;
        }

        internal void ClearTrackingPoints()
        {
            _lastIndex = 0;
            // Diz que parou de editar
            Editing = false;
            Selecting = false;
            _trackPointList.Clear();
            _layer.RemoveAllAdornments();
            _selectedTrackPointList.Clear();
        }

        private void AddTrackingPoint(CaretPosition caretPosition)
        {
            ITrackingPoint trackingPoint = _view.TextSnapshot.CreateTrackingPoint(caretPosition.BufferPosition.Position, PointTrackingMode.Positive);

            if (trackingPoint.GetPosition(_view.TextSnapshot) >= 0)
            {
                _trackPointList.Add(trackingPoint);
            }
            else
            {
                trackingPoint = _view.TextSnapshot.CreateTrackingPoint(0, PointTrackingMode.Positive);
                _trackPointList.Add(trackingPoint);
            }

            if (caretPosition.VirtualSpaces > 0)
            {
                _view.Caret.MoveTo(trackingPoint.GetPoint(_view.TextSnapshot));
            }
        }

        internal void CancelarEdicao()
        {
            ClearTrackingPoints();
            RedrawScreen();
        }

        internal void CancelSelecting()
        {
            Selecting = false;
            _selectedTrackPointList.Clear();
            _view.Caret.IsHidden = false;
            //_layer.RemoveMatchingAdornments();
            RedrawScreen();
        }

        internal void SplitIntoLines()
        {
            // TODO: Get selected lines, hide selection and create an cursor/caret/trackpoint at the end of line
        }
    }
}

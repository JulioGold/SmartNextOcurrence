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
        private List<TextSelection> _selectionList = new List<TextSelection>();
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
                _layer.AddAdornment(AdornmentPositioningBehavior.TextRelative, span, "NextOcurrence", image, null);
            }
        }

        internal void RedrawScreen()
        {
            _layer.RemoveAllAdornments();

            // Draw the selections
            foreach (var item in _selectionList)
            {
                DrawSingleSelection(
                    new Tuple<ITrackingPoint, ITrackingPoint>
                    (
                        _view.TextSnapshot.CreateTrackingPoint(new SnapshotPoint(_view.TextSnapshot, item.Start), PointTrackingMode.Positive),
                        _view.TextSnapshot.CreateTrackingPoint(new SnapshotPoint(_view.TextSnapshot, item.End), PointTrackingMode.Positive)
                    ));
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
            string selectedText;
            
            // Se já tem algo selecionado
            if (_selectionList.Count <= 0 && !_view.Selection.IsEmpty)
            {
                selectedText = _view.Selection.StreamSelectionSpan.GetText();

                // Coloca o cursor na própria palavra já selecionada
                _lastIndex = (_lastIndex == 0) ? _view.Selection.ActivePoint.Position.Position - selectedText.Length : _lastIndex;
            }
            else if (_selectionList.Count > 0)
            {
                var itemSelecionado = _selectionList[0];

                selectedText = _view.TextViewLines.FormattedSpan.GetText().Substring(itemSelecionado.Start, itemSelecionado.End - itemSelecionado.Start);

                // Coloca o cursor na própria palavra já selecionada
                _lastIndex = (_lastIndex == 0) ? itemSelecionado.End - selectedText.Length : _lastIndex;
            }
            else
            {
                // Se não tinha nada selecionado então retorno pois agora o filtro é este, o comando sempre vem para cá
                return;
            }

            _view.Selection.IsActive = false;
            _view.Selection.Clear();
            _view.Caret.IsHidden = true;

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

                        _selectionList.Add(new TextSelection(
                            new Tuple<ITrackingPoint, ITrackingPoint>
                            (
                                _view.TextSnapshot.CreateTrackingPoint(new SnapshotPoint(_view.TextSnapshot, match.Index), PointTrackingMode.Positive),
                                _view.TextSnapshot.CreateTrackingPoint(new SnapshotPoint(_view.TextSnapshot, match.Index + match.Length), PointTrackingMode.Positive)
                            ),
                            _view));

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
            _view.Selection.IsActive = false;
            _view.Caret.IsHidden = true;
            _view.Selection.Clear();

            // Faz com que o "cursor" vá para o início da palavra, pois aqui é SelectPreviousWord, ou seja a palavra anterior
            for (int i = 0; i < _trackPointList.Count; i++)
            {
                ITrackingPoint cursorPos = _trackPointList[i];

                int newCursorPosition = (_trackPointList[i].GetPosition(_view.TextSnapshot) == 0) ? 0 : (_trackPointList[i].GetPosition(_view.TextSnapshot) - 1);

                _trackPointList[i] = _view.TextSnapshot.CreateTrackingPoint(new SnapshotPoint(_view.TextSnapshot, newCursorPosition), PointTrackingMode.Positive);

                if (_selectionList.Count < _trackPointList.Count)
                {
                    _selectionList.Add(new TextSelection(
                            new Tuple<ITrackingPoint, ITrackingPoint>
                            (
                                _view.TextSnapshot.CreateTrackingPoint(new SnapshotPoint(_view.TextSnapshot, newCursorPosition), PointTrackingMode.Positive),
                                _view.TextSnapshot.CreateTrackingPoint(new SnapshotPoint(_view.TextSnapshot, cursorPos.GetPosition(_view.TextSnapshot)), PointTrackingMode.Positive)
                            ),
                            _view));
                }
                else
                {
                    _selectionList[i].Move(newCursorPosition);
                }
            }

            Selecting = true;

            RedrawScreen();

            // Diz que está editando
            Editing = true;
        }

        internal void SelectNextCharacter()
        {
            _view.Selection.IsActive = false;
            _view.Caret.IsHidden = true;
            _view.Selection.Clear();

            // Faz com que o "cursor" vá para o início da palavra, pois aqui é SelectPreviousWord, ou seja a palavra anterior
            for (int i = 0; i < _trackPointList.Count; i++)
            {
                ITrackingPoint cursorPos = _trackPointList[i];

                int textLength = _view.TextViewLines.FormattedSpan.GetText().Length;

                int newCursorPosition = ((cursorPos.GetPosition(_view.TextSnapshot) + 1) > textLength) ? textLength : (cursorPos.GetPosition(_view.TextSnapshot) + 1);

                _trackPointList[i] = _view.TextSnapshot.CreateTrackingPoint(new SnapshotPoint(_view.TextSnapshot, newCursorPosition), PointTrackingMode.Positive);
                
                if (_selectionList.Count < _trackPointList.Count)
                {
                    _selectionList.Add(new TextSelection(
                            new Tuple<ITrackingPoint, ITrackingPoint>
                            (
                                _view.TextSnapshot.CreateTrackingPoint(new SnapshotPoint(_view.TextSnapshot, cursorPos.GetPosition(_view.TextSnapshot)), PointTrackingMode.Positive),
                                _view.TextSnapshot.CreateTrackingPoint(new SnapshotPoint(_view.TextSnapshot, newCursorPosition), PointTrackingMode.Positive)
                            ),
                            _view));
                }
                else
                {
                    _selectionList[i].Move(newCursorPosition);
                }
            }

            Selecting = true;

            RedrawScreen();

            // Diz que está editando
            Editing = true;
        }

        internal void SelectPreviousWord()
        {
            _view.Selection.IsActive = false;
            _view.Caret.IsHidden = true;
            _view.Selection.Clear();

            // Faz com que o "cursor" vá para o início da palavra, pois aqui é SelectPreviousWord, ou seja a palavra anterior
            for (int i = 0; i < _trackPointList.Count; i++)
            {
                ITrackingPoint cursorPos = _trackPointList[i];

                WordPosition previousWord = WordPosition.PreviousWord(_view.TextViewLines.FormattedSpan.GetText(), cursorPos.GetPosition(_view.TextSnapshot));

                int newCursorPosition = ((_trackPointList[i].GetPosition(_view.TextSnapshot) - previousWord.Word.Length) >= 0) && (previousWord.Word.Length > 0) ? (_trackPointList[i].GetPosition(_view.TextSnapshot) - previousWord.Word.Length) : 0;

                _trackPointList[i] = _view.TextSnapshot.CreateTrackingPoint(new SnapshotPoint(_view.TextSnapshot, newCursorPosition), PointTrackingMode.Positive);

                if (_selectionList.Count < _trackPointList.Count)
                {
                    _selectionList.Add(new TextSelection(
                            new Tuple<ITrackingPoint, ITrackingPoint>
                            (
                                _view.TextSnapshot.CreateTrackingPoint(new SnapshotPoint(_view.TextSnapshot, previousWord.Start), PointTrackingMode.Positive),
                                _view.TextSnapshot.CreateTrackingPoint(new SnapshotPoint(_view.TextSnapshot, previousWord.End), PointTrackingMode.Positive)
                            ),
                            _view));
                }
                else
                {
                    _selectionList[i].Move(newCursorPosition);
                }
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
                ITrackingPoint cursorPos = _trackPointList[i];

                WordPosition nextWord = WordPosition.NextWord(_view.TextViewLines.FormattedSpan.GetText(), cursorPos.GetPosition(_view.TextSnapshot));

                int newCursorPosition = ((_trackPointList[i].GetPosition(_view.TextSnapshot) + nextWord.Word.Length) < _view.TextViewLines.FormattedSpan.GetText().Length) && (nextWord.Word.Length > 0) ?
                    (_trackPointList[i].GetPosition(_view.TextSnapshot) + nextWord.Word.Length) :
                    _view.TextViewLines.FormattedSpan.GetText().Length - 1;

                _trackPointList[i] = _view.TextSnapshot.CreateTrackingPoint(new SnapshotPoint(_view.TextSnapshot, newCursorPosition), PointTrackingMode.Positive);

                if (_selectionList.Count < _trackPointList.Count)
                {
                    _selectionList.Add(new TextSelection(
                            new Tuple<ITrackingPoint, ITrackingPoint>
                            (
                                _view.TextSnapshot.CreateTrackingPoint(new SnapshotPoint(_view.TextSnapshot, nextWord.Start), PointTrackingMode.Positive),
                                _view.TextSnapshot.CreateTrackingPoint(new SnapshotPoint(_view.TextSnapshot, nextWord.End), PointTrackingMode.Positive)
                            ),
                            _view));
                }
                else
                {
                    _selectionList[i].Move(newCursorPosition);
                }
            }

            Selecting = true;

            RedrawScreen();

            // Diz que está editando
            Editing = true;
        }

        private void DeleteSelection()
        {
            foreach (var item in _selectionList)
            {
                // Delete the selected text
                ITextEdit edit = _view.TextBuffer.CreateEdit();
                ITextSnapshot snapshot = edit.Snapshot;
                int caracteresToDelete = item.End - item.Start;
                edit.Delete(item.Start, caracteresToDelete);
                edit.Apply();
            }

            CancelSelecting();
        }

        internal void CopySelection()
        {
            StringBuilder sb = new StringBuilder();

            foreach (var item in _selectionList)
            {
                int quant = item.End - item.Start;
                sb.Append(_view.TextViewLines.FormattedSpan.GetText().Substring(item.Start, quant));
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
            _selectionList.Clear();
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
            _selectionList.Clear();
            _view.Caret.IsHidden = false;
            _view.Selection.IsActive = true;
            RedrawScreen();
        }

        internal void SplitIntoLines()
        {
            // TODO: Get selected lines, hide selection and create an cursor/caret/trackpoint at the end of line
        }

        internal void HandleClick()
        {
            AddTrackingPoint(_view.Caret.Position);

            // Diz que está editando
            Editing = true;

            RedrawScreen();

            /*
            if (addCursor && _view.Selection.SelectedSpans.All(span => span.Length == 0))
            {
                if (_view.Selection.SelectedSpans.Count == 1)
                {
                    if (m_trackList.Count == 0)
                    {
                        AddSyncPoint(lastCaretPosition);
                    }

                    AddSyncPoint(_view.Caret.Position);
                    RedrawScreen();
                }
                else
                {
                    foreach (var span in _view.Selection.SelectedSpans)
                    {
                        AddSyncPoint(span.Start.Position);
                    }

                    _view.Selection.Clear();
                    RedrawScreen();
                }
            }
            else if (m_trackList.Any())
            {
                ClearSyncPoints();
                RedrawScreen();
            }

            lastCaretPosition = _view.Caret.Position;
            */
        }
    }
}

using Microsoft.VisualStudio;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Text.Editor;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;

namespace SmartNextOcurrence
{
    public class NextOcurrenceCommandFilter : IOleCommandTarget
    {
        // Indica que já foi adicionado o filtro de comandos isso para não adicionar novamente.
        internal bool IsAdded { get; set; }

        private NextOcurrence _nextOcurrence;

        private NextOcurrence NextOcurrence
        {
            get
            {
                if (_nextOcurrence == null)
                {
                    return _nextOcurrence = new NextOcurrence(_textView, _nextTarget);
                }

                return _nextOcurrence;
            }
            set { }
        }

        internal IOleCommandTarget _nextTarget;

        private IWpfTextView _textView;

        public NextOcurrenceCommandFilter(IWpfTextView textView)
        {
            _textView = textView;
        }

        public int Exec(ref Guid pguidCmdGroup, uint nCmdID, uint nCmdexecopt, IntPtr pvaIn, IntPtr pvaOut)
        {
            if (pguidCmdGroup == typeof(VSConstants.VSStd2KCmdID).GUID)
            {
                switch (nCmdID)
                {
                    case ((uint)VSConstants.VSStd2KCmdID.UP):
                    case ((uint)VSConstants.VSStd2KCmdID.DOWN):
                    case ((uint)VSConstants.VSStd2KCmdID.LEFT):
                    case ((uint)VSConstants.VSStd2KCmdID.RIGHT):

                        /* Se estava selecionando algo, cancela a seleção. */
                        if (NextOcurrence.Selecting)
                        {
                            NextOcurrence.CancelSelecting();
                        }

                        if (NextOcurrence.Editing)
                        {
                            return NextOcurrence.SyncedOperation(ref pguidCmdGroup, nCmdID, nCmdexecopt, pvaIn, pvaOut);
                        }

                        break;

                    case ((uint)VSConstants.VSStd2KCmdID.TYPECHAR):
                    case ((uint)VSConstants.VSStd2KCmdID.BACKSPACE):
                    case ((uint)VSConstants.VSStd2KCmdID.DELETEWORDRIGHT):
                    case ((uint)VSConstants.VSStd2KCmdID.DELETEWORDLEFT):
                    case ((uint)VSConstants.VSStd2KCmdID.TAB):
                    case ((uint)VSConstants.VSStd2KCmdID.END):
                    case ((uint)VSConstants.VSStd2KCmdID.HOME):
                    case ((uint)VSConstants.VSStd2KCmdID.PAGEDN):
                    case ((uint)VSConstants.VSStd2KCmdID.PAGEUP):
                    case ((uint)VSConstants.VSStd2KCmdID.PASTE):
                    case ((uint)VSConstants.VSStd2KCmdID.PASTEASHTML):
                    case ((uint)VSConstants.VSStd2KCmdID.BOL):
                    case ((uint)VSConstants.VSStd2KCmdID.EOL):
                    case ((uint)VSConstants.VSStd2KCmdID.RETURN):
                    case ((uint)VSConstants.VSStd2KCmdID.BACKTAB):
                    case ((uint)VSConstants.VSStd2KCmdID.WORDPREV):
                    case ((uint)VSConstants.VSStd2KCmdID.WORDNEXT):

                        if (NextOcurrence.Editing)
                        {
                            return NextOcurrence.SyncedOperation(ref pguidCmdGroup, nCmdID, nCmdexecopt, pvaIn, pvaOut);
                        }

                        break;

                    /* Seleciona uma letra para a esquerda */
                    case ((uint)VSConstants.VSStd2KCmdID.LEFT_EXT):
                        
                        NextOcurrence.SelectPreviousCharacter();

                        break;

                    /* Seleciona uma letra para a direita */
                    case ((uint)VSConstants.VSStd2KCmdID.RIGHT_EXT):

                        NextOcurrence.SelectNextCharacter();

                        break;

                    /* Seleciona uma palavra para a esquerda */
                    case ((uint)VSConstants.VSStd2KCmdID.WORDPREV_EXT):

                        NextOcurrence.SelectPreviousWord();

                        break;

                    /* Seleciona uma palavra para a direita */
                    case ((uint)VSConstants.VSStd2KCmdID.WORDNEXT_EXT):

                        NextOcurrence.SelectNextWord();

                        break;

                    /* Ctrl+Shift+L */
                    case ((uint)VSConstants.VSStd2KCmdID.DELETELINE):

                        NextOcurrence.SplitIntoLines();
                        
                        break;

                    /* Quando tecla Esc cancela as operações */
                    case ((uint)VSConstants.VSStd2KCmdID.CANCEL):

                        NextOcurrence.CancelarEdicao();

                        break;

                    default:
                        break;
                }
            }

            if (pguidCmdGroup == typeof(VSConstants.VSStd97CmdID).GUID)
            {
                switch (nCmdID)
                {
                    /* Ctrl+D */
                    case ((uint)VSConstants.VSStd97CmdID.SearchCombo):

                        NextOcurrence.SelectNextOcurrence();

                        break;

                    /* Quando tecla Esc cancela as operações */
                    case ((uint)VSConstants.VSStd97CmdID.Cancel):

                        NextOcurrence.CancelarEdicao();

                        break;

                    /* Ctrl+C */
                    case ((uint)VSConstants.VSStd97CmdID.Copy):

                        /* Se estava selecionando algo, cancela a seleção. */
                        if (NextOcurrence.Selecting)
                        {
                            NextOcurrence.CopySelection();

                            // Stop event to don't propagate action of Copy
                            return 0;
                        }

                        break;

                    /* Ctrl+V */
                    case ((uint)VSConstants.VSStd97CmdID.Paste):

                        /* Se estava selecionando algo, cancela a seleção. */
                        if (NextOcurrence.Selecting)
                        {
                            return NextOcurrence.SyncedOperation(ref pguidCmdGroup, nCmdID, nCmdexecopt, pvaIn, pvaOut);
                        }

                        break;

                    default:
                        break;
                }
            }

            #region Desativado

            //if (Keyboard.IsKeyDown(Key.LeftCtrl) && Keyboard.IsKeyDown(Key.LeftShift) && Keyboard.IsKeyDown(Key.Left))
            //{
            //    //// Se tem algo selecionado
            //    //if (!_textView.Selection.IsEmpty)
            //    //{
            //    //NextOcurrence.SelectNextOcurrence();
            //    //}
            //    if (NextOcurrence.Editing)
            //    {
            //        return NextOcurrence.SyncedOperation(ref pguidCmdGroup, nCmdID, nCmdexecopt, pvaIn, pvaOut);
            //    }
            //}

            //if (Keyboard.IsKeyDown(Key.LeftCtrl) && Keyboard.IsKeyDown(Key.LeftShift) && Keyboard.IsKeyDown(Key.Right))
            //{
            //    //// Se tem algo selecionado
            //    //if (!_textView.Selection.IsEmpty)
            //    //{
            //    //NextOcurrence.SelectNextOcurrence();
            //    //}
            //    if (NextOcurrence.Editing)
            //    {
            //        return NextOcurrence.SyncedOperation(ref pguidCmdGroup, nCmdID, nCmdexecopt, pvaIn, pvaOut);
            //    }
            //}

            //if (Keyboard.IsKeyDown(Key.LeftCtrl) && Keyboard.IsKeyDown(Key.D))
            //{
            //    // Se tem algo selecionado
            //    if (!_textView.Selection.IsEmpty)
            //    {
            //        NextOcurrence.SelectNextOcurrence();
            //    }
            //}

            #endregion

            return _nextTarget.Exec(ref pguidCmdGroup, nCmdID, nCmdexecopt, pvaIn, pvaOut);
        }

        public int QueryStatus(ref Guid pguidCmdGroup, uint cCmds, OLECMD[] prgCmds, IntPtr pCmdText)
        {
            return _nextTarget.QueryStatus(ref pguidCmdGroup, cCmds, prgCmds, pCmdText);
        }

        public void HandleClick(bool addCursor)
        {
            if (addCursor)
            {
                NextOcurrence.HandleClick();
            }
            else
            {
                NextOcurrence.CancelSelecting();
                NextOcurrence.CancelarEdicao();
            }
        }
    }
}

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
    class NextOcurrenceMouseProcessor : IMouseProcessor
    {
        private readonly IWpfTextView _wpfTextView;

        public NextOcurrenceMouseProcessor(IWpfTextView wpfTextView)
        {
            _wpfTextView = wpfTextView;
        }

        public void PostprocessMouseLeftButtonUp(MouseButtonEventArgs e)
        {
            // Nas propriedades do do editor vou buscar o command filter
            var commandFilter = _wpfTextView.Properties.GetProperty<NextOcurrenceCommandFilter>(typeof(NextOcurrenceCommandFilter));

            if (commandFilter != null)
            {
                commandFilter.HandleClick();
            }
        }

        #region Unused

        public void PostprocessDragEnter(DragEventArgs e)
        {
        }

        public void PostprocessDragLeave(DragEventArgs e)
        {
        }

        public void PostprocessDragOver(DragEventArgs e)
        {
        }

        public void PostprocessDrop(DragEventArgs e)
        {
        }

        public void PostprocessGiveFeedback(GiveFeedbackEventArgs e)
        {
        }

        public void PostprocessMouseDown(MouseButtonEventArgs e)
        {
        }

        public void PostprocessMouseEnter(MouseEventArgs e)
        {
        }

        public void PostprocessMouseLeave(MouseEventArgs e)
        {
        }

        public void PostprocessMouseLeftButtonDown(MouseButtonEventArgs e)
        {
        }

        public void PostprocessMouseMove(MouseEventArgs e)
        {
        }

        public void PostprocessMouseRightButtonDown(MouseButtonEventArgs e)
        {
        }

        public void PostprocessMouseRightButtonUp(MouseButtonEventArgs e)
        {
        }

        public void PostprocessMouseUp(MouseButtonEventArgs e)
        {
        }

        public void PostprocessMouseWheel(MouseWheelEventArgs e)
        {
        }

        public void PostprocessQueryContinueDrag(QueryContinueDragEventArgs e)
        {
        }

        public void PreprocessDragEnter(DragEventArgs e)
        {
        }

        public void PreprocessDragLeave(DragEventArgs e)
        {
        }

        public void PreprocessDragOver(DragEventArgs e)
        {
        }

        public void PreprocessDrop(DragEventArgs e)
        {
        }

        public void PreprocessGiveFeedback(GiveFeedbackEventArgs e)
        {
        }

        public void PreprocessMouseDown(MouseButtonEventArgs e)
        {
        }

        public void PreprocessMouseEnter(MouseEventArgs e)
        {
        }

        public void PreprocessMouseLeave(MouseEventArgs e)
        {
        }

        public void PreprocessMouseLeftButtonDown(MouseButtonEventArgs e)
        {
        }

        public void PreprocessMouseLeftButtonUp(MouseButtonEventArgs e)
        {
        }

        public void PreprocessMouseMove(MouseEventArgs e)
        {
        }

        public void PreprocessMouseRightButtonDown(MouseButtonEventArgs e)
        {
        }

        public void PreprocessMouseRightButtonUp(MouseButtonEventArgs e)
        {
        }

        public void PreprocessMouseUp(MouseButtonEventArgs e)
        {
        }

        public void PreprocessMouseWheel(MouseWheelEventArgs e)
        {
        }

        public void PreprocessQueryContinueDrag(QueryContinueDragEventArgs e)
        {
        }

        #endregion
    }
}

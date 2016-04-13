using System.ComponentModel.Composition;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Utilities;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Editor;
using Microsoft.VisualStudio.TextManager.Interop;
using Microsoft.VisualStudio;

namespace SmartNextOcurrence
{
    [Export(typeof(IWpfTextViewCreationListener))]
    [ContentType("text")]
    [TextViewRole(PredefinedTextViewRoles.Document)]
    internal sealed class NextOcurrenceTextViewCreationListener : IWpfTextViewCreationListener
    {
        [Import(typeof(IVsEditorAdaptersFactoryService))]
        internal IVsEditorAdaptersFactoryService editorFactory = null;

        // Disable "Field is never assigned to..." and "Field is never used" compiler's warnings. Justification: the field is used by MEF.
#pragma warning disable 649, 169

        /// <summary>
        /// Defines the adornment layer for the adornment. This layer is ordered
        /// after the selection layer in the Z-order
        /// </summary>
        [Export(typeof(AdornmentLayerDefinition))]
        [Name("NextOcurrence")]
        [Order(After = PredefinedAdornmentLayers.Selection, Before = PredefinedAdornmentLayers.Text)]
        private AdornmentLayerDefinition editorAdornmentLayer;

#pragma warning restore 649, 169

        #region IWpfTextViewCreationListener

        public void TextViewCreated(IWpfTextView textView)
        {
            AddCommandFilter(textView, new NextOcurrenceCommandFilter(textView));
        }

        private void AddCommandFilter(IWpfTextView textView, NextOcurrenceCommandFilter commandFilter)
        {
            // Se ainda não foi adicionado
            if (commandFilter.IsAdded == false)
            {
                IOleCommandTarget next;
                IVsTextView view = editorFactory.GetViewAdapter(textView);

                int hr = view.AddCommandFilter(commandFilter, out next);

                if (hr == VSConstants.S_OK)
                {
                    commandFilter.IsAdded = true;
                    textView.Properties.AddProperty(typeof(NextOcurrenceCommandFilter), commandFilter);

                    if (next != null)
                    {
                        commandFilter._nextTarget = next;
                    }
                }
            }
        }
        #endregion
    }
}

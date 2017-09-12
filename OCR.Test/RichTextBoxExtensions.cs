using System.Drawing;
using System.Windows.Forms;

namespace OCR.Test
{
    public static class RichTextBoxExtensions
    {
        public static void AppendText(this RichTextBox box, string text, Color color)
        {
            box.SelectionStart = box.TextLength;
            box.SelectionLength = 0;

            box.SelectionColor = color;
            box.AppendText(text);
            box.SelectionColor = box.ForeColor;
        }

        public static void AddText(this RichTextBox box, string text, Color color)
        {
            box.SelectionColor = color;
            box.Text = text;
            box.SelectionColor = box.ForeColor;
        }
    }
}

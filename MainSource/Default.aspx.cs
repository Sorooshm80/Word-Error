using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class _Default : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        Microsoft.Office.Interop.Word.Application objWord = new Microsoft.Office.Interop.Word.Application();
        Microsoft.Office.Interop.Word.Document objDoc = new Microsoft.Office.Interop.Word.Document();

        objWord.Options.ArabicNumeral = Microsoft.Office.Interop.Word.WdArabicNumeral.wdNumeralContext;
        objDoc.Application.Options.ArabicNumeral = Microsoft.Office.Interop.Word.WdArabicNumeral.wdNumeralContext;
        Microsoft.Office.Interop.Word.Range _range = objDoc.Range(Type.Missing, Type.Missing);

    }
}
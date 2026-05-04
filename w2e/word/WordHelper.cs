using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace w2e.word
{
    public static class WordHelper
    {
        public static (int? numId, int? level) GetNumberingInfo( Paragraph a_pars, StyleDefinitionsPart a_stylePart )
        {
            NumberingProperties numPr = a_pars.ParagraphProperties?.NumberingProperties;
            if( null != numPr )
            {
                return (
                    (int?)numPr.NumberingId?.Val?.Value,
                    (int?)numPr.NumberingLevelReference?.Val?.Value
                );
            }

            string styleId = a_pars.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            if( null == styleId ) { return (null, null); }

            Style style = a_stylePart?.Styles?.Elements<Style>().FirstOrDefault( s => s.StyleId == styleId );
            NumberingProperties styleNumPr = style?.StyleParagraphProperties?.NumberingProperties;
            return (
                (int?)styleNumPr?.NumberingId?.Val?.Value,
                (int?)styleNumPr?.NumberingLevelReference?.Val?.Value
            );
        }

        public static Dictionary<int, NumberingDefinition> LoadNumbring( WordprocessingDocument a_doc )
        {
            Dictionary<int, NumberingDefinition> result = new Dictionary<int, NumberingDefinition>();
            NumberingDefinitionsPart part = a_doc.MainDocumentPart.NumberingDefinitionsPart;
            if( null == part ) { return result; }

            Dictionary<int, NumberingDefinition> abstractMap = new Dictionary<int, NumberingDefinition>();

            foreach( AbstractNum abs in part.Numbering.Elements<AbstractNum>() )
            {
                NumberingDefinition def = new NumberingDefinition();
                foreach( Level lvl in abs.Elements<Level>() )
                {
                    def.Levels[(int)lvl.LevelIndex.Value] = new LevelDefinition
                    {
                        Text = null != lvl.LevelText ? lvl.LevelText.Val.Value : "",
                        Format = null != lvl.NumberingFormat ? lvl.NumberingFormat.Val.Value : (NumberFormatValues?)null,
                        Start = null != lvl.StartNumberingValue ? lvl.StartNumberingValue.Val.Value : 1
                    };
                }
                abstractMap[(int)abs.AbstractNumberId.Value] = def;
            }

            foreach( NumberingInstance num in part.Numbering.Elements<NumberingInstance>() )
            {
                int numId = (int)num.NumberID.Value;
                int absId = (int)num.AbstractNumId.Val.Value;

                if( abstractMap.ContainsKey( absId ) )
                {
                    result[numId] = abstractMap[absId];
                }
            }

            return result;
        }

        public static string GetVisibleText( OpenXmlElement a_element )
        {
            StringBuilder sb = new StringBuilder();

            string currentField = null;

            foreach( OpenXmlElement element in a_element.Descendants() )
            {
                /* フィールド開始・終了 */
                FieldChar fieldChar = element as FieldChar;
                if( null != fieldChar )
                {
                    if( FieldCharValues.End == fieldChar.FieldCharType )
                    {
                        currentField = null;
                    }

                    continue;
                }

                /* フィールドコード */
                FieldCode fieldCode = element as FieldCode;
                if( null != fieldCode )
                {
                    string txt = fieldCode.Text;

                    if( txt.Contains( "SEQ" ) )
                    {
                        currentField = "SEQ";
                    }
                    else if( txt.Contains( "STYLEREF" ) )
                    {
                        currentField = "STYLEREF";
                    }
                    else
                    {
                        /* 処理なし */
                    }

                    continue;
                }

                /* 表示テキスト */
                Text text = element as Text;
                if( null != text )
                {
                    string value = text.Text;

                    /* SEQ結果の前に "-" を補完 */
                    if( currentField == "SEQ" )
                    {
                        if( 0 < sb.Length &&
                            char.IsDigit( sb[sb.Length - 1] ) &&
                            0 < value.Length &&
                            char.IsDigit( value[0] ) )
                        {
                            sb.Append( "-" );
                        }
                    }

                    sb.Append( value );
                }
            }

            return sb.ToString();
        }

    }
}

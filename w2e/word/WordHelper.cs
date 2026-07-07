using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace w2e.word
{
    /// <summary>
    /// Wordファイルを操作するクラス
    /// </summary>
    public static class WordHelper
    {
        /// <summary>
        /// Word文書に定義されている番号付け情報（Numbering）の種別
        /// </summary>
        public enum NumberingTypeEn
        {
            NONE,
            HEADING,    /* 見出し */
            LIST        /* 箇条書き */
        }

        /// <summary>
        /// 指定された段落から章番号の情報を取得する
        /// </summary>
        /// <param name="a_pars">対象の段落</param>
        /// <param name="a_stylePart">スタイル定義パート（段落に直接番号情報が存在しない場合、スタイルから番号情報を取得するために使用する）</param>
        /// <returns>番号リストID, アウトラインレベル, 番号付け種別</returns>
        public static (int? numId, int? level, NumberingTypeEn numberingType) GetNumberingInfo( Paragraph a_pars, StyleDefinitionsPart a_stylePart )
        {
            /* スタイルが設定されていない、または見出しスタイルでない場合は空の番号情報を返す */

            /* 段落に設定されているスタイルIDを取得 */
            string styleId = a_pars.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            if( null == styleId )
            {
                /* スタイルが設定されていない場合は空の番号情報を返す */
                return ( null, null, NumberingTypeEn.NONE );
            }

            /* スタイルIDに一致するスタイル定義を取得 */
            Style style = a_stylePart?.Styles?.Elements<Style>().FirstOrDefault( s => s.StyleId == styleId );

            /* スタイル名を取得 */
            string styleName = style?.StyleName?.Val?.Value;

            /* 段落の番号情報を取得 */
            NumberingProperties numPr = a_pars.ParagraphProperties?.NumberingProperties;
]
            /* 見出しスタイルか判定
             * 箇条書きの番号を章番号として誤検出しないため、
             * heading 系スタイルのみ章番号対象とする
             */
            bool isHeadingStyle_flg = !string.IsNullOrEmpty( styleName ) &&
                                    styleName.StartsWith( "heading", StringComparison.OrdinalIgnoreCase );
            if( !isHeadingStyle_flg )
            {
                /* 見出しスタイルでない場合 */

                if( null != numPr )
                {
                    /* 見出しスタイルでないが段落の番号情報がある場合 */
                    return (
                        (int?)numPr.NumberingId?.Val?.Value,
                        (int?)numPr.NumberingLevelReference?.Val?.Value,
                        NumberingTypeEn.LIST
                    );
                }

                /* 見出しスタイルも段落の番号情報もない場合は空の番号情報を返す */
                return ( null, null, NumberingTypeEn.NONE );
            }

            /* 番号情報を取得する */

            /* 優先順位
             * 1. 段落に直接設定されている番号情報
             * 2. 段落スタイルに設定されている番号情報
             */

            /* 1. 段落に番号情報が設定されていた場合はその番号情報を返す */
            if( null != numPr )
            {
                return (
                    (int?)numPr.NumberingId?.Val?.Value,
                    (int?)numPr.NumberingLevelReference?.Val?.Value,
                    NumberingTypeEn.HEADING
                );
            }

            /* 2. 段落スタイルに設定されている番号情報を返す */
            NumberingProperties styleNumPr = style?.StyleParagraphProperties?.NumberingProperties;
            return (
                (int?)styleNumPr?.NumberingId?.Val?.Value,
                (int?)styleNumPr?.NumberingLevelReference?.Val?.Value,
                NumberingTypeEn.HEADING
            );
        }


        /// <summary>
        /// Word文書に定義されている番号付け情報（Numbering）を読み込んでNumberingIdをキーとした辞書として取得する
        /// </summary>
        /// <param name="a_doc">Wordドキュメント</param>
        /// <returns>NumberingIdをキーとする番号付け情報の辞書</returns>
        public static Dictionary<int, NumberingDefinition> LoadNumbring( WordprocessingDocument a_doc )
        {
            /* 処理内容：
             * ・AbstractNum（抽象番号定義）を読み込み
             * ・Levelごとの書式（テキスト、フォーマット、開始値）を取得
             * ・NumberingInstance（実体）と関連付けて最終的なマッピングを構築
             * 注意：
             * ・NumberingDefinitionsPartが存在しない場合は空の辞書を返します
             * ・AbstractNumに存在しない参照は無視されます
             */

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
                        text = null != lvl.LevelText ? lvl.LevelText.Val.Value : "",
                        format = null != lvl.NumberingFormat ? lvl.NumberingFormat.Val.Value : (NumberFormatValues?)null,
                        start = null != lvl.StartNumberingValue ? lvl.StartNumberingValue.Val.Value : 1
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


        /// <summary>
        /// OpenXml要素からフィールドコードを考慮してユーザーに表示されるテキストを抽出する
        /// </summary>
        /// <param name="a_element">OpenXml要素</param>
        /// <returns>ユーザーに表示されるテキスト</returns>
        public static string GetVisibleText( OpenXmlElement a_element )
        {
            /* 以下の処理を行います：
             * ・フィールド（FieldChar / FieldCode）を解析
             * ・SEQフィールドおよびSTYLEREFフィールドを識別
             * ・SEQフィールドの連番表示時に必要な区切り文字（"-"）を補完
             * 注意点：
             * ・フィールドコード自体は出力せず、表示結果のみを対象とします
             * ・複雑なフィールド構造（ネストなど）には完全対応していません
             */

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


        /// <summary>
        /// Word の表のセル内の文字列を取得する
        /// </summary>
        /// <param name="a_cell">文字列を取得する Word の表セル</param>
        /// <returns>セル内の文字列</returns>
        /// <remarks>
        /// Word の段落区切りや改行は改行コードに変換する
        /// </remarks>
        public static string GetCellText( TableCell a_cell )
        {
            StringBuilder sb = new StringBuilder();

            /* 最初の段落かどうかを判定するフラグ
             * 段落の先頭には改行を付与したくないため2段落目以降のみ改行する
             */
            bool firstParagraph = true;
            
            /* セル内の Paragraph（段落）を順番に処理する */
            foreach( Paragraph para in a_cell.Elements<Paragraph>() )
            {
                /* 2段落目以降の場合 */
                if( !firstParagraph )
                {
                    sb.Append( Environment.NewLine );
                }

                /* 1段落目の処理が終わったためフラグをオフにする */
                firstParagraph = false;

                /* 段落内の要素を順番に処理する */
                foreach( OpenXmlElement elem in para.Descendants() )
                {
                    if( elem is Break )
                    {
                        /* 改行の場合 */
                        sb.Append( Environment.NewLine );
                    }
                    else
                    {
                        if( elem is Text text )
                        {
                            /* 文字列の場合 */
                            sb.Append( text.Text );
                        }
                    }
                }
            }

            return sb.ToString();
        }


    }
}

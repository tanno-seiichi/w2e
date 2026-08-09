using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using Word = DocumentFormat.OpenXml.Wordprocessing;


namespace w2e.word
{
    /// <summary>
    /// Word文書内の画像を取得するヘルパークラス
    /// </summary>
    public static class WordImageHelper
    {
        /// <summary>
        /// 指定したインデックスの段落を起点に、画像を取得する。
        /// 画像を伴わない図形のみが複数の段落にまたがって離れた位置に配置されている場合は、
        /// それらをまとめて1枚の画像として合成し、消費した段落数を返す。
        /// （そうでない場合は通常通りその段落単体の画像取得と同じ結果になり、消費数は1になる）
        /// </summary>
        /// <param name="a_mainDocumentPart">MainDocumentPart</param>
        /// <param name="a_elements">Word本文の要素一覧</param>
        /// <param name="a_currentIndex">対象段落のインデックス</param>
        /// <param name="a_consumedCount">この呼び出しで消費した（読み飛ばすべき）要素数。呼び出し側はループのインデックスをこの数だけ進める</param>
        /// <returns>画像情報一覧</returns>
        public static List<WordImageData> GetImages( MainDocumentPart a_mainDocumentPart, IReadOnlyList<DocumentFormat.OpenXml.OpenXmlElement> a_elements, int a_currentIndex, out int a_consumedCount )
        {
            a_consumedCount = 1;

            if( !( a_elements[a_currentIndex] is Word.Paragraph currentParagraph ) )
            {
                return new List<WordImageData>();
            }

            /* 段落をまたいで離れた位置に配置されている図形のみの段落群が無いか、まず確認する */
            WordImageData multiParagraphResult = ShapeOverlayCompositor.TryComposeAcrossParagraphs( a_mainDocumentPart, a_elements, a_currentIndex, out int consumed );
            if( null != multiParagraphResult )
            {
                a_consumedCount = consumed;
                return new List<WordImageData>() { multiParagraphResult };
            }

            /* 対象外だった場合は、通常通りこの段落単体の画像取得処理を行う */
            return GetImages( a_mainDocumentPart, currentParagraph );
        }


        /// <summary>
        /// 指定した段落内に存在する画像を取得する。
        /// </summary>
        /// <param name="a_mainDocumentPart">MainDocumentPart</param>
        /// <param name="a_paragraph">対象段落</param>
        /// <returns>画像情報一覧</returns>
        public static List<WordImageData> GetImages( MainDocumentPart a_mainDocumentPart, Word.Paragraph a_paragraph )
        {
            List<WordImageData> imageList = new List<WordImageData>();

            /* 引数をチェックする */
            if( null == a_mainDocumentPart )
            {
                throw new ArgumentNullException( nameof( a_mainDocumentPart ) );
            }

            if( null == a_paragraph )
            {
                throw new ArgumentNullException( nameof( a_paragraph ) );
            }

            /* 画像に図形（矢印・強調枠など）が重なっている場合は、1枚に合成した画像として取得する */
            WordImageData composedImage = ShapeOverlayCompositor.TryCompose( a_mainDocumentPart, a_paragraph );

            if( null != composedImage )
            {
                imageList.Add( composedImage );
            }
            else
            {
                /* Drawing画像を取得する */
                GetDrawingImages( a_mainDocumentPart, a_paragraph, imageList );
            }

            /* VML画像を取得する（旧.doc形式の画像保持方式に対応） */
            GetVmlImages( a_mainDocumentPart, a_paragraph, imageList );

            return imageList;
        }


        /// <summary>
        /// 指定した段落内に存在するDrawing画像を取得する。
        /// </summary>
        /// <param name="a_mainDocumentPart">MainDocumentPart</param>
        /// <param name="a_paragraph">対象段落</param>
        /// <param name="a_imageList">画像情報一覧</param>
        private static void GetDrawingImages( MainDocumentPart a_mainDocumentPart, Word.Paragraph a_paragraph, List<WordImageData> a_imageList )
        {
            /* 段落内の画像を順番に取得する */
            IEnumerable<Word.Drawing> drawingList = a_paragraph.Descendants<Word.Drawing>();

            foreach( Word.Drawing drawing in drawingList )
            {
                Blip blip = GetBlip( drawing );

                if( null == blip ||
                    null == blip.Embed )
                {
                    continue;
                }

                ImagePart imagePart = a_mainDocumentPart.GetPartById( blip.Embed.Value ) as ImagePart;

                if( null == imagePart )
                {
                    continue;
                }

                WordImageData imageData = CreateImageData( drawing, imagePart, blip.Embed.Value );

                if( null != imageData )
                {
                    a_imageList.Add( imageData );
                }
            }
        }


        /// <summary>
        /// 指定した段落内に存在するVML画像を取得する。
        /// </summary>
        /// <param name="a_mainDocumentPart">MainDocumentPart</param>
        /// <param name="a_paragraph">対象段落</param>
        /// <param name="a_imageList">画像情報一覧</param>
        private static void GetVmlImages( MainDocumentPart a_mainDocumentPart, Word.Paragraph a_paragraph, List<WordImageData> a_imageList )
        {
            /* 段落内の画像を順番に取得する */
            IEnumerable<DocumentFormat.OpenXml.Vml.ImageData> imageDataList = a_paragraph.Descendants<DocumentFormat.OpenXml.Vml.ImageData>();

            foreach( DocumentFormat.OpenXml.Vml.ImageData imageData in imageDataList )
            {
                /* mc:AlternateContent の mc:Fallback（後方互換用のVML表現）内にある画像は、
                 * 対応する mc:Choice 側に「実際に解決できる」現代形式（w:drawing）の画像がある場合に限り、
                 * GetDrawingImages()で取得済みの画像と重複するため除外する。
                 * 万一 mc:Choice 側の画像が壊れている・取得できない文書だった場合に画像自体を失わないよう、
                 * Choice側が使えないときはVML側をそのまま採用する。
                 */
                if( IsInsideAlternateContentFallback( imageData ) &&
                    HasUsableChoiceDrawing( imageData, a_mainDocumentPart ) )
                {
                    continue;
                }

                if( null == imageData.RelationshipId )
                {
                    continue;
                }

                ImagePart imagePart = a_mainDocumentPart.GetPartById( imageData.RelationshipId ) as ImagePart;

                if( null == imagePart )
                {
                    continue;
                }

                WordImageData wordImageData = CreateImageData( imagePart, imageData.RelationshipId );

                if( null != wordImageData )
                {
                    a_imageList.Add( wordImageData );
                }
            }
        }


        /// <summary>
        /// 指定した要素が、mc:AlternateContent の mc:Fallback（後方互換用の代替コンテンツ）の内側にあるかどうかを判定する。
        /// </summary>
        /// <param name="a_element">判定対象の要素</param>
        /// <returns>mc:Fallbackの内側にある場合はtrue</returns>
        private static bool IsInsideAlternateContentFallback( DocumentFormat.OpenXml.OpenXmlElement a_element )
        {
            for( DocumentFormat.OpenXml.OpenXmlElement current = a_element.Parent; null != current; current = current.Parent )
            {
                if( current is DocumentFormat.OpenXml.AlternateContentFallback )
                {
                    return true;
                }
            }

            return false;
        }


        /// <summary>
        /// mc:Fallback内の要素に対して、同じ mc:AlternateContent の mc:Choice 側に
        /// 実際に画像データを解決できる w:drawing が存在するかどうかを判定する。
        /// </summary>
        /// <param name="a_fallbackElement">mc:Fallback内の要素（VMLのImageDataなど）</param>
        /// <param name="a_mainDocumentPart">MainDocumentPart</param>
        /// <returns>mc:Choice側に解決可能な画像がある場合はtrue</returns>
        private static bool HasUsableChoiceDrawing( DocumentFormat.OpenXml.OpenXmlElement a_fallbackElement, MainDocumentPart a_mainDocumentPart )
        {
            /* この要素を含む mc:AlternateContent を探す */
            DocumentFormat.OpenXml.AlternateContent alternateContent = null;
            for( DocumentFormat.OpenXml.OpenXmlElement current = a_fallbackElement.Parent; null != current; current = current.Parent )
            {
                if( current is DocumentFormat.OpenXml.AlternateContent ac )
                {
                    alternateContent = ac;
                    break;
                }
            }

            if( null == alternateContent )
            {
                /* mc:AlternateContentが見つからない場合は判定できないため、Choice側は無いものとして扱う */
                return false;
            }

            DocumentFormat.OpenXml.AlternateContentChoice choice =
                alternateContent.GetFirstChild<DocumentFormat.OpenXml.AlternateContentChoice>();

            if( null == choice )
            {
                return false;
            }

            /* mc:Choice側にあるw:drawingのうち、1つでも画像データを実際に解決できればtrueとする */
            foreach( Word.Drawing drawing in choice.Descendants<Word.Drawing>() )
            {
                Blip blip = GetBlip( drawing );

                if( null == blip || null == blip.Embed )
                {
                    continue;
                }

                if( a_mainDocumentPart.GetPartById( blip.Embed.Value ) is ImagePart )
                {
                    return true;
                }
            }

            return false;
        }


        /// <summary>
        /// ImagePartから画像情報を生成する。
        /// </summary>
        /// <param name="a_drawing">Drawing</param>
        /// <param name="a_imagePart">ImagePart</param>
        /// <param name="a_relationshipId">RelationshipId</param>
        /// <returns>画像情報</returns>
        private static WordImageData CreateImageData( Word.Drawing a_drawing, ImagePart a_imagePart, string a_relationshipId )
        {
            /* 画像データを取得する */
            byte[] imageData = GetImageData( a_imagePart );

            /* コンテンツタイプを取得する */
            string contentType = GetContentType( a_imagePart );

            /* 画像サイズを取得する（EMU単位） */
            GetImageSize( a_drawing, out long widthEmu, out long heightEmu );

            /* AltTextを取得する */
            string altText = GetAltText( a_drawing );

            /* 画像情報を生成する */
            WordImageData result = new WordImageData();

            result.imageData = imageData;
            result.contentType = contentType;
            result.relationshipId = a_relationshipId;
            result.widthEmu = widthEmu;
            result.heightEmu = heightEmu;
            result.altText = altText;

            /* 画像サイズ・AltTextは後続コミットで設定する */
            return result;
        }


        private static WordImageData CreateImageData( ImagePart a_imagePart, string a_relationshipId )
        {
            byte[] imageData = GetImageData( a_imagePart );

            string contentType = GetContentType( a_imagePart );

            WordImageData result = new WordImageData();

            result.imageData = imageData;
            result.contentType = contentType;
            result.relationshipId = a_relationshipId;

            return result;
        }


        /// <summary>
        /// ImagePartから画像データを取得する。
        /// </summary>
        /// <param name="a_imagePart">ImagePart</param>
        /// <returns>画像データ</returns>
        internal static byte[] GetImageData( ImagePart a_imagePart )
        {
            /* 引数チェック */
            if( null == a_imagePart )
            {
                throw new ArgumentNullException( nameof( a_imagePart ) );
            }

            /* ストリームから画像データを取得する */
            using( System.IO.Stream stream = a_imagePart.GetStream() )
            {
                using( System.IO.MemoryStream memoryStream = new System.IO.MemoryStream() )
                {
                    stream.CopyTo( memoryStream );
                    return memoryStream.ToArray();
                }
            }
        }


        /// <summary>
        /// ImagePartからコンテンツタイプを取得する。
        /// </summary>
        /// <param name="a_imagePart">ImagePart</param>
        /// <returns>コンテンツタイプ</returns>
        internal static string GetContentType( ImagePart a_imagePart )
        {
            /* 引数チェック */
            if( null == a_imagePart )
            {
                throw new ArgumentNullException( nameof( a_imagePart ) );
            }

            /* コンテンツタイプを返却する */
            return a_imagePart.ContentType;
        }


        /// <summary>
        /// Drawingから画像サイズを取得する。
        /// </summary>
        /// <param name="a_drawing">Drawing</param>
        /// <param name="a_widthEmu">画像幅（EMU）</param>
        /// <param name="a_heightEmu">画像高さ（EMU）</param>
        internal static void GetImageSize( Word.Drawing a_drawing, out long a_widthEmu, out long a_heightEmu )
        {
            /* 初期化 */
            a_widthEmu = 0;
            a_heightEmu = 0;

            /* 引数チェック */
            if( null == a_drawing )
            {
                return;
            }

            /* inline画像のサイズ取得 */
            DocumentFormat.OpenXml.Drawing.Extents extents =
        a_drawing.Descendants<DocumentFormat.OpenXml.Drawing.Extents>().FirstOrDefault();

            if( null == extents )
            {
                return;
            }

            a_widthEmu = extents.Cx;
            a_heightEmu = extents.Cy;
        }


        /// <summary>
        /// DrawingからAltTextを取得する。
        /// </summary>
        /// <param name="a_drawing">Drawing</param>
        /// <returns>AltText</returns>
        private static string GetAltText( Word.Drawing a_drawing )
        {
            /* 引数チェック */
            if( null == a_drawing )
            {
                return string.Empty;
            }

            /* wp:docPr から代替テキストを取得する */
            DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties docProps =
        a_drawing.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties>().FirstOrDefault();

            if( null == docProps )
            {
                return string.Empty;
            }

            if( null == docProps.Description )
            {
                return string.Empty;
            }

            return docProps.Description.Value;
        }


        /// <summary>
        /// DrawingからBlipを取得する。
        /// </summary>
        /// <param name="a_drawing">Drawing</param>
        /// <returns>Blip。取得できない場合はnull。</returns>
        internal static Blip GetBlip( Word.Drawing a_drawing )
        {
            if( null == a_drawing )
            {
                return null;
            }

            return a_drawing
                .Descendants<Blip>()
                .FirstOrDefault();
        }


    }
}
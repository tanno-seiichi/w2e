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

            /* Drawing画像を取得する */
            GetDrawingImages( a_mainDocumentPart, a_paragraph, imageList );

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
        private static byte[] GetImageData( ImagePart a_imagePart )
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
        private static string GetContentType( ImagePart a_imagePart )
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
        private static void GetImageSize( Word.Drawing a_drawing, out long a_widthEmu, out long a_heightEmu )
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
        private static Blip GetBlip( Word.Drawing a_drawing )
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
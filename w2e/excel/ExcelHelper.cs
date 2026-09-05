using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using w2e.word;
using A = DocumentFormat.OpenXml.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace w2e.excel
{
    /// <summary>
    /// Excelを操作するクラス
    /// </summary>
    public static class ExcelHelper
    {
        /// <summary>
        /// Excelのワークシートを生成する
        /// </summary>
        /// <param name="a_wbPart">ワークシートの追加先となるブック</param>
        /// <param name="a_sheets">ブック内のシートコレクション</param>
        /// <param name="a_sheetName">作成するワークシートの名前</param>
        /// <param name="a_sheetId">ワークシートに割り当てるID（ブック内で一意である必要がある）</param>
        /// <param name="a_sheetData">ワークシートの内容</param>
        /// <returns> Excelワークシート</returns>
        public static WorksheetPart CreateWorksheet( WorkbookPart a_wbPart, Sheets a_sheets, string a_sheetName, uint a_sheetId, out SheetData a_sheetData )
        {
            WorksheetPart wsPart = a_wbPart.AddNewPart<WorksheetPart>();

            a_sheetData = new SheetData();
            wsPart.Worksheet = new Worksheet( a_sheetData );

            Sheet sheet = new Sheet()
            {
                Id = a_wbPart.GetIdOfPart( wsPart ),
                SheetId = a_sheetId,
                Name = a_sheetName
            };

            a_sheets.Append( sheet );
            return wsPart;
        }


        /// <summary>
        /// Excelシート名の禁止文字を除去して返す
        /// </summary>
        /// <param name="a_name">禁止文字を除去する前のシート名</param>
        /// <returns>禁止文字を除去したシート名</returns>
        public static string SafeSheetName( string a_name )
        {
            /* Excelシート名の禁止文字を半角スペースに置換 */
            char[] invalidid = { '\\', '/', '*', '[', ']', ':', '?', ',', '、', '／', '：' };
            foreach( char c in invalidid )
            {
                a_name = a_name.Replace( c, ' ' );
            }

            /* 全角スペースを除去 */
            a_name = a_name.Replace( "　", "" );

            /* 改行コードを除去 */
            a_name = a_name.Replace( Environment.NewLine, "" );

            /* Excelシート名の長さ制限チェック */
            if( 31 < a_name.Length )
            {
                a_name = a_name.Substring( 0, 31 );
            }

            return string.IsNullOrWhiteSpace( a_name ) ? "Sheet" : a_name.Trim();
        }


        /// <summary>
        /// WorkbookStylesPart と Stylesheet を安全に初期化する。
        /// Excelが要求する最小構造（Fonts/Fills/Borders/CellFormats）を必ず1件以上持たせる。
        /// </summary>
        /// <param name="a_wbPart">WorkbookPart（スタイル情報のルート）</param>
        public static void InitializeStylesheet( WorkbookPart a_wbPart )
        {
            /* WorkbookStylesPart取得（無ければ作成） */
            WorkbookStylesPart stylesPart = a_wbPart.WorkbookStylesPart;
            if( null == stylesPart )
            {
                stylesPart = a_wbPart.AddNewPart<WorkbookStylesPart>();
                stylesPart.Stylesheet = new Stylesheet();
            }

            /* Stylesheet取得（無ければ生成） */
            Stylesheet styles = stylesPart.Stylesheet;
            if( null == styles )
            {
                styles = new Stylesheet();
                stylesPart.Stylesheet = styles;
            }

            /* Fonts初期化（最低1件必要） */
            if( null == styles.Fonts )
            {
                styles.Fonts = new Fonts( new Font() );
            }
            else if( styles.Fonts.Count() == 0 )
            {
                styles.Fonts.AppendChild( new Font() );
            }

            /* Fills初期化（最低1件必要） */
            if( null == styles.Fills )
            {
                styles.Fills = new Fills( new Fill() );
            }
            else if( styles.Fills.Count() == 0 )
            {
                styles.Fills.AppendChild( new Fill() );
            }

            /* Borders初期化（最低1件必要） */
            if( null == styles.Borders )
            {
                styles.Borders = new Borders( new Border() );
            }
            else if( styles.Borders.Count() == 0 )
            {
                styles.Borders.AppendChild( new Border() );
            }

            /* CellFormats初期化（最低1件必要） */
            if( null == styles.CellFormats )
            {
                styles.CellFormats = new CellFormats( new CellFormat() );
            }
            else if( styles.CellFormats.Count() == 0 )
            {
                styles.CellFormats.AppendChild( new CellFormat() );
            }

            /* Count属性を整合させる */
            styles.Fonts.Count = (uint)styles.Fonts.Count();
            styles.Fills.Count = (uint)styles.Fills.Count();
            styles.Borders.Count = (uint)styles.Borders.Count();
            styles.CellFormats.Count = (uint)styles.CellFormats.Count();
        }


        /// <summary>
        /// 画像サイズが不明な場合に使用する既定の幅（EMU）
        /// </summary>
        private const long DEFAULT_IMAGE_WIDTH_EMU = 3048000;

        /// <summary>
        /// 画像サイズが不明な場合に使用する既定の高さ（EMU）
        /// </summary>
        private const long DEFAULT_IMAGE_HEIGHT_EMU = 2286000;

        /// <summary>
        /// Excelの既定の行の高さ（EMU）画像が占有する行数の算出に使用する
        /// </summary>
        private const long DEFAULT_ROW_HEIGHT_EMU = 190500;

        /// <summary>
        /// EMUとポイントの変換係数（OOXML標準：1ポイント = 12700EMU）
        /// </summary>
        private const double EMU_PER_POINT = 12700.0;

        /// <summary>
        /// 画像を貼付ける行の高さを計算する際、画像の下に余白として追加するポイント数
        /// </summary>
        private const double IMAGE_ROW_HEIGHT_MARGIN_POINTS = 4.0;

        /// <summary>
        /// セルの罫線と画像が重ならないよう、画像をセルの左上から右下にずらすオフセット（EMU）
        /// 行の高さに追加する余白（IMAGE_ROW_HEIGHT_MARGIN_POINTS）のうち、上側の分に相当する
        /// </summary>
        private const long IMAGE_CELL_OFFSET_EMU = 25400;


        /// <summary>
        /// 画像の高さから、その画像を貼付ける行に設定すべき行の高さ（pt）を算出する。
        /// セルの罫線内に画像がきれいに収まるよう、行の高さを画像の高さに合わせて拡張する用途に使用する。
        /// </summary>
        /// <param name="a_image">対象の画像情報（サイズ不明の場合は既定サイズとして扱う）</param>
        /// <returns>行の高さ（pt）</returns>
        public static double CalculateRowHeightForImage( WordImageData a_image )
        {
            long heightEmu = ( null != a_image && 0 < a_image.heightEmu ) ? a_image.heightEmu : DEFAULT_IMAGE_HEIGHT_EMU;

            return ( heightEmu / EMU_PER_POINT ) + IMAGE_ROW_HEIGHT_MARGIN_POINTS;
        }


        /// <summary>
        /// 行の高さの見積りに使用する既定の列幅（px, 96DPI換算）
        /// 本アプリでは列幅を明示的に指定していないため、Excelの既定列幅（8.43文字相当）を用いる
        /// </summary>
        private const double DEFAULT_COLUMN_WIDTH_PIXELS = 64.0;

        /// <summary>
        /// セル内側の余白（左右合計、px, 96DPI換算）。折返し判定の際に列幅から差し引く
        /// </summary>
        private const double CELL_PADDING_PIXELS = 10.0;

        /// <summary>
        /// 行数見積りに使用するフォント名（実際にセルへ設定しているフォントに関わらず、
        /// 折返し行数の目安を得るための計測用フォントとして使用する）
        /// </summary>
        private const string ESTIMATE_FONT_NAME = "MS Pゴシック";

        /// <summary>
        /// 行数見積りに使用するフォントサイズ（pt）
        /// </summary>
        private const float ESTIMATE_FONT_SIZE = 11.0f;


        /// <summary>
        /// セルのテキストが実際に何行に折返されるかを、実フォントでの文字幅計測により見積もる。
        /// 明示的な改行（段落区切り等）に加え、列幅に収まらない場合の自動折返しも考慮する。
        /// </summary>
        /// <param name="a_text">対象セルのテキスト（Environment.NewLine 区切りを想定）</param>
        /// <param name="a_columnSpan">セルが占める列数（横結合されている場合は2以上）</param>
        /// <returns>折返し後の行数（1以上）</returns>
        public static int EstimateWrappedLineCount( string a_text, int a_columnSpan )
        {
            if( string.IsNullOrEmpty( a_text ) )
            {
                return 1;
            }

            double columnWidthPixels = ( DEFAULT_COLUMN_WIDTH_PIXELS * Math.Max( 1, a_columnSpan ) ) - CELL_PADDING_PIXELS;
            if( columnWidthPixels < 1.0 )
            {
                columnWidthPixels = 1.0;
            }

            string[] explicitLines = a_text.Split( new[] { "\r\n", "\n" }, StringSplitOptions.None );

            int totalLines = 0;

            using( System.Drawing.Font font = new System.Drawing.Font( ESTIMATE_FONT_NAME, ESTIMATE_FONT_SIZE ) )
            using( System.Drawing.Bitmap dummyBitmap = new System.Drawing.Bitmap( 1, 1 ) )
            using( System.Drawing.Graphics g = System.Drawing.Graphics.FromImage( dummyBitmap ) )
            {
                foreach( string line in explicitLines )
                {
                    if( string.IsNullOrEmpty( line ) )
                    {
                        /* 空行も1行として数える */
                        totalLines += 1;
                        continue;
                    }

                    /* 折返し無しの状態でその行全体の幅を計測し、列幅で割って折返し行数を求める */
                    System.Drawing.SizeF size = g.MeasureString( line, font, int.MaxValue );
                    int wrappedCount = (int)Math.Ceiling( size.Width / columnWidthPixels );

                    totalLines += Math.Max( 1, wrappedCount );
                }
            }

            return Math.Max( 1, totalLines );
        }


        /// <summary>
        /// セル内のテキストから、Excelの自動調整相当の行の高さ（pt）を見積もる。
        /// 実フォントでの文字幅計測により、列幅に収まらない場合の自動折返しも考慮した行数を求め、
        /// 1行あたり Excel の既定の行の高さ（DEFAULT_ROW_HEIGHT_EMU）分の高さが必要と仮定して計算する。
        /// </summary>
        /// <param name="a_text">対象セルのテキスト（Environment.NewLine 区切りを想定）</param>
        /// <param name="a_columnSpan">セルが占める列数（横結合されている場合は2以上）</param>
        /// <returns>行の高さ（pt）</returns>
        public static double EstimateRowHeightForText( string a_text, int a_columnSpan = 1 )
        {
            double singleLineHeightPoints = DEFAULT_ROW_HEIGHT_EMU / EMU_PER_POINT;
            int lineCount = EstimateWrappedLineCount( a_text, a_columnSpan );

            return lineCount * singleLineHeightPoints;
        }


        /// <summary>
        /// Excelワークシートに画像を挿入する
        /// </summary>
        /// <remarks>
        /// ワークシートに対する最初の呼び出し時に DrawingsPart を生成し、
        /// 以降の呼び出しでは同じ DrawingsPart に画像を追加していく
        /// シートが切り替わった場合は呼び出し元で a_drawingsPart / a_imageId を初期化すること
        /// </remarks>
        /// <param name="a_wsPart">画像を挿入する対象のワークシート</param>
        /// <param name="a_drawingsPart">ワークシートに紐づくDrawingsPart（未生成の場合はnullを渡し、このメソッドが生成する）</param>
        /// <param name="a_imageId">画像に割り当てる一意なID（呼び出す度に採番され更新される）</param>
        /// <param name="a_image">挿入する画像情報</param>
        /// <param name="a_rowIndex">画像を配置する行番号（0始まり）</param>
        /// <param name="a_colIndex">画像を配置する列番号（0始まり）</param>
        /// <returns>画像の高さから算出した、確保すべき行数</returns>
        public static int AddImage( WorksheetPart a_wsPart, ref DrawingsPart a_drawingsPart, ref uint a_imageId, WordImageData a_image, int a_rowIndex, int a_colIndex )
        {
            /* 引数チェック */
            if( null == a_wsPart ||
                null == a_image ||
                null == a_image.imageData ||
                0 == a_image.imageData.Length )
            {
                return 0;
            }

            /* DrawingsPartが未生成の場合は生成し、ワークシートに関連付ける */
            if( null == a_drawingsPart )
            {
                a_drawingsPart = a_wsPart.AddNewPart<DrawingsPart>();
                a_drawingsPart.WorksheetDrawing = new Xdr.WorksheetDrawing();

                a_wsPart.Worksheet.Append( new Drawing() { Id = a_wsPart.GetIdOfPart( a_drawingsPart ) } );
            }

            /* 画像パートを追加し、画像データを書込む */
            PartTypeInfo imagePartType = GetImagePartType( a_image.contentType );
            ImagePart imagePart = a_drawingsPart.AddImagePart( imagePartType );
            using( System.IO.MemoryStream stream = new System.IO.MemoryStream( a_image.imageData ) )
            {
                imagePart.FeedData( stream );
            }

            /* 画像サイズを決定する（サイズ不明の場合は既定値を使用する） */
            long widthEmu = ( 0 < a_image.widthEmu ) ? a_image.widthEmu : DEFAULT_IMAGE_WIDTH_EMU;
            long heightEmu = ( 0 < a_image.heightEmu ) ? a_image.heightEmu : DEFAULT_IMAGE_HEIGHT_EMU;

            uint id = a_imageId++;
            string picName = "Picture " + id;
            string altText = a_image.altText ?? "";

            /* 画像の配置情報（アンカー）を生成する
             * セルの罫線（上・左）と画像が重ならないよう、少しだけ右下にオフセットして配置する
             */
            Xdr.OneCellAnchor anchor = new Xdr.OneCellAnchor(
                new Xdr.FromMarker()
                {
                    ColumnId = new Xdr.ColumnId( a_colIndex.ToString() ),
                    ColumnOffset = new Xdr.ColumnOffset( IMAGE_CELL_OFFSET_EMU.ToString() ),
                    RowId = new Xdr.RowId( a_rowIndex.ToString() ),
                    RowOffset = new Xdr.RowOffset( IMAGE_CELL_OFFSET_EMU.ToString() )
                },
                new Xdr.Extent() { Cx = widthEmu, Cy = heightEmu },
                new Xdr.Picture(
                    new Xdr.NonVisualPictureProperties(
                        new Xdr.NonVisualDrawingProperties() { Id = id, Name = picName, Description = altText },
                        new Xdr.NonVisualPictureDrawingProperties( new A.PictureLocks() { NoChangeAspect = true } )
                    ),
                    new Xdr.BlipFill(
                        new A.Blip() { Embed = a_drawingsPart.GetIdOfPart( imagePart ) },
                        new A.Stretch( new A.FillRectangle() )
                    ),
                    new Xdr.ShapeProperties(
                        new A.Transform2D(
                            new A.Offset() { X = 0, Y = 0 },
                            new A.Extents() { Cx = widthEmu, Cy = heightEmu }
                        ),
                        new A.PresetGeometry( new A.AdjustValueList() ) { Preset = A.ShapeTypeValues.Rectangle }
                    )
                ),
                new Xdr.ClientData()
            );

            a_drawingsPart.WorksheetDrawing.Append( anchor );

            /* 画像の高さから、後続コンテンツと重ならないよう確保すべき行数を算出する（余白として1行加算） */
            return (int)Math.Ceiling( (double)heightEmu / DEFAULT_ROW_HEIGHT_EMU ) + 1;
        }


        /// <summary>
        /// 画像のコンテンツタイプから PartTypeInfo（画像パート種別）を判定する
        /// </summary>
        /// <param name="a_contentType">画像のコンテンツタイプ（例: "image/png"）</param>
        /// <returns>対応する PartTypeInfo (判定できない場合は Png)</returns>
        private static PartTypeInfo GetImagePartType( string a_contentType )
        {
            switch( ( a_contentType ?? "" ).ToLower() )
            {
                case "image/png":
                    return ImagePartType.Png;
                case "image/jpeg":
                case "image/jpg":
                    return ImagePartType.Jpeg;
                case "image/gif":
                    return ImagePartType.Gif;
                case "image/bmp":
                    return ImagePartType.Bmp;
                case "image/tiff":
                    return ImagePartType.Tiff;
                default:
                    return ImagePartType.Png;
            }
        }


        /// <summary>
        /// Excelの列名を取得する
        /// </summary>
        /// <param name="a_index">列インデックス（例：1 = A、26 = Z、27 = AA）</param>
        /// <returns>Excel形式の列名文字列</returns>
        private static string GetColumnName( int a_index )
        {
            /* 前提：Excelは 26進数、ただし0がない（A=1） */

            string name = "";   // 生成する列名

            /* インデックスを26進数（A-Z）として扱い、右側の桁から順に算出する */
            while( 0 < a_index )
            {
                /* 0始まりに補正（A=0として扱うため） */
                a_index--;

                /* 現在の桁の文字を算出して先頭に追加 */
                name = (char)( 'A' + ( a_index % 26 ) ) + name;

                /* 次の桁へ */
                a_index /= 26;
            }
            return name;
        }


        /// <summary>
        /// Excelの行を追加する
        /// </summary>
        /// <param name="a_wbPart">WorkbookPart（スタイル情報のルート）</param>
        /// <param name="a_sheetData">Excelワークシート</param>
        /// <param name="a_rowIndex">行番号</param>
        /// <param name="a_values">行に設定するセルデータの一覧</param>
        /// <param name="a_cache">Excelのスタイルシートに登録済のスタイルを再利用するためのキャッシュ</param>
        /// <param name="a_rowHeightPoints">
        /// 行の高さ（pt）を明示的に指定する場合に指定する（画像を貼付ける行など）。
        /// 指定しない場合はExcelの既定の高さ・自動調整に委ねる。
        /// </param>
        public static void SetRow( WorkbookPart a_wbPart, SheetData a_sheetData, int a_rowIndex, List<CellData> a_values, Dictionary<string, uint> a_cache, double? a_rowHeightPoints = null )
        {
            Row row = new Row();
            row.RowIndex = (uint)a_rowIndex;

            if( a_rowHeightPoints.HasValue )
            {
                row.Height = a_rowHeightPoints.Value;
                row.CustomHeight = true;
            }

            a_sheetData.Append( row );

            for( int i = 0; i < a_values.Count; i++ )
            {
                CellData data = a_values[i];

                Cell cell = new Cell();
                cell.CellReference = GetColumnName( i + 1 ) + a_rowIndex;

                /* セルに値を設定 */
                if( string.IsNullOrEmpty( data.text ) )
                {
                    /* セルがブランクの場合 */
                    cell.DataType = null;
                    cell.CellValue = null;
                }
                else
                {
                    /* セルがブランクでない場合 */
                    cell.DataType = CellValues.String;
                    cell.CellValue = new CellValue( data.text ?? "" );
                }

                /* セルに枠線を設定 */
                BorderHelper.ApplyBorder( a_wbPart, cell, data.topBorder, data.bottomBorder, data.leftBorder, data.rightBorder, data.rightAlign, data.bold, a_cache );

                row.Append( cell );
            }
        }


    }
}

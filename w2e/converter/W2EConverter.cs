using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using w2e.delegates;
using w2e.excel;
using w2e.file;
using w2e.word;
using Excel = DocumentFormat.OpenXml.Spreadsheet;
using Word = DocumentFormat.OpenXml.Wordprocessing;

namespace w2e.converter
{
    /// <summary>
    /// Wordファイルの内容をExcelファイルに書き出すクラス
    /// </summary>
    public class W2EConverter : IConverter
    {
        /// <summary>
        /// 開始時の進捗値を表す定数
        /// </summary>
        private const int PROGRESS_MIN_VALUE = 1;

        /// <summary>
        /// Word読込完了時の進捗値を表す定数
        /// </summary>
        private const int PROGRESS_WORD_RANGE = 40;

        /// <summary>
        /// Excel書出完了時の進捗値を表す定数
        /// </summary>
        private const int PROGRESS_EXCEL_RANGE = 60;

        /// <summary>
        /// 先頭シートのシート名
        /// </summary>
        private const string TOP_SHEET_NAME = "トップ";

        /// <summary>
        /// 画像を配置する列番号 (0始まり。番号列,テキスト列と重ならない列に配置する)
        /// </summary>
        private const int IMAGE_COLUMN_INDEX = 2;
        
        /// <summary>
        /// 進捗情報が更新された時の処理
        /// </summary>
        public Delegates.UpdateProgressDelegate onProgressUpdate { get; set; }

        /// <summary>
        /// ログが出力された時の処理
        /// </summary>
        public Delegates.UpdateLogDelegate onLogUpdate { get; set; }


        /// <summary>
        /// Word → Excel の変換処理を実行する
        /// </summary>
        /// <param name="a_wordPath">Wordファイルのパス</param>
        /// <param name="a_excelPath">Excelファイルのパス</param>
        /// <param name="a_outputImage_flg">画像を使用するか否か</param>
        /// <param name="a_outputListNumber_flg">箇条書きに番号（Wordの実際の記号）を使用するか否か。falseの場合は固定で「・」を使用する</param>
        /// <param name="a_token">処理中断通知</param>
        public void Convert( string a_wordPath, string a_excelPath, bool a_outputImage_flg, bool a_outputListNumber_flg, CancellationToken a_token )
        {
            onProgressUpdate?.Invoke( PROGRESS_MIN_VALUE );
            string tempPath = FileCopy.CreateTempCopy( a_wordPath );

            try
            {
                /* Wordファイルを読込 */
                using( WordprocessingDocument doc = WordprocessingDocument.Open( tempPath, false ) )
                {
                    Word.Body body = doc.MainDocumentPart.Document.Body;
                    StyleDefinitionsPart stylePart = doc.MainDocumentPart.StyleDefinitionsPart;

                    Dictionary<int, NumberingDefinition> numberingMap = WordHelper.LoadNumbring( doc );
                    NumberingEngine engine = new NumberingEngine();

                    /* Excelファイルを生成 */
                    using( SpreadsheetDocument spreadsheet = SpreadsheetDocument.Create( a_excelPath, SpreadsheetDocumentType.Workbook ) )
                    {
                        /* Excelワークブックを生成 */
                        WorkbookPart wbPart = spreadsheet.AddWorkbookPart();
                        wbPart.Workbook = new Excel.Workbook();

                        /* Excelのスタイルシートを初期化 */
                        ExcelHelper.InitializeStylesheet( wbPart );

                        /* Excelのスタイルシートに登録済のスタイルを再利用するためのキャッシュを生成 */
                        var cache = new Dictionary<string, uint>();

                        /* Excelワークブックのシートを追加する準備 */
                        Excel.Sheets sheets = wbPart.Workbook.AppendChild( new Excel.Sheets() );

                        WorksheetPart wsPart = null;
                        Excel.SheetData sheetData = null;
                        string sheetName = TOP_SHEET_NAME;
                        uint sheetId = 1;
                        int row = 1;

                        /* 画像貼付用のDrawingsPartと画像ID（シートが切り替わるたびに初期化する） */
                        DrawingsPart drawingsPart = null;
                        uint imageId = 1;

                        /* 空行が3行以上連続しないようにするための直前の連続空行数（シートが切り替わるたびに初期化する） */
                        int consecutiveBlankRows = 0;

                        /* 箇条書きの記号（①、a)、・ など）を生成するためのnumIdごとのカウンタ（文書全体で共有し、シートが切り替わっても初期化しない） */
                        var listCounters = new Dictionary<int, Dictionary<int, int>>();

                        /* 直前に新規ファイルとして採用した章タイトルを保持する
                         * （見出しテキストが空で、次の非空段落をタイトルとして補完する際に、
                         *   その補完先が直前の章タイトルと同一だった場合は、
                         *   実体のない空の見出し段落（コピー&ペースト等で紛れ込んだもの）を
                         *   誤って新しい章として分割してしまわないようにするために使用する）
                         */
                        string lastChapterTitle = null;

                        /* 直前に処理した図形（矢印・強調枠等）の中に設定されていたテキストを保持する。
                         * Word上でその図形の陰に隠れて表示されない「重複した通常の段落」を検出するために使用する
                         * （詳細は使用箇所のコメントを参照）
                         */
                        var recentShapeTexts = new HashSet<string>();

                        int total = body.Elements().Count();
                        int current = 0;

                        /* プログレスバーをWordファイル読込終了まで進める */
                        onProgressUpdate?.Invoke( PROGRESS_WORD_RANGE );

                        /* Wordファイルの要素一覧（見出しのタイトル補完のため、次要素を先読みできるようリスト化する） */
                        List<OpenXmlElement> elements = body.Elements().ToList();

                        /* Wordファイルの要素ごとに処理 */
                        for( int elementIndex = 0; elementIndex < elements.Count; elementIndex++ )
                        {
                            OpenXmlElement element = elements[elementIndex];

                            /* 処理中断が要求されていたらループを抜ける */
                            if( a_token.IsCancellationRequested ) { break; }

                            /* プログレスバーを更新 */
                            current++;
                            int progress = PROGRESS_WORD_RANGE + (int)(current * PROGRESS_EXCEL_RANGE / total );
                            onProgressUpdate?.Invoke( progress );

                            /* Wordファイル「段落」の処理 */
                            Word.Paragraph para = element as Word.Paragraph;
                            if( null != para )
                            {
                                var info = WordHelper.GetNumberingInfo( para, stylePart );
                                int? numId = info.numId;
                                int? level = info.level;
                                WordHelper.NumberingTypeEn numberingType = info.numberingType;

                                CellData textData = new CellData() { text = WordHelper.GetVisibleText( para ) };
                                CellData numData = new CellData() { text = "" };

                                /* この段落のテキストが、直前に処理した図形（矢印や強調枠）の中のテキストと
                                 * 完全に一致する場合は、Word上でその図形の陰に隠れて表示されない「重複した通常の段落」
                                 * （コピー&ペースト等の残骸）とみなし、出力自体を丸ごとスキップする
                                 */
                                if( a_outputImage_flg &&
                                    !string.IsNullOrEmpty( textData.text ) &&
                                    recentShapeTexts.Contains( textData.text.Trim() ) )
                                {
                                    continue;
                                }

                                bool isHeading_flg = WordHelper.NumberingTypeEn.HEADING == numberingType;
                                bool isList_flg = WordHelper.NumberingTypeEn.LIST == numberingType;

                                /* この段落が新しい章の開始（新規シートを作るべき見出し）かどうか */
                                bool isNewChapter_flg = false;

                                /* この段落が、直前の章と重複する見出し（誤って紛れ込んだ空・重複見出し段落）であり、
                                 * 出力自体を丸ごとスキップすべきかどうか
                                 */
                                bool isDuplicateHeading_flg = false;

                                if( isHeading_flg )
                                {
                                    /* engine.Generateはカウンタをインクリメントするため、後で重複見出しと判明した場合に
                                     * 巻き戻せるよう、呼び出し前の状態を保存しておく
                                     */
                                    Dictionary<int, int> engineStateBeforeGenerate = engine.SaveState();
                                    bool usedEngineGenerate_flg = true;

                                    if( numId.HasValue && numberingMap.ContainsKey( numId.Value ) )
                                    {
                                        /* Wordの番号定義（numPr）から章番号を取得できる場合 */
                                        int levelValue = level.HasValue ? level.Value : 0;
                                        numData.text = engine.Generate( numberingMap[numId.Value], levelValue );
                                        usedEngineGenerate_flg = true;
                                    }
                                    else if( WordHelper.TryExtractLeadingNumber( textData.text, out string extractedNum, out string extractedTitle ) )
                                    {
                                        /* Wordの番号定義が無い場合は、見出しテキスト先頭の数字パターン（"3" "4.1"など）を章番号として代用する */
                                        numData.text = extractedNum;
                                        textData.text = extractedTitle;
                                    }

                                    /* シートを分割する条件はあくまで「章番号を取得できたかどうか」とする。
                                     * （本文が空でアウトラインレベルだけ設定された見出しや、番号を伴わない見出しスタイルの段落まで
                                     *   新しい章として分割してしまわないようにするため）
                                     */
                                    if( !string.IsNullOrEmpty( numData.text ) )
                                    {
                                        /* 章番号は取得できたがタイトルが空の場合（見出しが章番号のみの場合）は、
                                         * 次の空白でない段落をタイトルとして補完する
                                         */
                                        if( string.IsNullOrEmpty( textData.text ) )
                                        {
                                            string fallbackTitle = WordHelper.FindNextNonBlankParagraphText( elements, elementIndex + 1 );
                                            if( !string.IsNullOrEmpty( fallbackTitle ) )
                                            {
                                                textData.text = fallbackTitle;
                                            }
                                        }

                                        isNewChapter_flg = !string.IsNullOrEmpty( textData.text );

                                        /* 採用予定のタイトルが、直前に採用した章タイトルと完全に同一である場合は、
                                         * Word上に紛れ込んだ空・非表示の重複見出し段落（コピー&ペーストの残骸や
                                         * 削除履歴の残骸など）を誤って新章として分割してしまったものとみなし、
                                         * 新規ファイル作成を取り消す
                                         * （見出しテキストが直接重複している場合・空欄で次段落から補完した場合の
                                         *   両方に対応する）
                                         */
                                        if( isNewChapter_flg && textData.text == lastChapterTitle )
                                        {
                                            isNewChapter_flg = false;
                                            isDuplicateHeading_flg = true;

                                            /* 章番号カウンタが実際には消費されなかったことにするため、
                                             * engine.Generate呼び出し前の状態に巻き戻す
                                             */
                                            if( usedEngineGenerate_flg )
                                            {
                                                engine.RestoreState( engineStateBeforeGenerate );
                                            }
                                        }
                                    }
                                }

#if DEBUG
                                onLogUpdate?.Invoke( $"[DEBUG] idx={elementIndex} numId={numId} level={level} num=\"{numData.text}\" text=\"{textData.text}\" isNewChapter={isNewChapter_flg} isDuplicatedHeading={isDuplicateHeading_flg} lastChapterTitle=\"{lastChapterTitle}\"" );
#endif

                                /* 重複見出しと判定した段落は、章番号・タイトルとして採用しないだけでなく、
                                 * 本文としても出力せずに丸ごとスキップする
                                 * （Word上に紛れ込んだ空・重複見出し段落の残骸を出力してしまわないようにするため）
                                 */
                                if( isDuplicateHeading_flg )
                                {
                                    continue;
                                }

                                /* 箇条書きの場合は、記号を生成する
                                 * ・箇条書き番号が有効な場合：Wordの番号定義に基づいて実際の記号（①、a)、・ など）を生成する
                                 * ・箇条書き番号が無効な場合：固定で「・」を使用する
                                 */
                                string listMarker = "";
                                if( isList_flg &&
                                    numId.HasValue &&
                                    numberingMap.ContainsKey( numId.Value ) )
                                {
                                    if( a_outputListNumber_flg )
                                    {
                                        int levelValue = level.HasValue ? level.Value : 0;

                                        /* このnumId専用のカウンタを用意する（章番号用のengineとは状態を共有しない） */
                                        if( !listCounters.ContainsKey( numId.Value ) )
                                        {
                                            listCounters[numId.Value] = new Dictionary<int, int>();
                                        }

                                        listMarker = engine.GenerateListMarker( numberingMap[numId.Value], levelValue, listCounters[numId.Value] );
                                    }
                                    else
                                    {
                                        listMarker = "・";
                                    }
                                }

                                /* シートが未登録の場合、または章番号を取得した場合は新規シートを追加する */
                                if( null == wsPart )
                                {
                                    /* シートが未登録の場合 */

                                    /* 先頭シートを追加 */
                                    wsPart = ExcelHelper.CreateWorksheet( wbPart, sheets, sheetName, sheetId++, out sheetData );

                                    /* シートを新規作成したので画像貼付用の状態を初期化する */
                                    drawingsPart = null;
                                    imageId = 1;

                                    /* シートを新規作成したので空行判定の状態を初期化する */
                                    consecutiveBlankRows = 0;

                                    /* ログにシート名を表示 */
                                    onLogUpdate( sheetName );
                                }
                                else if( isNewChapter_flg )
                                {
                                    /* 章番号を取得した場合 */

                                    /* 今回採用した章タイトルを記憶する */
                                    lastChapterTitle = textData.text;

                                    /* 章番号 章タイトル のシートを追加 */
                                    sheetName = ExcelHelper.SafeSheetName( ( numData.text + " " + textData.text ).Trim() );
                                    wsPart = ExcelHelper.CreateWorksheet( wbPart, sheets, sheetName, sheetId++, out sheetData );

                                    /* シートを新規作成したので画像貼付用の状態を初期化する */
                                    drawingsPart = null;
                                    imageId = 1;

                                    /* シートを新規作成したので空行判定の状態を初期化する */
                                    consecutiveBlankRows = 0;

                                    /* ログにシート名を表示 */
                                    onLogUpdate( sheetName );

                                    /* シートが変わったので行を先頭に戻す */
                                    row = 1;
                                }

                                /* Wordファイル「画像」の処理
                                 * （画像を伴わない図形のみが複数の段落にまたがって離れた位置に配置されている場合、
                                 *   それらをまとめて1枚の画像として合成することがあるため、消費した段落数の分だけ
                                 *   ループのインデックスを進める）
                                 */
                                List<WordImageData> imageList;
                                if( a_outputImage_flg )
                                {
                                    imageList = WordImageHelper.GetImages( doc.MainDocumentPart, elements, elementIndex, out int consumedCount );
                                    elementIndex += consumedCount - 1;
                                }
                                else
                                {
                                    imageList = new List<WordImageData>();
                                }

                                foreach( WordImageData imageData in imageList )
                                {
                                    /* 画像を貼付け、画像の高さ分の行を確保する（0始まりの行番号を渡す） */
                                    int usedRows = ExcelHelper.AddImage( wsPart, ref drawingsPart, ref imageId, imageData, row - 1, IMAGE_COLUMN_INDEX );
                                    row += usedRows;
                                }

                                /* この段落の図形に設定されていたテキストを記録しておく（次の段落以降の重複判定に使用する）。
                                 * 図形が無かった場合は、直前までの記録をクリアして重複判定の対象範囲を限定する
                                 * （図形の直後に連続する段落だけを重複判定の対象とするため）
                                 */
                                List<string> shapeTexts = a_outputImage_flg ? ShapeOverlayCompositor.GetShapeTexts( para ) : new List<string>();
                                if( 0 < shapeTexts.Count )
                                {
                                    foreach( string shapeText in shapeTexts )
                                    {
                                        recentShapeTexts.Add( shapeText );
                                    }
                                }
                                else
                                {
                                    recentShapeTexts.Clear();
                                }

                                /* 章番号・テキスト・画像のいずれも無い行を「空行」とみなす */
                                bool isBlankRow = string.IsNullOrEmpty( numData.text ) &&
                                                   string.IsNullOrEmpty( textData.text ) &&
                                                   0 == imageList.Count;

                                /* 直前までの連続空行がすでに2行に達している場合は、3行以上連続しないようにこの行の出力をスキップする */
                                if( isBlankRow && 2 <= consecutiveBlankRows )
                                {
                                    continue;
                                }

                                consecutiveBlankRows = isBlankRow ? consecutiveBlankRows + 1 : 0;

                                /* 行出力（段落内改行(Shift+Enter)がある場合は、番号列を空白にして行を分けて出力する） */
                                string[] textLines = textData.text.Split( new[] { "\r\n", "\n" }, StringSplitOptions.None );

                                /* 箇条書きの記号が生成された場合は、記号列(B列)と内容列(C列)に分けて出力する */
                                bool hasListMarker_flg = !string.IsNullOrEmpty( listMarker );

                                /* 章の見出しとして扱われた行は、太字で表示する */
                                bool isHeadingRow_flg = isNewChapter_flg;

                                for( int i = 0; i < textLines.Length; i++ )
                                {
                                    /* 2行目以降は番号列を空白にする */
                                    CellData lineNumData = ( 0 == i ) ? numData : new CellData() { text = "" };
                                    CellData lineTextData;
                                    CellData lineContentData;

                                    if( hasListMarker_flg )
                                    {
                                        if( 0 == i )
                                        {
                                            /* 箇条書きの記号（①、a)、・ など）を右揃えでB列に、内容をC列に表示する */
                                            lineTextData = new CellData() { text = listMarker, rightAlign = true };
                                            lineContentData = new CellData() { text = textLines[i] };
                                        }
                                        else
                                        {
                                            /* 箇条書き項目内の改行による継続行：記号なしでC列にそのまま表示する */
                                            lineTextData = new CellData() { text = "" };
                                            lineContentData = new CellData() { text = textLines[i] };
                                        }
                                    }
                                    else
                                    {
                                        /* 箇条書きでない通常の行 */
                                        lineTextData = new CellData() { text = textLines[i] };
                                        lineContentData = new CellData() { text = "" };
                                    }

                                    /* 見出し行の場合は、行内のすべてのセルを太字にする */
                                    lineNumData.bold = isHeadingRow_flg;
                                    lineTextData.bold = isHeadingRow_flg;
                                    lineContentData.bold = isHeadingRow_flg;

                                    ExcelHelper.SetRow( wbPart, sheetData, row++, new List<CellData>() { lineNumData, lineTextData, lineContentData }, cache );
                                }

                                continue;
                            }

                            /* Wordファイル「表」の処理 */
                            Word.Table table = element as Word.Table;
                            if( null != table )
                            {
                                /* Excelワークシートを追加 */
                                if( null == wsPart )
                                {
                                    /* シートが未登録の場合 */

                                    /* 先頭シートを追加 */
                                    wsPart = ExcelHelper.CreateWorksheet( wbPart, sheets, sheetName, sheetId++, out sheetData );

                                    /* シートを新規作成したので画像貼付用の状態を初期化する */
                                    drawingsPart = null;
                                    imageId = 1;

                                    /* シートを新規作成したので空行判定の状態を初期化する */
                                    consecutiveBlankRows = 0;

                                    /* ログにシート名を表示 */
                                    onLogUpdate( sheetName );
                                }

                                ConvertTable( doc.MainDocumentPart, wbPart, wsPart, ref drawingsPart, ref imageId, table, sheetData, ref row, cache, a_outputImage_flg );

                                /* 表の後に区切りの空行を1行確保する（この行は空行として扱う） */
                                row++;
                                consecutiveBlankRows = 1;
                                continue;
                            }
                        }
                    }
                }
            }
            catch( Exception ex )
            {
                string errMsg = "エラーが発生しました" + Environment.NewLine + Environment.NewLine + ex.Message;
                Console.WriteLine( errMsg );
                onLogUpdate( errMsg );
                System.Windows.MessageBox.Show( errMsg );
            }
            finally
            {
                if( !a_token.IsCancellationRequested )
                {
                    onProgressUpdate?.Invoke( 100 );
                }

                /* 一時ファイルを削除 */
                try
                {
                    System.IO.File.Delete( tempPath );
                }
                catch( Exception ex )
                {
                    Console.WriteLine( ex.Message );
                    onLogUpdate( ex.Message );
                }
            }
        }


        /// <summary>
        /// Word の表を Excel の表形式データへ変換し、指定された SheetData に追記する。
        /// </summary>
        /// <param name="a_mainDocumentPart">MainDocumentPart（セル内画像の取得に使用）</param>
        /// <param name="a_wbPart">出力先 WorkbookPart</param>
        /// <param name="a_wsPart">出力先 WorksheetPart（セル内画像の貼付けに使用）</param>
        /// <param name="a_drawingsPart">画像貼付用のDrawingsPart（未生成の場合は生成される）</param>
        /// <param name="a_imageId">画像に付与するID（画像を貼付けるたびに更新される）</param>
        /// <param name="a_table">変換元の Word の表</param>
        /// <param name="a_sheetData">出力先 Excel シートデータ</param>
        /// <param name="a_row">Excel の出力開始行番号（出力後は次の行番号へ更新される）</param>
        /// <param name="a_outputImage_flg">画像を出力するか否か</param>
        private void ConvertTable( MainDocumentPart a_mainDocumentPart, WorkbookPart a_wbPart, WorksheetPart a_wsPart, ref DrawingsPart a_drawingsPart, ref uint a_imageId, Word.Table a_table, SheetData a_sheetData, ref int a_row, Dictionary<string, uint> a_cache, bool a_outputImage_flg )
        {
            /* -----------------------------------------------------------------
             * Word 表の全行を List にして index で参照できるようにする
             * ----------------------------------------------------------------- */
            List<Word.TableRow> rows = a_table.Elements<Word.TableRow>().ToList();

            /* -----------------------------------------------------------------
             * Word 表の各行を順番に処理する
             * ----------------------------------------------------------------- */
            for( int rowIndex = 0; rowIndex < rows.Count; rowIndex++ )
            {
                Word.TableRow tr = rows[rowIndex];

                /* -------------------------------------------------------------
                 * 1 行分の Excel セルデータを格納するリスト
                 * この values がそのまま Excel の 1 行になる
                 * ------------------------------------------------------------- */
                List<CellData> values = new List<CellData>();

                /* -------------------------------------------------------------
                 * 先頭列は章番号など別用途で使用するため、
                 * Word 表の内容は 1 列右にずらして出力する
                 * ------------------------------------------------------------- */
                values.Add( new CellData() { text = "" } );

                /* -------------------------------------------------------------
                 * この行のセル内に存在する画像を、貼付け先の列 index と対応付けて記録するリスト
                 * （行のテキストを出力した後、実際の貼付け処理を行う）
                 * ------------------------------------------------------------- */
                List<KeyValuePair<int, List<WordImageData>>> cellImages = new List<KeyValuePair<int, List<WordImageData>>>();

                /* -------------------------------------------------------------
                 * この行のセル一覧を取得（列 index 用）
                 * ------------------------------------------------------------- */
                List<Word.TableCell> cells = tr.Elements<Word.TableCell>().ToList();

                /* -------------------------------------------------------------
                 * Word 行内の各セルを順番に処理する
                 * ------------------------------------------------------------- */
                for( int colIndex = 0; colIndex < cells.Count; colIndex++ )
                {
                    Word.TableCell tc = cells[colIndex];

                    /* ---------------------------------------------------------
                     * セルのプロパティを取得
                     * （結合情報、罫線情報などが含まれる）
                     * --------------------------------------------------------- */
                    Word.TableCellProperties props = tc.TableCellProperties;

                    /* ---------------------------------------------------------
                     * GridSpan（横結合の列数）を取得
                     * --------------------------------------------------------- */
                    Word.GridSpan gridSpan = props?.GetFirstChild<Word.GridSpan>();
                    int span = ( null != gridSpan ) ? gridSpan.Val.Value : 1;

                    /* ---------------------------------------------------------
                     * VerticalMerge を取得
                     *
                     * VerticalMerge.Val
                     * ・MergedCellValues.Restart : 縦結合セルの開始セル
                     * ・MergedCellValues.Continue : 縦結合セルの継続セル
                     * --------------------------------------------------------- */
                    Word.VerticalMerge vertical = props?.GetFirstChild<Word.VerticalMerge>();
                    bool isRestart_flg = false;
                    bool isContinue_flg = false;
                    if( null != vertical )
                    {
                        if( null == vertical.Val )
                        {
                            /* 結合セル */
                            isContinue_flg = true;
                        }
                        else
                        {
                            if( Word.MergedCellValues.Restart == vertical.Val.Value )
                            {
                                /* 開始セル */
                                isRestart_flg = true;
                            }
                        }
                    }

                    /* ---------------------------------------------------------
                     * 次の行の同じ列のセルで縦結合が継続しているか判定
                     * --------------------------------------------------------- */
                    bool hasNextVerticalMerge_flg = false;

                    if( isContinue_flg &&
                        rowIndex + 1 < rows.Count )
                    {
                        Word.TableRow nextRow = rows[rowIndex + 1];
                        List<Word.TableCell> nextCells = nextRow.Elements<Word.TableCell>().ToList();

                        if( colIndex < nextCells.Count )
                        {
                            var nextVmerge = nextCells[colIndex].TableCellProperties?.GetFirstChild<Word.VerticalMerge>();
                            hasNextVerticalMerge_flg = ( null != nextVmerge );
                        }
                    }

                    /* ---------------------------------------------------------
                     * 枠線（下）判定
                     * ・非縦結合            ： あり
                     * ・Restart             ： なし
                     * ・Continue + 次もあり ： なし（縦結合セルの中間）
                     * ・Continue + 次はなし ： あり（縦結合セルの末尾）
                     * --------------------------------------------------------- */
                    bool bottomBorder_flg = ( null == vertical ) ? true : ( isContinue_flg ? !hasNextVerticalMerge_flg : false );

                    /* ---------------------------------------------------------
                     * セル内の画像を取得する（縦結合の継続セルはWord上テキストを持たないため対象外とする）
                     * この時点の values.Count が、このセルが配置される列 index（0始まり）と一致する
                     * --------------------------------------------------------- */
                    if( a_outputImage_flg &&
                        !isContinue_flg )
                    {
                        List<WordImageData> images = GetCellImages( a_mainDocumentPart, tc );

                        if( 0 < images.Count )
                        {
                            cellImages.Add( new KeyValuePair<int, List<WordImageData>>( values.Count, images ) );
                        }
                    }

                    /* ---------------------------------------------------------
                     * セルのテキストを取得する
                     * --------------------------------------------------------- */
                    string cellText = isContinue_flg ? "" : WordHelper.GetCellText( tc );

                    /* ---------------------------------------------------------
                     * セルデータを追加（先頭セル）
                     * --------------------------------------------------------- */
                    values.Add(
                        new CellData()
                        {
                            text = cellText,
                            topBorder = ( null == vertical ) ? true : isRestart_flg,
                            bottomBorder = bottomBorder_flg,
                            leftBorder = true,
                            rightBorder = ( 1 < span ) ? false : true
                        } );

                    /* ---------------------------------------------------------
                     * 横方向に結合されている残りの列数分、
                     * Excel 側で位置合わせ用の空セルを追加する
                     * 結合セルの末尾まで枠線（右）は設定しない
                     *
                     * 先頭セルはすでに追加済みなので i = 1 から開始
                     * --------------------------------------------------------- */
                    for( int i = 1; i < span; i++ )
                    {
                        values.Add(
                            new CellData()
                            {
                                text = "",
                                topBorder = ( null == vertical ) ? true : isRestart_flg,
                                bottomBorder = bottomBorder_flg,
                                leftBorder = false,
                                rightBorder = ( i == span - 1 ) ? true : false
                            } );
                    }
                }

                /* -------------------------------------------------------------
                 * 1 行分のセルデータを Excel に出力する。
                 *
                 * セル内に画像が無い行は、高さを指定せずExcelの自動調整に任せる
                 * （Excelは実際のフォントで折返しを計算するため、こちらの方が正確）
                 *
                 * セル内に画像がある行は、浮動画像はExcelの自動調整の対象外のため、
                 * 高さを明示的に指定する必要がある。その際、セル内のテキストを実フォントで
                 * 計測して見積もった高さ（オート調整相当）と、画像が必要とする高さを比較し、
                 * 大きい方を採用する（画像がある行だけの処理なので、コストの高い文字幅計測は
                 * この場合に限定して行う）
                 *
                 * 出力後、次の行番号へ進める
                 * ------------------------------------------------------------- */
                double? rowHeightPoints = null;

                if( 0 < cellImages.Count )
                {
                    double maxImageHeightPoints = 0;

                    foreach( KeyValuePair<int, List<WordImageData>> pair in cellImages )
                    {
                        foreach( WordImageData imageData in pair.Value )
                        {
                            double heightPoints = ExcelHelper.CalculateRowHeightForImage( imageData );

                            if( maxImageHeightPoints < heightPoints )
                            {
                                maxImageHeightPoints = heightPoints;
                            }
                        }
                    }

                    double maxTextHeightPoints = 0;

                    foreach( CellData cellData in values )
                    {
                        if( string.IsNullOrEmpty( cellData.text ) )
                        {
                            continue;
                        }

                        double heightPoints = ExcelHelper.EstimateRowHeightForText( cellData.text, 1 );

                        if( maxTextHeightPoints < heightPoints )
                        {
                            maxTextHeightPoints = heightPoints;
                        }
                    }

                    rowHeightPoints = Math.Max( maxTextHeightPoints, maxImageHeightPoints );
                }

                int currentRowNumber = a_row;
                ExcelHelper.SetRow( a_wbPart, a_sheetData, a_row++, values, a_cache, rowHeightPoints );

                /* -------------------------------------------------------------
                 * この行のセルに存在した画像を、対応する列位置に貼付ける
                 * （行の高さは上ですでに画像に合わせて拡張済みのため、追加の行確保は不要）
                 * ------------------------------------------------------------- */
                foreach( KeyValuePair<int, List<WordImageData>> pair in cellImages )
                {
                    foreach( WordImageData imageData in pair.Value )
                    {
                        ExcelHelper.AddImage( a_wsPart, ref a_drawingsPart, ref a_imageId, imageData, currentRowNumber - 1, pair.Key );
                    }
                }
            }
        }


        /// <summary>
        /// 表のセル内に存在する画像を取得する（セル内の各段落を対象とする）。
        /// </summary>
        /// <param name="a_mainDocumentPart">MainDocumentPart</param>
        /// <param name="a_cell">対象セル</param>
        /// <returns>画像情報一覧</returns>
        private List<WordImageData> GetCellImages( MainDocumentPart a_mainDocumentPart, Word.TableCell a_cell )
        {
            List<WordImageData> result = new List<WordImageData>();

            foreach( Word.Paragraph para in a_cell.Elements<Word.Paragraph>() )
            {
                result.AddRange( WordImageHelper.GetImages( a_mainDocumentPart, para ) );
            }

            return result;
        }


    }
}

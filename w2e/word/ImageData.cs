namespace w2e.word
{
    /// <summary>
    /// Word文書内の画像情報を保持するクラス
    /// </summary>
    public class ImageData
    {
        /* 保存ファイル名 */
        public string fileName { get; set; }

        /* ファイル拡張子 */
        public string extension { get; set; }

        /* 画像データ */
        public byte[] data { get; set; }

        /* 画像の幅（EMU） */
        public long widthEmu { get; set; }

        /* 画像の高さ（EMU） */
        public long heightEmu { get; set; }

        /* 画像の代替テキスト */
        public string altText { get; set; }

        /* Word内のRelationshipId */
        public string relationshipId { get; set; }
    }
}
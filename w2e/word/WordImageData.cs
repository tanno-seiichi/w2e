using System;

namespace w2e.word
{
    /// <summary>
    /// Word文書内の画像情報を保持するクラス
    /// </summary>
    public class WordImageData
    {
        /// <summary>
        /// 画像データを取得または設定する。
        /// </summary>
        public byte[] imageData { get; set; }

        /// <summary>
        /// 画像の種別を取得または設定する。
        /// </summary>
        public string contentType { get; set; }

        /// <summary>
        /// 画像の幅（EMU）を取得または設定する。
        /// </summary>
        public long widthEmu { get; set; }

        /// <summary>
        /// 画像の高さ（EMU）を取得または設定する。
        /// </summary>
        public long heightEmu { get; set; }

        /// <summary>
        /// 代替テキストを取得または設定する。
        /// </summary>
        public string altText { get; set; }

        /// <summary>
        /// RelationshipId を取得または設定する。
        /// </summary>
        public string relationshipId { get; set; }
    }
}
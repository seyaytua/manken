# PDF統合・変換ツール Pro

PySide6で作成されたPDF処理アプリケーション

## 機能

- 📚 PDF統合：複数のPDFを1つに統合
- 🖼️ 画像変換：PDFをJPEG/PNG画像に変換
- ✂️ PDF分割：PDFを1ページずつ分割
- 📦 PDF圧縮：PDFファイルを圧縮
- 🔄 PDF回転：選択したページを回転
- 📑 ページ抽出：特定のページを抽出
- 🔒 パスワード保護：PDFにパスワードを設定

## インストール

### 必要なもの
- Python 3.11以上
- Poppler（pdf2imageを使用するため）

### セットアップ

```bash
pip install -r requirements.txt
Popplerのインストール
macOS:

Copybrew install poppler
Windows: Poppler for Windowsをダウンロードして、binフォルダをPATHに追加

Linux:

Copysudo apt-get install poppler-utils
実行方法
Copypython main.py
Windows実行ファイル
Releasesページから実行ファイル（.exe）をダウンロードできます。

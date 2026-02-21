# 確定申告ツール

IT/AIツール開発の個人事業主向け帳簿作成ツール。

## 概要
- レシート撮影 → AI仕訳判定 → 弥生会計CSV出力
- Google Apps Script + Vue 3 + Tailwind CSS
- Gemini AI によるレシート解析・勘定科目判定
- スマホ操作に最適化したモバイルファーストWebアプリ

## ファイル構成
- `コード.js` - バックエンド（GAS）全ロジック
- `index.html` - フロントエンド SPA（Vue 3 + Tailwind CSS）
- `appsscript.json` - GAS マニフェスト
- `.clasp.json` - clasp デプロイ設定

## 必要なスクリプトプロパティ
- `GEMINI_API_KEY` - Gemini API キー
- `RECEIPT_FOLDER_ID` - レシート画像保存先の Google Drive フォルダ ID

## スプレッドシート
- シート名「経費データ」を使用（自動作成）

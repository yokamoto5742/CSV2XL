import datetime

import polars as pl

from services.csv_processor import dataframe_to_documents, process_csv_data
from services.document_repository import NewDocument

DOC_COL_IDX = 3  # 処理後の col_6_文書名(元G列)
DOCTOR_COL_IDX = 5  # 処理後の col_9_医師名(元J列)


class TestCsvProcessor:
    """CSV処理機能のテストクラス"""

    def create_test_dataframe(self):
        """テスト用のDataFrameを作成（実際のCSVの11列構造を模擬）"""
        return pl.DataFrame({
            "患者ID": [1, 2, 3, 4],
            "col1": ["A", "B", "C", "D"],
            "col2": ["E", "F", "G", "H"],
            "col3": ["I", "J", "K", "L"],
            "col4": ["M", "N", "O", "P"],
            "col5": ["Q", "R", "S", "T"],
            "文書名": ["検査結果 A *", "診断書 B", "処方箋 C", "除外文書*"],  # G列：スペースとアスタリスクを含むデータ
            "日付": ["20240101", "20240102", "20240103", "20240104"],
            "col8": ["削除予定", "削除予定", "削除予定", "削除予定"],  # I列（インデックス8）削除される
            "医師名": ["田中 医師 *", "除外 医師", "佐藤医師", "鈴木 医師"],  # J列：スペースとアスタリスクを含むデータ
            "col10": ["削除予定", "削除予定", "削除予定", "削除予定"]  # K列（インデックス10）削除される
        })

    def column_values(self, df: pl.DataFrame, index: int) -> list[str]:
        return df[df.columns[index]].to_list()

    def test_process_csv_data_drops_columns(self):
        """A～C列・I列・K列が削除され、残りの列が元の順で並ぶことのテスト"""
        result_df = process_csv_data(self.create_test_dataframe(), [], [])

        assert result_df.columns == [
            "col_3_col3", "col_4_col4", "col_5_col5", "col_6_文書名", "col_7_日付", "col_9_医師名"
        ]

    def test_process_csv_data_no_exclusions(self):
        """除外設定なしでのCSVデータ処理テスト"""
        result_df = process_csv_data(self.create_test_dataframe(), [], [])

        assert len(result_df) == 4  # 全ての行が残る
        doc_values = self.column_values(result_df, DOC_COL_IDX)
        doctor_values = self.column_values(result_df, DOCTOR_COL_IDX)
        # スペースとアスタリスクが除去されることを確認
        assert "除外文書" in doc_values
        assert "除外医師" in doctor_values
        assert "田中医師" in doctor_values
        assert "検査結果A" in doc_values

    def test_process_csv_data_with_doc_exclusions(self):
        """文書除外設定ありでのCSVデータ処理テスト"""
        result_df = process_csv_data(self.create_test_dataframe(), ["除外文書"], [])

        assert len(result_df) == 3
        assert "除外文書" not in self.column_values(result_df, DOC_COL_IDX)
        assert "除外医師" in self.column_values(result_df, DOCTOR_COL_IDX)  # 医師は残る

    def test_process_csv_data_with_doctor_exclusions(self):
        """医師除外設定ありでのCSVデータ処理テスト"""
        result_df = process_csv_data(self.create_test_dataframe(), [], ["除外医師"])

        assert len(result_df) == 3
        assert "除外文書" in self.column_values(result_df, DOC_COL_IDX)  # 文書は残る
        assert "除外医師" not in self.column_values(result_df, DOCTOR_COL_IDX)

    def test_process_csv_data_with_both_exclusions(self):
        """文書と医師両方の除外設定ありでのCSVデータ処理テスト"""
        result_df = process_csv_data(self.create_test_dataframe(), ["除外文書"], ["除外医師"])

        assert len(result_df) == 2
        assert "除外文書" not in self.column_values(result_df, DOC_COL_IDX)
        assert "除外医師" not in self.column_values(result_df, DOCTOR_COL_IDX)

    def test_process_csv_data_partial_match_exclusions(self):
        """部分一致での除外機能テスト"""
        result_df = process_csv_data(self.create_test_dataframe(), ["検査"], ["田中"])

        # "検査結果A"（スペース・*除去後）と"田中医師"の行が除外される
        assert len(result_df) == 3
        assert not any("検査" in doc for doc in self.column_values(result_df, DOC_COL_IDX))
        assert not any("田中" in doctor for doctor in self.column_values(result_df, DOCTOR_COL_IDX))

    def test_process_csv_data_multiple_exclusions(self):
        """複数の除外条件でのテスト"""
        result_df = process_csv_data(self.create_test_dataframe(), ["検査", "診断"], ["田中", "佐藤"])

        assert len(result_df) == 1
        assert "除外文書" in self.column_values(result_df, DOC_COL_IDX)
        assert "鈴木医師" in self.column_values(result_df, DOCTOR_COL_IDX)


class TestDataframeToDocuments:
    def test_converts_rows(self):
        """日付は YYYY-MM-DD、患者IDは数値の取り込み用の行にする"""
        df = pl.DataFrame({
            "預り日": [datetime.date(2025, 1, 5)],
            "患者ID": [12345],
            "患者名": ["山田　太郎"],
            "文書名": ["診断書"],
            "診療科": ["内科"],
            "医師名": ["田中"],
        })

        assert dataframe_to_documents(df) == [
            NewDocument('2025-01-05', 12345, '山田　太郎', '診断書', '内科', '田中')
        ]

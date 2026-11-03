from unittest.mock import patch, MagicMock

import polars as pl

from services.csv_processor import process_csv_data


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
    
    @patch('services.csv_processor.ConfigManager')
    def test_process_csv_data_drops_columns(self, mock_config_manager):
        """A～C列・I列・K列が削除され、残りの列が元の順で並ぶことのテスト"""
        mock_config_manager.return_value.get_exclude_docs.return_value = []
        mock_config_manager.return_value.get_exclude_doctors.return_value = []

        result_df = process_csv_data(self.create_test_dataframe())

        assert result_df.columns == [
            "col_3_col3", "col_4_col4", "col_5_col5", "col_6_文書名", "col_7_日付", "col_9_医師名"
        ]

    @patch('services.csv_processor.ConfigManager')
    def test_process_csv_data_no_exclusions(self, mock_config_manager):
        """除外設定なしでのCSVデータ処理テスト"""
        # ConfigManagerのモック設定
        mock_config = MagicMock()
        mock_config.get_exclude_docs.return_value = []
        mock_config.get_exclude_doctors.return_value = []
        mock_config_manager.return_value = mock_config
        
        # テストデータ作成
        df = self.create_test_dataframe()
        
        # 処理実行
        result_df = process_csv_data(df)
        
        # 結果確認（除外されていないこと）
        assert len(result_df) == 4  # 全ての行が残る
        # 処理後は文書名がcol_6_文書名（インデックス3）、医師名がcol_9_医師名（インデックス5）
        doc_col_idx = 3  # col_6_文書名（元G列）
        doctor_col_idx = 5  # col_9_医師名（元J列）
        # スペースとアスタリスクが除去されることを確認
        doc_values = result_df[result_df.columns[doc_col_idx]].to_list()
        doctor_values = result_df[result_df.columns[doctor_col_idx]].to_list()
        assert "除外文書" in doc_values  # *が除去された後
        assert "除外医師" in doctor_values  # スペースが除去された後
        # 具体的な除去が行われていることを確認
        assert "田中医師" in doctor_values  # スペースとアスタリスクが除去されている
        assert "検査結果A" in doc_values  # スペースとアスタリスクが除去されている
    
    @patch('services.csv_processor.ConfigManager')
    def test_process_csv_data_with_doc_exclusions(self, mock_config_manager):
        """文書除外設定ありでのCSVデータ処理テスト"""
        # ConfigManagerのモック設定
        mock_config = MagicMock()
        mock_config.get_exclude_docs.return_value = ["除外文書"]
        mock_config.get_exclude_doctors.return_value = []
        mock_config_manager.return_value = mock_config
        
        # テストデータ作成
        df = self.create_test_dataframe()
        
        # 処理実行
        result_df = process_csv_data(df)
        
        # 結果確認（除外文書の行が削除されていること）
        assert len(result_df) == 3
        doc_col_idx = 3  # col_6_文書名（元G列）
        doctor_col_idx = 5  # col_9_医師名（元J列）
        assert "除外文書" not in result_df[result_df.columns[doc_col_idx]].to_list()
        assert "除外医師" in result_df[result_df.columns[doctor_col_idx]].to_list()  # 医師は残る
    
    @patch('services.csv_processor.ConfigManager')
    def test_process_csv_data_with_doctor_exclusions(self, mock_config_manager):
        """医師除外設定ありでのCSVデータ処理テスト"""
        # ConfigManagerのモック設定
        mock_config = MagicMock()
        mock_config.get_exclude_docs.return_value = []
        mock_config.get_exclude_doctors.return_value = ["除外医師"]
        mock_config_manager.return_value = mock_config
        
        # テストデータ作成
        df = self.create_test_dataframe()
        
        # 処理実行
        result_df = process_csv_data(df)
        
        # 結果確認（除外医師の行が削除されていること）
        assert len(result_df) == 3
        doc_col_idx = 3  # col_6_文書名（元G列）
        doctor_col_idx = 5  # col_9_医師名（元J列）
        assert "除外文書" in result_df[result_df.columns[doc_col_idx]].to_list()  # 文書は残る
        assert "除外医師" not in result_df[result_df.columns[doctor_col_idx]].to_list()
    
    @patch('services.csv_processor.ConfigManager')
    def test_process_csv_data_with_both_exclusions(self, mock_config_manager):
        """文書と医師両方の除外設定ありでのCSVデータ処理テスト"""
        # ConfigManagerのモック設定
        mock_config = MagicMock()
        mock_config.get_exclude_docs.return_value = ["除外文書"]
        mock_config.get_exclude_doctors.return_value = ["除外医師"]
        mock_config_manager.return_value = mock_config
        
        # テストデータ作成
        df = self.create_test_dataframe()
        
        # 処理実行
        result_df = process_csv_data(df)
        
        # 結果確認（両方の除外条件に該当する行が削除されていること）
        assert len(result_df) == 2
        doc_col_idx = 3  # col_6_文書名（元G列）
        doctor_col_idx = 5  # col_9_医師名（元J列）
        assert "除外文書" not in result_df[result_df.columns[doc_col_idx]].to_list()
        assert "除外医師" not in result_df[result_df.columns[doctor_col_idx]].to_list()
    
    @patch('services.csv_processor.ConfigManager')
    def test_process_csv_data_partial_match_exclusions(self, mock_config_manager):
        """部分一致での除外機能テスト"""
        # ConfigManagerのモック設定
        mock_config = MagicMock()
        mock_config.get_exclude_docs.return_value = ["検査"]  # 「検査結果 A」に部分一致
        mock_config.get_exclude_doctors.return_value = ["田中"]  # 「田中医師」に部分一致
        mock_config_manager.return_value = mock_config
        
        # テストデータ作成
        df = self.create_test_dataframe()
        
        # 処理実行
        result_df = process_csv_data(df)
        
        # 結果確認（部分一致で除外されていること）
        assert len(result_df) == 3  # "検査結果A"（スペース・*除去後）と"田中医師"の行が除外される
        doc_col_idx = 3  # col_6_文書名（元G列）
        doctor_col_idx = 5  # col_9_医師名（元J列）
        doc_list = result_df[result_df.columns[doc_col_idx]].to_list()
        doctor_list = result_df[result_df.columns[doctor_col_idx]].to_list()
        
        # 「検査」を含む文書が除外されていることを確認
        assert not any("検査" in doc for doc in doc_list)
        # 「田中」を含む医師が除外されていることを確認
        assert not any("田中" in doctor for doctor in doctor_list)
    
    @patch('services.csv_processor.ConfigManager')
    def test_process_csv_data_multiple_exclusions(self, mock_config_manager):
        """複数の除外条件でのテスト"""
        # ConfigManagerのモック設定
        mock_config = MagicMock()
        mock_config.get_exclude_docs.return_value = ["検査", "診断"]
        mock_config.get_exclude_doctors.return_value = ["田中", "佐藤"]
        mock_config_manager.return_value = mock_config
        
        # テストデータ作成
        df = self.create_test_dataframe()
        
        # 処理実行
        result_df = process_csv_data(df)
        
        # 結果確認（複数の除外条件に該当する行が全て削除されていること）
        assert len(result_df) == 1  # 「処方箋C」と「佐藤医師」の組み合わせも除外される
        doc_col_idx = 3  # col_6_文書名（元G列）
        doctor_col_idx = 5  # col_9_医師名（元J列）
        doc_list = result_df[result_df.columns[doc_col_idx]].to_list()
        doctor_list = result_df[result_df.columns[doctor_col_idx]].to_list()
        
        # 除外条件に該当しない行のみ残っていることを確認
        assert "除外文書" in doc_list
        assert "鈴木医師" in doctor_list

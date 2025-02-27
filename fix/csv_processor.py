# csv_processor.py - CSVファイルの処理に関する機能
import polars as pl


class CSVProcessor:
    """CSVファイルの読み込みと処理を行うクラス"""
    
    def read_csv_with_encoding(self, file_path):
        """CSVファイルを適切なエンコーディングで読み込む"""
        encodings = ['shift-jis', 'utf-8']
        schema = {"患者ID": pl.Int64}

        for encoding in encodings:
            try:
                df = pl.read_csv(
                    file_path,
                    encoding=encoding,
                    separator=',',
                    skip_rows=3,
                    has_header=True,
                    infer_schema_length=0,
                    schema_overrides=schema
                )

                if len(df.columns) > 1:
                    print(f"エンコーディング {encoding} で正常に読み込みました（列数: {len(df.columns)}, 行数: {len(df)}）")
                    return df
            except Exception as e:
                print(f"{encoding}での読み込み試行中にエラー: {str(e)}")
                continue

        raise ValueError("CSVファイルの読み込みに失敗しました")

    def convert_date_format(self, df):
        """日付列の形式を変換"""
        try:
            date_col = df.columns[0]
            return df.with_columns([
                pl.col(date_col).str.strptime(pl.Date, format="%Y%m%d").alias(date_col)
            ])
        except Exception as e:
            print(f"日付変換中にエラーが発生しましたが、処理を継続します: {str(e)}")
            return df

    def process_csv_data(self, df, config):
        """CSVデータの処理: 列の選択、フィルタリング、書式設定"""
        try:
            # 列名を一意にする
            original_columns = df.columns
            unique_columns = [f"col_{i}_{col}" for i, col in enumerate(original_columns)]
            df = df.select([
                pl.col(old_name).alias(new_name)
                for old_name, new_name in zip(original_columns, unique_columns)
            ])

            # K列とI列を削除
            columns_to_keep = [i for i in range(len(df.columns)) if i not in [8, 10]]
            df = df.select([df.columns[i] for i in columns_to_keep])

            # A列からC列を削除
            df = df.select(df.columns[3:])

            df = self._apply_filters(df, config)
            df = self._clean_text_columns(df)

            return df

        except Exception as e:
            print(f"データ処理中にエラーが発生しました: {str(e)}")
            raise

    def _apply_filters(self, df, config):
        """設定に基づいてデータをフィルタリング"""
        exclude_docs = config.get_exclude_docs()
        exclude_doctors = config.get_exclude_doctors()

        # 除外設定によるフィルタリング
        if exclude_docs:
            filter_conditions = [~pl.col(df.columns[3]).str.contains(doc) for doc in exclude_docs]
            combined_filter = filter_conditions[0]
            for condition in filter_conditions[1:]:
                combined_filter = combined_filter & condition
            df = df.filter(combined_filter)

        if exclude_doctors:
            doctor_filter_conditions = [~pl.col(df.columns[5]).str.contains(doc) for doc in exclude_doctors]
            doctor_combined_filter = doctor_filter_conditions[0]
            for condition in doctor_filter_conditions[1:]:
                doctor_combined_filter = doctor_combined_filter & condition
            df = df.filter(doctor_combined_filter)
            
        return df

    def _clean_text_columns(self, df):
        """テキスト列のクリーニング"""
        # D列とF列のスペースと*を除去（4列目と6列目）
        return df.with_columns([
            pl.col(df.columns[3]).str.replace_all(r'[\s*]', ''),  # D列
            pl.col(df.columns[5]).str.replace_all(r'[\s*]', '')   # F列
        ])

    def convert_to_string_data(self, df):
        """DataFrameの全ての列を文字列に変換してリストに変換"""
        temp_df = df.select([pl.col('*').cast(pl.String)])
        return temp_df.to_numpy().tolist()

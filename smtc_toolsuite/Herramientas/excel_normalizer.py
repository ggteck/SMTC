import pandas as pd
import glob
from pathlib import Path


class ExcelNormalizer:
    """
    A utility to normalize tables across multiple Excel files and sheets
    based on a mapping DataFrame of raw headers to canonical column names.
    Expected mapping_df columns: ['file_name_like', 'sheet_name_like', 'std_name', 'other_name']
    """

    def __init__(self, mapping_df: pd.DataFrame):
        """
        Initialize the normalizer with a DataFrame that defines:
          file_name_like, sheet_name_like, std_name, other_name
        """
        # Validate required columns
        required = {'file_name_like', 'sheet_name_like', 'std_name', 'other_name'}
        if not required.issubset(mapping_df.columns):
            missing = required - set(mapping_df.columns)
            raise ValueError(f"Mapping DataFrame is missing columns: {missing}")
        self.map_df = mapping_df.copy().astype(str)

    def get_mapping_subset(self, file_path: str, sheet_name: str) -> pd.DataFrame:
        """
        Return the subset of mapping_df applicable to this file and sheet.
        """
        fname = Path(file_path).stem.lower()
        sname = sheet_name.lower()
        mask = (
            self.map_df['file_name_like'].str.lower().apply(lambda pat: pat in fname) &
            self.map_df['sheet_name_like'].str.lower().apply(lambda pat: pat in sname)
        )
        return self.map_df[mask]

    def extract_and_normalize_tables(self, file_path: str, sheet_name: str, header_map: dict) -> pd.DataFrame:
        """
        Read the sheet without headers, locate each table by header rows,
        rename columns via header_map, drop unmapped columns/empty rows,
        and return the concatenated DataFrame of all tables.
        """
        # Read with no header so every row is preserved
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=None, dtype=str)
        df = df.fillna('')
        rows = df.values.tolist()

        # Find header rows by presence in header_map
        header_idxs = [i for i, row in enumerate(rows) if any(cell in header_map for cell in row)]
        blocks = []
        for start, end in zip(header_idxs, header_idxs[1:] + [len(rows)]):
            raw_header = rows[start]
            mapped_cols = [header_map.get(col, None) for col in raw_header]
            block = pd.DataFrame(rows[start+1:end], columns=mapped_cols)
            # Keep only mapped columns and drop rows that are blank in all mapped cols
            block = block.loc[:, block.columns.notna()].dropna(how='all')
            blocks.append(block)

        if blocks:
            return pd.concat(blocks, ignore_index=True)
        # Return empty frame with canonical columns
        canonical_cols = self.map_df['std_name'].unique().tolist()
        return pd.DataFrame(columns=canonical_cols)

    def normalize_folder(self, input_folder: str, sheet_names: list = None) -> tuple[pd.DataFrame, list]:
        """
        Walk through all .xlsx/.xlsm files in input_folder, apply normalization on
        each sheet (or only those in sheet_names if provided), and return:
        - a master DataFrame including a "file_name_like" column
        - a list of errors (file_path, error_message)
        """
        all_tables = []
        errors = []

        for file_path in glob.glob(str(Path(input_folder) / '*.xlsx')) + glob.glob(str(Path(input_folder) / '*.xlsm')):
            try:
                xls = pd.ExcelFile(file_path)
            except Exception as e:
                errors.append((file_path, str(e)))
                continue

            for sheet in xls.sheet_names:
                if sheet_names and sheet not in sheet_names:
                    continue
                subset = self.get_mapping_subset(file_path, sheet)
                if subset.empty:
                    continue

                header_map = dict(zip(subset['other_name'], subset['std_name']))
                df_norm = self.extract_and_normalize_tables(file_path, sheet, header_map)
                if df_norm.empty:
                    continue

                # add file_name_like column
                file_like_vals = subset['file_name_like'].unique().tolist()
                df_norm['file_name_like'] = file_like_vals[0] if file_like_vals else ''
                all_tables.append(df_norm)

        if all_tables:
            df = pd.concat(all_tables, ignore_index=True)
        else:
            canonical_cols = list(self.map_df['std_name'].unique()) + ['file_name_like']
            df = pd.DataFrame(columns=canonical_cols)

        return df, errors


def main():
    import argparse

    parser = argparse.ArgumentParser(
        description="Normalize Excel tables into a single CSV based on a mapping DataFrame"
    )
    parser.add_argument("--mapping", required=True,
                        help="Path to the mapping CSV (will be loaded into a DataFrame)")
    parser.add_argument("--input_dir", required=True, help="Folder with .xlsx files")
    parser.add_argument("--output", required=True, help="Output CSV path")
    parser.add_argument(
        "--sheets", nargs="*", help="Optional list of sheet names to include"
    )
    args = parser.parse_args()

    # Load mapping CSV into DataFrame and pass that to ExcelNormalizer
    mapping_df = pd.read_csv(args.mapping, dtype=str)
    normalizer = ExcelNormalizer(mapping_df)
    master_df = normalizer.normalize_folder(args.input_dir, args.sheets)
    master_df.to_csv(args.output, index=False)
    print(f"Master CSV written to: {args.output}")


if __name__ == "__main__":
    main()

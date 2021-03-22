import pandas as pd

class CloseOutStatement:
    """
    Object containing close out statement information pertaining to a settlement.
    
    ASSUMPTIONS:
    ------------
    - Item names always reside in column 1 (Col B)
    - Numerical values for each item resides in column 9 (Col J) of the corresponding row
    - Sub-total labels reside in column 2 (Col C)
    - Only 3 sections in closeout statement in following order: Settlement info, Expense, Medical/Lien

    """
    def __init__(self, file_path: str):
        self.file_path = file_path
        self.closeout_df = self.import_closeout_as_df(file_path)
        self.preprocess_closeout_df()

    
    def import_closeout_as_df(self, file_path: str):
        """
        Import closeout statement excel file as a pandas DataFrame
        """
        df = pd.read_excel(file_path, engine='openpyxl', header=None)
        return df

    def preprocess_closeout_df(self):
        """
        Preprocess closeout_df by performing the following operations:
            1. Drop unessential columns
            2. Combine separate item columns together (to get sub-total items into same column as other items)
            3. Rename columns
            4. Remove empty rows
            5. Format all strings
        """
        self._drop_columns_closeout()
        self._combine_item_columns_closeout()
        self._rename_columns_closeout()
        self._drop_empty_rows_closeout()
        self._drop_verbose_rows_closeout()
        self._format_item_entries_closeout()
    
    def _drop_columns_closeout(self):
        """
        Remove unessential columns from closeout_df 
        """
        self.closeout_df = self.closeout_df[[1,2,9]]

    def _combine_item_columns_closeout(self):
        """
        Combine separate item columns together (to get sub-total items into same column as other items)
        """
        self.closeout_df = self.closeout_df.combine_first(self.closeout_df[[2]].rename(columns={2: 1})).drop(columns=[2])
    
    def _rename_columns_closeout(self):
        """
        Rename columns of closeout_df to "item" and "amount"
        """
        self.closeout_df = self.closeout_df.rename(columns={1: 'item', 9: 'amount'})

    def _drop_empty_rows_closeout(self):
        """
        Drop empty rows from closeout_df and reset the indices to be contiguous
        """
        self.closeout_df = self.closeout_df.dropna(how='all', axis=0).reset_index(drop=True)

    def _drop_verbose_rows_closeout(self, len_thresh = 150):
        """
        Drop rows where item column is verbose (> 150 characters)
        """
        len_mask = (self.closeout_df.astype(str).applymap(len) > len_thresh).any(axis=1)
        self.closeout_df = self.closeout_df[~len_mask].reset_index(drop=True)

    def _format_item_entries_closeout(self):
        """
        Format entries in closeout_df
        """
        self.closeout_df['item'] = self.closeout_df['item'].apply(self._format_items)
    
    def _format_items(self, item: str):
        """
        Method for formatting a single string entry in "item" column of closeout_df. Meant to be used with pandas.DataFrame.apply
        """
        item = item.lower()
        item = item.replace(":", '')
        item = item.replace("-", '')
        return item


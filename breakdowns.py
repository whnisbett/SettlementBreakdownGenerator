import pandas as pd
import regex

class CloseOutStatement:
    """
    Object containing close out statement information pertaining to a settlement.
    
    ASSUMPTIONS:
    ------------
    - Item names always reside in column 1 (Col B)
    - Numerical values for each item resides in column 9 (Col J) of the corresponding row
    - Sub-total labels reside in column 2 (Col C)
    - Only 3 sections in closeout statement in following order: Settlement info, Expense, Medical/Lien
    - Net to client is a row (and is unique)
    - Total expenses is a row (and is unique
    - Total medical is a row (and is unique)
    - "Amount of Settlement" is a row (and is unique
    - Medical items have one extra row above ("Lien" or "Medical" header) and three extra rows below ("Total Medical", "Total Medical Reductions", and "Subtotal Medical")
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
        item = item.strip()
        return item

    def parse_closeout_df(self):
        """
        Parse information in closeout_df and store key information as attributes
        """
        self.client_name = self.get_client_name()
        self.settlement_amount = self.get_settlement_amount()
        self.net_to_client_amount = self.get_net_to_client_amount()
        self.total_expenses_amount = self.get_total_expenses()
        self.total_medical_amount = self.get_total_medical()
        self.medical_items = self.get_medical_items()


    def split_closeout_df(self):
        """
        Split the closeout df into 3 dfs based on the items items each section contains: settlement info, expenses, and medical
        """
        subtotal_mask = self._fuzzy_match_series(self.closeout_df['item'], 'subtotal', errors=2)
        subtotal_idxs = self.closeout_df[subtotal_mask].index
        
        settlement_df = self.closeout_df.iloc[:subtotal_idxs[0] + 1].reset_index(drop=True)
        expenses_df = self.closeout_df.iloc[subtotal_idxs[0] + 1: subtotal_idxs[1] + 1].reset_index(drop=True)
        medical_df = self.closeout_df.iloc[subtotal_idxs[1] + 1:subtotal_idxs[2] + 1].reset_index(drop=True)

        return settlement_df, expenses_df, medical_df

    def get_medical_items(self):
        """
        Get itemized medical expenses from closeout_df
        """
        _, _, medical_df= self.split_closeout_df()
        medical_items = medical_df.iloc[1:-3].reset_index(drop=True)
        return medical_items
    
    def get_settlement_amount(self):
        """
        Get settlement amount from closeout_df
        """
        settlement_amount_mask = self._fuzzy_match_series(self.closeout_df['item'], 'amount of settlement', errors=3)
        settlement_amount = self.closeout_df[settlement_amount_mask]['amount'].values[0]
        return settlement_amount

    def get_net_to_client_amount(self):
        """
        Get net amount to client from closeout_df
        """
        net_to_client_mask = self._fuzzy_match_series(self.closeout_df['item'], 'net to client', errors=3)
        net_to_client_amount = self.closeout_df[net_to_client_mask]['amount'].values[0]
        return net_to_client_amount

    def get_total_expenses(self):
        """
        Get total amount of all expenses in closeout_df
        """
        total_expenses_mask = self._fuzzy_match_series(self.closeout_df['item'], 'total expenses', errors=3)
        total_expenses = self.closeout_df[total_expenses_mask]['amount'].values[0]
        return total_expenses

    def get_total_medical(self):
        """
        Get total medical medical expenses from closeout_df
        """
        total_medical_mask = self._fuzzy_match_series(self.closeout_df['item'], 'total medical', errors=3)
        total_medical = self.closeout_df[total_medical_mask]['amount'].values[0]
        return total_medical

    def get_client_name(self):
        """
        Get client name from closeout_df
        """
        client_name_mask = self._fuzzy_match_series(self.closeout_df['item'], 'name', errors=1)
        name = self.closeout_df[client_name_mask]['item'].values[0]
        name = name.replace('name', '').strip()
        return name

    def _fuzzy_match_series(self, series: pd.Series, match: str, errors=3):
        """
        Find elements in series that are similar to pattern within some number of errors. Output is a mask of shape series.shape
        """
        series = series.astype(str)
        mask = series.apply(lambda item: self._is_fuzzy_match(item, match=match, errors=errors))
        return mask
        
    def _is_fuzzy_match(self, item: str, match: str, errors=3):
        """
        Returns whether item matches pattern within some number of errors
        """ 
        regex_pattern = f"({match}){{e<={errors}}}"
        matches = regex.findall(regex_pattern, item)
        return len(matches) > 0
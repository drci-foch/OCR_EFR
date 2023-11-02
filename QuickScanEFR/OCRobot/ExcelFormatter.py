import pandas as pd


class DataReshaper:
    def __init__(self, df: pd.DataFrame):
        self.df = df

    def reshape(self) -> pd.DataFrame:
        reshaped_dataframes = []
        date_indices = self.df[self.df['Paramètres']
                               == 'Date test'].index.tolist()
        for start_idx in date_indices:
            end_idx = date_indices[date_indices.index(
                start_idx) + 1] if date_indices.index(start_idx) + 1 < len(date_indices) else len(self.df)
            slice_df = self.df.iloc[start_idx:end_idx].copy()
            date = slice_df.loc[slice_df['Paramètres']
                                == 'Date test', 'Pré'].values[0]
            slice_pivot = slice_df.pivot_table(
                index='Paramètres', values='Pré', aggfunc='first').reset_index()
            reshaped_slice = pd.DataFrame({'Date': [date]})
            for col in slice_pivot['Paramètres']:
                reshaped_slice[col] = slice_pivot.loc[slice_pivot['Paramètres']
                                                      == col, 'Pré'].values[0]
            if 'Date test' in reshaped_slice:
                reshaped_slice = reshaped_slice.drop(columns=['Date test'])
            if 'Heure test' in reshaped_slice:
                reshaped_slice = reshaped_slice.drop(columns=['Heure test'])
            reshaped_dataframes.append(reshaped_slice)
        final_df = pd.concat(reshaped_dataframes, ignore_index=True)
        return final_df

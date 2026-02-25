import pandas as pd
import numpy as np
from colorama import Fore, Style

try:
    from sklearn.tree import DecisionTreeClassifier
    from sklearn.preprocessing import LabelEncoder
    SKLEARN_AVAIL = True
except ImportError:
    SKLEARN_AVAIL = False

class ValidatorAnalyzer:
    def __init__(self, ignore_cols=None):
        if ignore_cols is None:
            self.ignore_cols = ['CONO', 'MESSAGE', 'RGDT', 'LMDT', 'RGTM', 'CHID', 'LMTS']
        else:
            self.ignore_cols = ignore_cols

    def analyze_column_pair(self, df_legacy, col_leg, df_m3, col_m3):
        """
        Compares two columns and returns a dictionary of suggestions.
        Used by unit tests.
        """
        # Ensure Series
        s_leg = df_legacy[col_leg].astype(str)
        s_m3 = df_m3[col_m3].astype(str)
        
        # Check if M3 is Constant
        if s_m3.nunique() == 1:
            val = s_m3.iloc[0]
            return {'TYPE': 'CONST', 'LOGIC': val, 'CONFIDENCE': '100%'}
            
        # Check Direct Match
        if s_leg.equals(s_m3):
            return {'TYPE': 'DIRECT', 'SOURCE': col_leg, 'CONFIDENCE': '100%'}
            
        return {'TYPE': 'TODO', 'LOGIC': 'Analysis inconclusive'}

    def _prepare_df(self, df, keep_keys=[]):
        df = df.copy()
        cols_to_drop = [c for c in self.ignore_cols if c in df.columns and c not in keep_keys]
        if cols_to_drop:
            df.drop(columns=cols_to_drop, inplace=True)
        return df

    def _explain_deviation(self, df_source, y_target, dominant_val):
        if not SKLEARN_AVAIL: return None
        y_binary = (y_target != dominant_val).astype(int)
        if y_binary.sum() < 2: return None

        valid_predictors = []
        df_encoded = pd.DataFrame()
        
        for col in df_source.columns:
            if df_source[col].nunique() < 20:
                try:
                    le = LabelEncoder()
                    df_encoded[col] = le.fit_transform(df_source[col].astype(str))
                    valid_predictors.append(col)
                except: pass
        
        if not valid_predictors: return None
        
        clf = DecisionTreeClassifier(max_depth=2)
        clf.fit(df_encoded[valid_predictors], y_binary)
        
        # Simple extraction of top feature
        importances = clf.feature_importances_
        if importances.max() > 0.5:
            top_idx = importances.argmax()
            return f"Depends on {valid_predictors[top_idx]}"
        return None
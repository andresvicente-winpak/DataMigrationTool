import pandas as pd
import numpy as np
from sklearn.tree import DecisionTreeClassifier, _tree
from sklearn.preprocessing import LabelEncoder
from colorama import Fore, Style

class PatternHunter:
    def __init__(self, df_combined, src_cols, tgt_cols):
        self.df = df_combined
        self.src_cols = src_cols
        self.tgt_cols = tgt_cols
        
        # Pre-encode sources for Decision Tree
        self.df_encoded = pd.DataFrame()
        self.encoders = {}
        self.valid_predictors = []

        print("   -> Pre-processing data for AI...")
        
        for col in self.src_cols:
            # Ignore High Cardinality (IDs) if > 80% unique
            # Exception: if total rows are small (<20), allow it
            if len(self.df) > 20 and self.df[col].nunique() > (len(self.df) * 0.8): 
                continue
            
            series = self.df[col].astype(str).fillna("NULL")
            le = LabelEncoder()
            self.df_encoded[col] = le.fit_transform(series)
            self.encoders[col] = le
            self.valid_predictors.append(col)

    def analyze_target(self, tgt_col):
        y = self.df[tgt_col].astype(str).fillna("")
        counts = y.value_counts(normalize=True)
        if counts.empty: return None
        
        top_val = counts.index[0]
        top_pct = counts.iloc[0] * 100

        # 1. CONSTANT
        if top_pct == 100.0:
            return {"Type": "CONST", "Logic": f"Value is '{top_val}'", "Prob": 100.0}

        # 2. DIRECT COPY
        best_src = None; best_score = 0
        for src in self.src_cols:
            match_pct = (self.df[src].astype(str).fillna("") == y).mean() * 100
            if match_pct > best_score:
                best_score = match_pct
                best_src = src
        
        if best_score == 100.0:
            clean = best_src.replace('_SRC', '')
            return {"Type": "DIRECT", "Logic": f"Copy {clean}", "Prob": 100.0}

        # 3. LOGIC
        if top_pct > 80.0 and self.valid_predictors:
            reason = self._explain(y, top_val)
            if reason:
                return {"Type": "LOGIC", "Logic": f"Mostly '{top_val}'. {reason}", "Prob": top_pct}
            else:
                return {"Type": "CONST?", "Logic": f"Mostly '{top_val}'. No clear pattern for outliers.", "Prob": top_pct}

        return {"Type": "UNK", "Logic": "Variable Data", "Prob": top_pct}

    def _explain(self, y, top_val):
        y_binary = (y != top_val).astype(int)
        if y_binary.sum() < 2: return None

        # Ensure we have valid predictors
        if self.df_encoded.empty: return None

        clf = DecisionTreeClassifier(max_depth=1)
        try:
            clf.fit(self.df_encoded[self.valid_predictors], y_binary)
        except: return None
        
        if clf.tree_.node_count < 3: return None
        
        feat_idx = clf.tree_.feature[0]
        if feat_idx < 0: return None # No split found
        
        feat = self.valid_predictors[feat_idx]
        
        # Find correlation
        deviants = self.df[y_binary == 1]
        if deviants.empty: return None
        
        # FIX: Safety check for mode
        mode_result = deviants[feat].mode()
        if mode_result.empty:
            cause = "Unknown"
        else:
            cause = mode_result[0]
        
        clean_feat = feat.replace('_SRC', '')
        return f"Exceptions correlate with {clean_feat} == '{cause}'"
import pandas as pd
import numpy as np
from sklearn.tree import DecisionTreeClassifier, _tree
from sklearn.preprocessing import LabelEncoder
from colorama import Fore, Style

class PatternHunter:
    def __init__(self, df):
        self.df = df
        
        # Identify Source vs Target columns
        self.src_cols = [c for c in df.columns if c.endswith('_SRC') or (c in df.columns and not c.endswith('_TGT'))]
        self.tgt_cols = [c for c in df.columns if c.endswith('_TGT')]
        
        # Pre-filter "Noise" columns from Source
        # We don't want to use unique IDs or Descriptions as predictors
        self.valid_predictors = []
        print("[AI] Profiling source data to exclude noise...")
        
        for col in self.src_cols:
            # Check cardinality (number of unique values)
            unique_count = df[col].nunique()
            total_rows = len(df)
            
            # Rule: If column is > 50% unique (like an ID or Description), ignore it
            if unique_count > 1 and unique_count < (total_rows * 0.5):
                self.valid_predictors.append(col)
        
        print(f"[AI] Found {len(self.valid_predictors)} valid predictor columns (out of {len(self.src_cols)}).")

        # Encode Predictors for AI
        self.encoders = {}
        self.X_encoded = pd.DataFrame()
        
        for col in self.valid_predictors:
            le = LabelEncoder()
            # Fill NaNs with a placeholder string
            clean_col = df[col].fillna("<<EMPTY>>").astype(str)
            self.X_encoded[col] = le.fit_transform(clean_col)
            self.encoders[col] = le

    def analyze_target(self, tgt_col):
        # Get Target Data
        y_raw = self.df[tgt_col].fillna("").astype(str)
        total = len(y_raw)
        
        # 1. Check Constant
        counts = y_raw.value_counts()
        top_val = counts.index[0]
        top_count = counts.iloc[0]
        pct = (top_count / total) * 100
        
        # Case A: 100% Constant
        if pct == 100.0:
            return {
                "Type": "CONST",
                "Logic": f"Set to '{top_val}'",
                "Prob": 100.0
            }

        # Case B: Direct Map (100%)
        # Check against original source columns
        for src in self.src_cols:
            # Simple equality check
            src_data = self.df[src].fillna("").astype(str)
            if (src_data == y_raw).all():
                clean_src = src.replace('_SRC', '')
                return {
                    "Type": "DIRECT",
                    "Logic": f"Map from {clean_src}",
                    "Prob": 100.0
                }

        # Case C: High Probability (>90%) - Find the Exception Pattern
        if pct >= 90.0:
            # We have a dominant value, but some exceptions.
            # Let's ask the Decision Tree: "What makes the other rows different?"
            
            # 0 = Dominant, 1 = Exception
            y_binary = (y_raw != top_val).astype(int)
            
            # Train small tree (depth 2 is usually enough for human-readable rules)
            clf = DecisionTreeClassifier(max_depth=2, min_samples_leaf=1)
            clf.fit(self.X_encoded, y_binary)
            
            # Extract Logic
            rules = self._extract_tree_rules(clf, self.valid_predictors)
            
            if rules:
                logic = f"Mostly '{top_val}'. EXCEPT: {rules[0]}"
            else:
                logic = f"Mostly '{top_val}'. (No clear pattern for exceptions found)"
                
            return {
                "Type": "LOGIC",
                "Logic": logic,
                "Prob": pct
            }

        return {
            "Type": "UNKNOWN",
            "Logic": "Data too variable (requires Map or complex logic)",
            "Prob": pct
        }

    def _extract_tree_rules(self, tree, feature_names):
        """
        Converts the decision tree path into a human readable string.
        """
        tree_ = tree.tree_
        feature_name = [
            feature_names[i] if i != _tree.TREE_UNDEFINED else "undefined!"
            for i in tree_.feature
        ]
        
        # We only care about paths that lead to Class 1 (The Exception)
        # This is a simplified extractor that grabs the strongest split
        
        if tree_.node_count < 3: return []
        
        # Root split
        node = 0
        feat = feature_name[node]
        thresh = tree_.threshold[node]
        
        # Decode threshold back to real value
        le = self.encoders[feat]
        # Find value closest to threshold (LabelEncoder maps to integers 0..N)
        val_idx = int(thresh)
        if val_idx < len(le.classes_):
            real_val = le.classes_[val_idx]
        else:
            real_val = "Unknown"
            
        # Check left (<=) and right (>) children to see where the 1s are
        # (Simple heuristic: just describe the primary split)
        
        clean_feat = feat.replace('_SRC', '')
        return [f"If {clean_feat} != '{real_val}'..."] 
        # Note: Real tree parsing is complex. 
        # For a POC, we often just correlate:
        # "Exception correlates with ITTY='900'"
        
        return []
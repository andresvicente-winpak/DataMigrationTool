import pandas as pd
import os
import configparser
from colorama import Fore, Style

# Try importing SQL libraries (fail gracefully if not installed)
try:
    from sqlalchemy import create_engine
    import urllib
    SQL_AVAILABLE = True
except ImportError:
    SQL_AVAILABLE = False

class DataExtractor:
    def __init__(self):
        self.db_config_path = 'config/db_config.ini'

    def _get_sql_connection(self):
        """
        Reads config/db_config.ini and returns a SQLAlchemy engine.
        """
        if not SQL_AVAILABLE:
            raise ImportError("SQLAlchemy/pyodbc not installed. Run: pip install sqlalchemy pyodbc")

        if not os.path.exists(self.db_config_path):
            raise FileNotFoundError(f"DB Config not found at {self.db_config_path}. Configure it in the Map Editor.")

        config = configparser.ConfigParser()
        config.read(self.db_config_path)

        if 'DEFAULT' not in config:
            raise ValueError("Invalid DB Config format.")

        settings = config['DEFAULT']
        driver = settings.get('Driver', 'ODBC Driver 17 for SQL Server')
        server = settings.get('Server', '.')
        database = settings.get('Database', '')
        trusted = settings.get('Trusted_Connection', 'yes')

        # Build Connection String
        params = urllib.parse.quote_plus(
            f"DRIVER={{{driver}}};SERVER={server};DATABASE={database};Trusted_Connection={trusted};"
        )
        
        # Create Engine
        engine = create_engine(f"mssql+pyodbc:///?odbc_connect={params}")
        return engine

    def load_data(self, file_path_or_query, format_type='MOVEX', sheet_name=0):
        """
        Loads data from Excel, CSV, or SQL.
        - file_path_or_query: Filename OR 'SQL:SELECT * FROM Table'
        """
        # --- 1. SQL QUERY HANDLING ---
        if str(file_path_or_query).upper().startswith("SQL:"):
            print(f"   -> Detected SQL Source: {file_path_or_query[:50]}...")
            query = file_path_or_query[4:].strip() # Remove 'SQL:' prefix
            
            try:
                engine = self._get_sql_connection()
                with engine.connect() as conn:
                    # Read SQL and force all columns to string (dtype=str not supported in read_sql directly)
                    df = pd.read_sql(query, conn)
                    # Convert all columns to string to match Excel behavior (avoid NaN/float issues)
                    df = df.astype(str)
                    
                    # Clean up 'None' strings that result from SQL NULLs converted to str
                    df.replace({'None': '', 'nan': ''}, inplace=True)
                    
                print(f"   -> Loaded {len(df)} rows from SQL.")
                return df
                
            except Exception as e:
                raise ValueError(f"SQL Error: {e}")

        # --- 2. FILE HANDLING ---
        if not os.path.exists(file_path_or_query):
            raise FileNotFoundError(f"File not found: {file_path_or_query}")

        print(f"Reading {format_type} file: {file_path_or_query} ({sheet_name})...")
        
        try:
            # CSV Handling
            if file_path_or_query.endswith('.csv'):
                df = pd.read_csv(file_path_or_query, dtype=str)
            else:
                # Excel Handling
                if format_type == 'M3_SDT':
                    df = pd.read_excel(file_path_or_query, sheet_name=sheet_name, header=0, dtype=str)
                    if len(df) > 2:
                        df = df.iloc[2:].reset_index(drop=True)
                    else:
                        df = pd.DataFrame(columns=df.columns)
                else:
                    # Standard Movex
                    df = pd.read_excel(file_path_or_query, sheet_name=sheet_name, dtype=str)

            # Cleanup
            df.columns = [str(c).strip() for c in df.columns]
            df.dropna(how='all', inplace=True)
            df.fillna("", inplace=True)
            
            print(f"   -> Loaded {len(df)} rows.")
            return df

        except Exception as e:
            raise ValueError(f"Error extracting data: {e}")

    def load_sdt_stitched(self, file_path, main_sheet, other_sheets=[]):
        """
        Loads the Main SDT sheet and merges additional sheets onto it.
        """
        print(f"   -> Loading Main Sheet: {main_sheet}")
        df_main = self.load_data(file_path, 'M3_SDT', main_sheet)
        
        for s in other_sheets:
            if not s or s == main_sheet: continue
            print(f"   -> Merging Sheet: {s}...")
            try:
                df_other = self.load_data(file_path, 'M3_SDT', s)
                
                # Find smart keys
                common = list(set(df_main.columns) & set(df_other.columns))
                keys = [k for k in ['CONO', 'DIVI', 'ITNO', 'CUNO', 'SUNO', 'FACI', 'WHLO'] if k in common]
                
                if keys:
                    df_main = pd.merge(df_main, df_other, on=keys, how='left', suffixes=('', f'_{s}'))
                else:
                    print(f"{Fore.YELLOW}      [Warning] No common keys found for {s}. Skipping merge.{Style.RESET_ALL}")
            except Exception as e:
                print(f"{Fore.RED}      Merge Error: {e}{Style.RESET_ALL}")
                
        return df_main
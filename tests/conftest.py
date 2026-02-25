import pytest
import pandas as pd
import os
import shutil

# Constants for Test Paths
TEST_CONFIG_DIR = "tests/config_temp"
TEST_DATA_DIR = "tests/data_temp"
TEST_OUTPUT_DIR = "tests/output_temp"
TEST_STAGE_DIR = "tests/surgical_staging"
TEST_UTIL_DIR = "tests/utilities_temp" # NEW

@pytest.fixture(scope="session", autouse=True)
def setup_test_env():
    """
    Runs once before all tests: Creates temp directories and dummy files.
    """
    # 1. Clean & Create Dirs
    for d in [TEST_CONFIG_DIR, TEST_DATA_DIR, TEST_OUTPUT_DIR, TEST_STAGE_DIR, TEST_UTIL_DIR]:
        if os.path.exists(d): shutil.rmtree(d)
        os.makedirs(d)
    
    # Create subfolders
    os.makedirs(f"{TEST_CONFIG_DIR}/rules")
    os.makedirs(f"{TEST_CONFIG_DIR}/sdt_templates")
    
    # Create Script subfolders
    os.makedirs(f"{TEST_UTIL_DIR}/TestCategory")

    # --- DATA GENERATION ---

    # A. Legacy Data (Item Master)
    # Added MMGR01 for Map Test
    df_legacy = pd.DataFrame({
        'MMITNO': ['1001', '1002', '1003'],
        'MMITDS': ['Item A', 'Item B', 'Item C'],
        'MMSTAT': ['20', '50', '20'],
        'MMGR01': ['G1', 'G2', 'G1'] 
    })
    df_legacy.to_excel(f"{TEST_DATA_DIR}/test_source.xlsx", index=False)

    # B. SDT Template
    df_sdt = pd.DataFrame(columns=['CONO', 'ITNO', 'ITDS', 'STAT', 'RESP', 'GR01', 'PRVG'])
    with pd.ExcelWriter(f"{TEST_CONFIG_DIR}/sdt_templates/TEST_API.xlsx", engine='xlsxwriter') as writer:
        df_sdt.to_excel(writer, sheet_name='AddBasic', startrow=3, header=False, index=False)
        ws = writer.sheets['AddBasic']
        for i, col in enumerate(df_sdt.columns): ws.write(0, i, col)
        
        df_sdt.to_excel(writer, sheet_name='AddWhs', startrow=3, header=False, index=False)
        ws2 = writer.sheets['AddWhs']
        for i, col in enumerate(df_sdt.columns): ws2.write(0, i, col)

    # C. Rule Config (Standard)
    # Added GR01 MAP rule
    # Added PRVG CONST 'NA' rule (Regression Test for NaN issue)
    df_rules = pd.DataFrame([
        {'TARGET_FIELD': 'ITNO', 'SOURCE_FIELD': 'MMITNO', 'RULE_TYPE': 'DIRECT', 'RULE_VALUE': '', 'SCOPE': 'GLOBAL'},
        {'TARGET_FIELD': 'ITDS', 'SOURCE_FIELD': 'MMITDS', 'RULE_TYPE': 'DIRECT', 'RULE_VALUE': '', 'SCOPE': 'GLOBAL'},
        {'TARGET_FIELD': 'STAT', 'SOURCE_FIELD': '', 'RULE_TYPE': 'CONST', 'RULE_VALUE': '20', 'SCOPE': 'GLOBAL'},
        {'TARGET_FIELD': 'RESP', 'SOURCE_FIELD': '', 'RULE_TYPE': 'CONST', 'RULE_VALUE': 'TESTUSER', 'SCOPE': 'GLOBAL'},
        {'TARGET_FIELD': 'GR01', 'SOURCE_FIELD': 'MMGR01', 'RULE_TYPE': 'MAP', 'RULE_VALUE': 'ITEM_GRP_MAP.xlsx|LEGACY_GRP|M3_GRP', 'SCOPE': 'GLOBAL'},
        {'TARGET_FIELD': 'PRVG', 'SOURCE_FIELD': '', 'RULE_TYPE': 'CONST', 'RULE_VALUE': 'NA', 'SCOPE': 'GLOBAL'}, 
    ])
    df_rules.to_excel(f"{TEST_CONFIG_DIR}/rules/TEST_RULES.xlsx", sheet_name='Rules', index=False)

    # D. Scoped Rule Config (For Context Testing)
    df_scoped = pd.DataFrame([
        {'TARGET_FIELD': 'ITNO', 'SOURCE_FIELD': 'MMITNO', 'RULE_TYPE': 'DIRECT', 'RULE_VALUE': '', 'SCOPE': 'GLOBAL'},
        {'TARGET_FIELD': 'STAT', 'SOURCE_FIELD': '', 'RULE_TYPE': 'CONST', 'RULE_VALUE': '10', 'SCOPE': 'GLOBAL'},
        {'TARGET_FIELD': 'STAT', 'SOURCE_FIELD': '', 'RULE_TYPE': 'CONST', 'RULE_VALUE': '90', 'SCOPE': 'DIV_US'},
    ])
    df_scoped.to_excel(f"{TEST_CONFIG_DIR}/rules/TEST_SCOPED.xlsx", sheet_name='Rules', index=False)

    # E. Maps
    pd.DataFrame([{'OBJECT_TYPE': 'ITEM', 'MCO_SHEET': 'Test_Sheet'}]).to_csv(f"{TEST_CONFIG_DIR}/surgical_def.csv", index=False)
    pd.DataFrame([{'MCO_SHEET': 'Test_Sheet', 'SOURCE_FILE': f"{TEST_DATA_DIR}/test_source.xlsx", 'JOIN_KEY': 'MMITNO'}]).to_csv(f"{TEST_CONFIG_DIR}/source_map.csv", index=False)
    # Note: 'Test_Sheet' is the MCO_SHEET name.
    pd.DataFrame([{'MCO_SHEET': 'Test_Sheet', 'API_NAME': 'TEST_API', 'SDT_TEMPLATE': 'TEST_API.xlsx', 'TRANSACTION_SHEET': 'AddBasic, AddWhs'}]).to_csv(f"{TEST_CONFIG_DIR}/migration_map.csv", index=False)

    # F. MCO Files
    # 1. Valid MCO
    df_mco_valid = pd.DataFrame({
        'Field name': ['ITNO', 'ITDS', 'STAT'],
        'Data Converison Source': ['MMITNO', 'MMITDS', ''],
        'Customer Required': ['1', '1', '1'],
        'Transformation Rule': ['Direct', 'Direct', 'Const 20']
    })
    with pd.ExcelWriter(f"{TEST_DATA_DIR}/MCO_VALID.xlsx") as writer:
        df_mco_valid.to_excel(writer, sheet_name='Sheet1', startrow=2, index=False)

    # 2. Update MCO (For Merge Test)
    df_mco_update = pd.DataFrame({
        'Field name': ['ITNO'],
        'Data Converison Source': ['MMITNO_UPDATED'],
        'Customer Required': ['1'],
        'Transformation Rule': ['Direct']
    })
    with pd.ExcelWriter(f"{TEST_DATA_DIR}/MCO_UPDATE.xlsx") as writer:
        df_mco_update.to_excel(writer, sheet_name='Sheet1', startrow=2, index=False)

    # 3. Duplicate/Bad MCO (For Checker Test)
    df_mco_bad = pd.DataFrame({
        'Field name': ['ITNO', 'ITNO'],
        'Data Converison Source': ['MMITNO', 'MMITNO_BACKUP'],
        'Customer Required': ['1', '1']
    })
    with pd.ExcelWriter(f"{TEST_DATA_DIR}/MCO_BAD.xlsx") as writer:
        df_mco_bad.to_excel(writer, sheet_name='Sheet1', startrow=2, index=False)

    # G. External Lookup Map (New for Test 19)
    df_map = pd.DataFrame({
        'LEGACY_GRP': ['G1', 'G2'],
        'M3_GRP': ['GRP_001', 'GRP_002']
    })
    # Save it in the rules folder so it can be found
    df_map.to_excel(f"{TEST_CONFIG_DIR}/rules/ITEM_GRP_MAP.xlsx", index=False)

    # H. Dummy Script (For Test 25)
    script_path = f"{TEST_UTIL_DIR}/TestCategory/hello_world.py"
    with open(script_path, "w") as f:
        f.write("# DESCRIPTION: Prints hello\nprint('Hello World')")

    yield
import cx_Oracle
from datetime import datetime, timedelta
import pandas as pd
from o365_drive import O365Drive

oracle_db_user = "__oracle_db_user__"
oracle_db_pass = "__oracle_db_pass__"
oracle_db_dns = "__oracle_db_dns__"

if __name__=="__main__":
    df = None
    query = "SELECT * FROM AP_BI.T_CRM_TLS_CALL_LIST_PROD"
    with cx_Oracle.connect(oracle_db_user,oracle_db_pass,oracle_db_dns) as con:
        df = pd.read_sql_query(query, con)

    bi_sharepoint = O365Drive(
                    client_id="__client_id__", 
                    client_secret="__client__secrete", 
                    host_name="homecreditgroup.sharepoint.com",
                    path_to_site="/sites/BusinessIntelligence",
                    token_file_path="D:\\BI\o365_token.txt",
                    scopes=["basic", "onedrive_all", "sharepoint_dl"]
                    )

    excel_file_path = "/Data/AP_SALES/CONTRACTS.xlsx"
    sharepoint = bi_sharepoint.update_excel_data(df=df, excel_file_path=excel_file_path, sheet_name="Sheet1")

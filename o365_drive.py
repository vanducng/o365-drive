from datetime import datetime, timedelta
import pandas as pd
from O365 import Account, FileSystemTokenBackend
from requests.exceptions import HTTPError
from O365.excel import WorkBook
import os
import math


class O365Drive():
    account = None
    drive = None

    def __init__(self,
                 client_id,
                 client_secret,
                 host_name,
                 path_to_site,
                 token_file_path,
                 drive_type="sharepoint",
                 scopes=["basic", "onedrive_all", "sharepoint_dl"]
                 ):
        self.credentials = (client_id, client_secret)
        self.host_name = host_name
        self.path_to_site = path_to_site
        self.drive_type = drive_type
        self.token_file_path = token_file_path
        self.scopes = scopes

        # Configure account
        self.configure_account()

        # If the token is not available, create new one which requires consent process
        if(not os.path.exists(token_file_path)):
            print(
                "No token found, creating new one. Please do censent following the instruction: ")
            self.generate_token()
        else:
            self.authenticate()

        # Get drive
        self.drive = self.__get_drive()

    def __get_folder_from_path(self, folder_path):
        """
        Get path folder instance within drive.
        """
        if folder_path is None:
            return self.drive

        subfolders = folder_path.split('/')
        if len(subfolders) == 0:
            return self.drive

        items = self.drive.get_items()
        for subfolder in subfolders:
            try:
                subfolder_drive = list(
                    filter(lambda x: subfolder in x.name, items))[0]
                items = subfolder_drive.get_items()
            except:
                raise (f"Path {folder_path} not exist")

        return subfolder_drive

    def __file_is_exist(self, file_path):
        """
        Check if a file is existed or not
        """
        try:
            self.drive.get_item_by_path(file_path)
        except HTTPError as e:
            if e.response.json().get('error', {}).get('code', '') == "itemNotFound":
                return False
            else:
                raise e
        return True

    def __get_excel_file_instance(self, remote_file_path, sheet_name, create_new=True):
        """
        Get file instance from drive, create empty excel file if not existed vi create_new
        """
        # Remote excel path check
        remote_path, file_name = os.path.split(remote_file_path)

        # Generate empty excel file and upload if file not existed
        if not self.__file_is_exist(remote_file_path):
            local_file_path = f"{file_name}"
            pd.DataFrame().to_excel(local_file_path, index=False, sheet_name=sheet_name)

            self.upload_file(local_file_path, remote_path[1:])
            os.remove(local_file_path)

            print(f"{remote_file_path} not exists, empty excel file created.")

        return self.drive.get_item_by_path(remote_file_path)

    def __convert_header_name(self, n):
        """
        Convert column index to character for excel sheet
        """
        string = ""
        while n > 0:
            n, remainder = divmod(n - 1, 26)
            string = chr(65 + remainder) + string
        return string

    def __insert_data(self, ws_instance, data, col_start, col_end, row_start, row_end):
        """
        Insert data into specific excel range. Eg: A1:C5
        """
        range_insert = ws_instance.get_range(
            f"{col_start}{row_start}:{col_end}{row_end}")
        range_insert.values = data[row_start-1:row_end]
        range_insert.update()

    def __df_to_excel(self, df, ws_instance, chunk=1000):
        """
        Update in chunk to overcome the limit of maximum 4MB string of API send to MS Graph.
        Input data is pandas data frame and number of rows per chunk
        """

        data = [df.columns.values.tolist()] + df.values.tolist()
        n_col = df.shape[1]
        n_row = df.shape[0] + 1  # count header in
        n_iter = math.ceil(n_row/chunk)
        col_name = self.__convert_header_name(n_col)

        row_start = 1
        row_end = chunk

        # If total rows is less than chunk size then insert in one shot
        # Else insert by chunk
        if n_iter < 2:
            self.__insert_data(ws_instance,
                               data,
                               col_start="A",
                               col_end=col_name,
                               row_start=row_start,
                               row_end=n_row
                               )
        else:
            for i in range(n_iter):
                row_end = n_row if n_row < row_end else row_end
                self.__insert_data(ws_instance,
                                   data,
                                   col_start="A",
                                   col_end=col_name,
                                   row_start=row_start,
                                   row_end=row_end
                                   )
                row_start += chunk
                row_end += chunk

    def configure_account(self):
        """
        Account consist of client_id and client_secret taken from Registration App on Azure portal
        """
        token_backend = FileSystemTokenBackend(
            token_path=os.path.split(self.token_file_path)[0],
            token_filename=os.path.split(self.token_file_path)[1]
        )
        self.account = Account(self.credentials, token_backend=token_backend)

    def generate_token(self):
        """
        Generate new token at the first run or anytime permission update from Registration App account.
        Simply delete the token file specified at run time to trigger this function to run
        """
        if self.account.authenticate(scopes=self.scopes):
            print(
                f"Your account is authenticated. Token is saved at {self.token_file_path}")

    def authenticate(self):
        """
        Token activation for limited of time, at the time it got expired, 
        old one will be replace by one generated by refresh_token function
        """
        if not self.account.is_authenticated:
            print("New token generated as old one was expired!")
            self.account.connection.refresh_token()

        print("You are authenticated, token loaded.")

    def __get_drive(self):
        """
        Storage location can get from Sharepoint or OneDrive.
        """
        drive = None

        if self.drive_type == "sharepoint":
            # initiate share_point account
            share_point = self.account.sharepoint()

            # structure get site call with host_name and path_to_site.
            site = share_point.get_site(self.host_name, self.path_to_site)

            # Get the drive pointed to Document folder
            drive = site.get_default_document_library(True)
        else:
            raise ValueError(
                f"Drive type of {self.drive_type} is not supported yet")

        return drive

    def get_drive_id(self):
        return self.drive.object_id

    def get_root_folder_list(self):
        return self.drive.get_root_folder().get_child_folders()

    def upload_file(self, local_file_path, remote_path):
        """
        Upload file from local_path to remote_path.
        """
        folder = self.__get_folder_from_path(remote_path)
        folder.upload_file(local_file_path)

    def get_workbook_instance(self, excel_file_path, sheet_name="Sheet1"):
        excel_file_instance = self.__get_excel_file_instance(
            excel_file_path, sheet_name=sheet_name)
        wb_instance = WorkBook(excel_file_instance,
                               use_session=True, persist=True)

        return wb_instance

    def worksheet_is_exist(self, wb_instance, sheet_name):
        wss = self.get_worksheets(wb_instance)
        ws_names = [ws.name for ws in wss]

        if sheet_name in set(ws_names):
            return True
        else:
            return False

    def get_worksheet(self, wb_instance, sheet_name):
        wss = self.get_worksheets(wb_instance)
        for ws in wss:
            if ws.name == sheet_name:
                return ws
        return None

    def get_worksheets(self, wb_instance):
        return wb_instance.get_worksheets()

    def get_worsheet_count(self, wb_instance):
        return len(wb_instance.get_worksheets())

    def rename_worksheet(self, wb_instance, old_name, new_name):
        ws = self.get_worksheet(wb_instance, old_name)
        if ws is not None:
            ws.update(name=new_name)
        else:
            print(f"Sheet {old_name} not exist.")

    def blank_worksheet(self, wb_instance, sheet_name, create_new_ws=True):
        """
        Clean old data to prepare for new data load in. If worksheet not exist, create it.
        - create_new_ws=True: replace existing worksheet with new worksheet id
        - create_new_ws=False: delete used range data, old worksheet id is kept
        """
        ws = self.get_worksheet(wb_instance, sheet_name)
        if ws is not None:
            if create_new_ws:
                # Change worksheet name to temp one
                temp_name = sheet_name + "__temp__"
                ws.update(name=temp_name)

                # Create new worksheet
                self.create_worksheet(wb_instance, sheet_name)

                # Delete temp one
                self.delete_worksheet(wb_instance, temp_name)
            else:
                self.delete_ws_used_range(wb_instance, sheet_name)
        else:
            self.create_worksheet(wb_instance, sheet_name)

    def create_worksheet(self, wb_instance, sheet_name):
        if not self.worksheet_is_exist(wb_instance, sheet_name):
            wb_instance.add_worksheet(sheet_name)

    def delete_worksheet(self, wb_instance, sheet_name):
        ws = self.get_worksheet(wb_instance, sheet_name)
        if ws is not None and self.get_worsheet_count(wb_instance) > 1:
            wb_instance.delete_worksheet(ws.object_id)

    def delete_ws_used_range(self, wb_instance, sheet_name):
        ws_instance = wb_instance.get_worksheet(sheet_name)
        range_addr = ws_instance.get_used_range().address
        range_data = ws_instance.get_range(range_addr)
        range_data.clear()

    def auto_fit_columns(self, excel_file_path, sheet_name):
        # Get used range
        wb_instance = self.get_workbook_instance(excel_file_path, sheet_name)
        ws_instance = wb_instance.get_worksheet(sheet_name)
        range_addr = ws_instance.get_used_range().address
        used_range = ws_instance.get_range(range_addr)

        # Update format
        fmt = used_range.get_format()
        fmt.auto_fit_columns()
        fmt.update()

    def update_excel_data(self, df, excel_file_path, sheet_name):
        # Get excel worksheet
        wb_instance = self.get_workbook_instance(excel_file_path, sheet_name)
        self.blank_worksheet(wb_instance, sheet_name, create_new_ws=True)
        ws_instance = wb_instance.get_worksheet(sheet_name)

        # Reload new data
        df.fillna('', inplace=True)
        self.__df_to_excel(df, ws_instance)

        # Auto fit columns
        self.auto_fit_columns(excel_file_path, sheet_name)

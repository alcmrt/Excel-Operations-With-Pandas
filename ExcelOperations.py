import pandas as pd


def read_column_as_list(file_name, sheet_name, column_name):
    """Read a column as list from excel file and return."""
    work_book = pd.ExcelFile(file_name)
    data_frame = pd.read_excel(work_book, sheet_name)
    return data_frame[column_name].values.tolist()


def read_excel(file_name, sheet_name):
    """Read excel file and return as pandas data frame."""
    work_book = pd.ExcelFile(file_name)
    return pd.read_excel(work_book, sheet_name)


def write_to_excel(data_frame, new_file_name, sheet_name=None):
    """Write data frame to a new excel file."""
    new_file_name = new_file_name + ".xlsx"
    data_frame.to_excel(new_file_name, sheet_name=sheet_name, index=False)
    print("{} file created successfully".format(new_file_name))


def drop_duplicate_rows(data_frame, column_name):
    """Drop duplicate rows in given column in pandas data_frame"""
    df = data_frame.drop_duplicates(subset=column_name, keep="first")
    return df


def append_data_frames(data_frame1, data_frame2):
    """Append rows of data_frame2 the end of data_frame1"""
    return data_frame1.append(data_frame2)


def fetch_columns(data_frame, column_list):
    """Fetch given columns of data frame in column_list and return."""
    return data_frame[data_frame.columns.intersection(column_list)]
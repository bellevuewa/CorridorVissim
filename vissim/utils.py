from datetime import datetime


def seconds_after_midnight(current_time_str):
    # Parse the current time string into a datetime object
    current_time = datetime.strptime(current_time_str, "%H:%M:%S")
    # Calculate midnight of the next day
    midnight = current_time.replace(hour=0, minute=0, second=0, microsecond=0)
    # Calculate the difference in seconds
    seconds_until_midnight = (current_time - midnight).seconds
    return seconds_until_midnight


def get_row_num(df, string='Count Summaries - All Vehicles'):
    for i_row in range(len(df)):
        try:
            if df.iloc[i_row, :].str.contains(string, case=True, na=False).any():
                break 
        except:
            continue
    return i_row


def start_vissim():
    # COM-Server
    import win32com.client as com
    ## Connecting the COM Server => Open a new Vissim Window:
    Vissim = com.Dispatch("Vissim.Vissim")
    return Vissim
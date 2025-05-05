import pandas as pd
import candas as cd
import os
import can
import matplotlib.pyplot as plt

def list_blf_files(): # List all BLF files in the current directory and return a list of file names.
    files = [f for f in os.listdir('.') if f.endswith('.blf')]
    for i, file in enumerate(files, 1):
        print(f"{i}: {file}")
    return files

def read_and_filter(df): # This function reads the DataFrame and filters the rows based on the SystemState_xdu8 and SupplyVoltage_xdu16 columns.

    # Select relevant columns
    #df = df.iloc[:, [0, 1, 4]]  # Adjust based on your column positions
    
    df = df.astype({df.columns[0]: 'float64'})
    df.columns = ['Start Time', 'SystemState_xdu8', 'SupplyVoltage_xdu16']

    # Add here options to filter rows
    filtered_df = df[df['SystemState_xdu8'] != 0 ]
    filtered_df = df[df['SupplyVoltage_xdu16'] <= 18.0 ]
    filtered_df = df[df['SupplyVoltage_xdu16'] > 0 ]
    filtered_df = df

    return filtered_df

def add_time_columns(df): # In simple terms, this function calculates the duration of each state by grouping consecutive rows with the same state.
    df['End Time'] = df['Start Time'] # Initialize the End Time column with the Time[s] values
    df['Duration'] = 0.0 # Initialize the Duration column with zeros as a float

    i = 0 # Initialize the start index
    while i < len(df) - 1: # Loop through all rows except the last one, because looping through the last row will cause an error
        if df.loc[df.index[i], 'SystemState_xdu8'] == df.loc[df.index[i + 1], 'SystemState_xdu8']: # Compares if the state in the current row is the same as the state in the next row
            start_time = df.loc[df.index[i], 'Start Time'] # At the beginning of a state, get the start time
            state = df.loc[df.index[i], 'SystemState_xdu8'] # Get the state value for the table
            j = i + 1 # Initialize the end index to be used in the next loop
            while j < len(df) and df.loc[df.index[j], 'SystemState_xdu8'] == state: # Loop until the state changes or the end of the DataFrame is reached
                j += 1 # Increment the end index so that we can get the end time and calculate the duration
            end_time = df.loc[df.index[j - 1], 'Start Time'] # Get the end time
            duration = end_time - start_time # Calculate the duration
            df.loc[df.index[i], 'End Time'] = end_time # Update the end time in the table
            df.loc[df.index[i], 'Duration'] = duration # Update the duration in the table
            df = df.drop(df.index[i + 1:j]) # Remove the intermediate rows
            df = df.reset_index(drop=True) # Reset the index of the DataFrame,
        else:
            i += 1 # Move to the next row
    return df # Return the modified DataFrame

def add_off_time_periods(df): # This function adds off time periods to the result. An off time period is defined as a period of time greater than 2 seconds between two consecutive states.
    i = 0
    while i < len(df) - 1:  # Adjust the loop condition to avoid index out of range error
        if df.loc[df.index[i + 1], 'Start Time'] - df.loc[df.index[i], 'End Time'] > 2:  # Check if the duration is greater than 2 seconds
            start_time = df.loc[df.index[i], 'End Time'] # Get the end time of the current state as the start time of the off time period
            end_time = df.loc[df.index[i + 1], 'Start Time'] # Get the start time of the next state as the end time of the off time period
            duration = end_time - start_time # Calculate the duration of the off time period
            df = pd.concat([df, pd.DataFrame({'Start Time': [start_time], 'End Time': [end_time], 'Duration': [duration], 'SystemState_xdu8': ['Null'], 'SupplyVoltage_xdu16': ['Null']})], ignore_index=True)
            df = df.sort_values(by='Start Time') # Sort the DataFrame by the Start Time column
            df = df.reset_index(drop=True) # Reset the index of the DataFrame
            i += 1  # Increment i to avoid re-evaluating the same rows when the DataFrame is updated with the off time period
        i += 1 # Increment i to move to the next row
    return df

def add_cycle_count(result_df): # This function adds a new column to the DataFrame with a global cycle counter for every time SystemState_xdu8 = 5 & Duration > 800 seconds.
# It would show the cycle number only on the rows where SystemState_xdu8 = 5 & Duration > 890 seconds in a new column named 'Cycle Count'.
    cycle_count = int(0) # Initialize the cycle count
    for i in range(len(result_df)): # Loop through all rows in the DataFrame
        if result_df.loc[result_df.index[i], 'SystemState_xdu8'] == 5 and result_df.loc[result_df.index[i], 'Duration'] > 800: # Check if the state is 5 and the duration is greater than 800 seconds
            cycle_count += 1 # Increment the cycle count
        result_df.loc[result_df.index[i], 'Cycle Count'] = cycle_count # Update the cycle count in the DataFrame
    return result_df

def save_picture(df, file_name):
    plt.figure(figsize=(5.1, 1.2))
    plt.axis('off')
    plt.table(cellText=df.values, colLabels=df.columns, cellLoc='center')
    plt.suptitle(file_name, fontsize=10, y=0.45)  # Add a heading with the file name positioned on the top center of the image
    plt.savefig(file_name + ".jpeg", dpi=300, bbox_inches='tight', pad_inches=0.5)  # Increase the dpi for higher quality
    plt.close()

def save_to_excel(df, file_name): # Save the DataFrame to an Excel file, using the original CSV file name, without the extension and with the .xlsx extension.
    file_path = file_name + ".xlsx"
    if os.path.exists(file_path):
        overwrite = input(f"The file {file_path} already exists. Do you want to overwrite it? (y/n): ")
        if overwrite.lower() != 'y':
            print("File not saved.")
            return
    try:
        df.to_excel(file_path, index=False)
        print(f"Result saved to {file_path}")
    except PermissionError:
        print(f"Permission denied. The file {file_path} is being used by another process.")

def main():
    Names = ["ID2S09_sApplI_SystemState_xdu8", "ID2S09_uApplI_SupplyVoltage_xdu16"]
    dbc_folder = "dbc"
    signal_db = cd.load_dbc(dbc_folder)
    previous_cycle_count = 0
    i=int(0)

    files = list_blf_files()
    if not files:
        print("No BLF files found in the current directory.")
        return

    # Loop over all .blf files in the current directory
    for file in os.listdir():
        if file.endswith(".blf"):
            file_name = os.path.splitext(file)[0]
            blf_log = cd.from_file(signal_db, file_name, always_convert=True, verbose=True)
            df = blf_log.to_dataframe(names=Names)

            # Open the BLF blf_log file to set the timestamps.
            with can.BLFReader(file_name + '.blf') as blf_log:
                # Access the start timestamp from the BLFReader object
                min_timestamp = blf_log.start_timestamp # BLF File Start Timestamp
                max_timestamp = blf_log.stop_timestamp # BLF File End Timestamp
                print(f"Starting time (in seconds since epoch): {min_timestamp}")
                print(f"Ending time (in seconds since epoch): {max_timestamp}")
    
            # Rename first column to 'Time [s]'
            df.rename(columns={df.columns[0]: 'Time [s]'}, inplace=True)

            # Apply the transformation in a single line
            df['Time [s]'] = df['Time [s]'].apply(lambda timestamp: f"{timestamp - min_timestamp:.6f}")

            # Add an element to the last row of the df with all information empty except for the time, which would be just the 'max_timestamp', add using pd.concat
            df = pd.concat([df, pd.DataFrame({'Time [s]': [f"{max_timestamp - min_timestamp:.6f}"]})], ignore_index=True)

            filtered_df = read_and_filter(df)

            result_df = add_time_columns(filtered_df)

            result_df_with_off_times = add_off_time_periods(result_df) # Add off time periods to the result

            # Rearrange the columns: SystemState_xdu8, SupplyVoltage_xdu16, Start Time, End Time, Duration
            result_df_with_off_times = result_df_with_off_times[['SystemState_xdu8', 'Duration',  'Start Time', 'End Time', 'SupplyVoltage_xdu16']]

            # Add a new column to the DataFrame with a global cycle counter for everytime SystemState_xdu8 = 5 & Duration > 890 seconds.
            result_df_with_off_times = add_cycle_count(result_df_with_off_times)
            i+=1
            result_df_with_off_times['Cycle Count'] = result_df_with_off_times['Cycle Count'] + previous_cycle_count
            previous_cycle_count = int(result_df_with_off_times['Cycle Count'].max())

            # Change ID2S09_sApplI_SystemState_xdu8 DataFrame values according to description: 0 =	NULL, 1 =	NmWait, 2 =	OldKL30Wait, 3 =	PreDrive, 4 =	DriveDown, 5 =	DriveUp, 6 =	PostRun, 7 =	Off, 8 =	Error, 9 =	Flash, 10 =	LowVolt.
            result_df_with_off_times['SystemState_xdu8'] = result_df_with_off_times['SystemState_xdu8'].map(
            {0: 'NULL', 1: '1: NmWait', 2: '2: OldKL30Wait', 3: '3: PreDrive', 4: '4: DriveDown', 5: '5: DriveUp', 6: '6: PostRun',
             7: '7: Off', 8: '8: Error', 9: '9: Flash', 10: '10: LowVolt', 11: '11', 12: '12', 13: '13', 14: '14', 15: '15'})

            # Print the result
            print(result_df_with_off_times[['SystemState_xdu8', 'SupplyVoltage_xdu16', 'Start Time', 'Duration']])
        
            # Save the result as a PNG image
            save_picture(result_df_with_off_times, file_name)
    
            # Save to Excel with the same name as the file
            save_to_excel(result_df_with_off_times, file_name)

            # Save the modified DataFrame to a CSV file with the same name as the original file
            # df.to_csv(file_name + ".csv", index=False)

            # Delete the .mat file
            # os.remove(file_name + '.mat')

if __name__ == "__main__":
    main()
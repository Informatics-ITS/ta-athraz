import pandas as pd
import re

input_files = [
    '0fad4bfb7d4697eaf966b88c9e623f454f843700195bb192cfcadcd256a4c431\\w-user-2min.csv',
    '0fad4bfb7d4697eaf966b88c9e623f454f843700195bb192cfcadcd256a4c431\\wo-user-2min.csv',
    
    '63e9c56ece51abcf78da3653ed4b03355f36982fdca931043a4bcfca7caf145e\\w-user-2min.csv',
    '63e9c56ece51abcf78da3653ed4b03355f36982fdca931043a4bcfca7caf145e\\wo-user-2min.csv',
    
    '85a7d0f873bde902bb66bc481fce0fb3e377113fcf483f8e6d5aca7ca1196141\\w-user-2min.csv',
    '85a7d0f873bde902bb66bc481fce0fb3e377113fcf483f8e6d5aca7ca1196141\\wo-user-2min.csv',
    
    '989b2568cd870052182bd637a9a5a1bf68461e2fa59763678f88f634bfd1d51c\\w-user-2min.csv',
    '989b2568cd870052182bd637a9a5a1bf68461e2fa59763678f88f634bfd1d51c\\wo-user-2min.csv',

    '2820b924640b10eb028a6fb30a77db1d3c53077e368fdf204fa914306e1ecc3a\\w-user-2min.csv',
    '2820b924640b10eb028a6fb30a77db1d3c53077e368fdf204fa914306e1ecc3a\\wo-user-2min.csv',
    
    '83681a0a9a4e5ffcaa56761f4cd6881de408e3408ae579da2a4d006bec0f39c9\\w-user-2min.csv',
    '83681a0a9a4e5ffcaa56761f4cd6881de408e3408ae579da2a4d006bec0f39c9\\wo-user-2min.csv',
    
    '491960fbf4a3f4f085b1a318c88d7f39208150b2fd11d1251bf251e912691a0b\\w-user-2min.csv',
    '491960fbf4a3f4f085b1a318c88d7f39208150b2fd11d1251bf251e912691a0b\\wo-user-2min.csv',
    
    'ad9093633b9ecaeea7bff69ab8d8781213fec82db6c7f2e963a40d2e0ee0e9ce\\w-user-2min.csv',
    'ad9093633b9ecaeea7bff69ab8d8781213fec82db6c7f2e963a40d2e0ee0e9ce\\wo-user-2min.csv',
    
    'def98259bba7c128a22dbb9100a3e9512911d9775ec82175f8a8a3c26b993dbf\\w-user-2min.csv',
    'def98259bba7c128a22dbb9100a3e9512911d9775ec82175f8a8a3c26b993dbf\\wo-user-2min.csv',
    
    'e6dff8475541ebddc1f0db47a311eb2c25581b7d5e62af8066d59c283114c2d3\\w-user-2min.csv',
    'e6dff8475541ebddc1f0db47a311eb2c25581b7d5e62af8066d59c283114c2d3\\wo-user-2min.csv',
]

output_files = [
    '0fad4bfb7d4697eaf966b88c9e623f454f843700195bb192cfcadcd256a4c431\\w-user-2min-malware.csv',
    '0fad4bfb7d4697eaf966b88c9e623f454f843700195bb192cfcadcd256a4c431\\wo-user-2min-malware.csv',
    
    '63e9c56ece51abcf78da3653ed4b03355f36982fdca931043a4bcfca7caf145e\\w-user-2min-malware.csv',
    '63e9c56ece51abcf78da3653ed4b03355f36982fdca931043a4bcfca7caf145e\\wo-user-2min-malware.csv',
    
    '85a7d0f873bde902bb66bc481fce0fb3e377113fcf483f8e6d5aca7ca1196141\\w-user-2min-malware.csv',
    '85a7d0f873bde902bb66bc481fce0fb3e377113fcf483f8e6d5aca7ca1196141\\wo-user-2min-malware.csv',
    
    '989b2568cd870052182bd637a9a5a1bf68461e2fa59763678f88f634bfd1d51c\\w-user-2min-malware.csv',
    '989b2568cd870052182bd637a9a5a1bf68461e2fa59763678f88f634bfd1d51c\\wo-user-2min-malware.csv',

    '2820b924640b10eb028a6fb30a77db1d3c53077e368fdf204fa914306e1ecc3a\\w-user-2min-malware.csv',
    '2820b924640b10eb028a6fb30a77db1d3c53077e368fdf204fa914306e1ecc3a\\wo-user-2min-malware.csv',
    
    '83681a0a9a4e5ffcaa56761f4cd6881de408e3408ae579da2a4d006bec0f39c9\\w-user-2min-malware.csv',
    '83681a0a9a4e5ffcaa56761f4cd6881de408e3408ae579da2a4d006bec0f39c9\\wo-user-2min-malware.csv',
    
    '491960fbf4a3f4f085b1a318c88d7f39208150b2fd11d1251bf251e912691a0b\\w-user-2min-malware.csv',
    '491960fbf4a3f4f085b1a318c88d7f39208150b2fd11d1251bf251e912691a0b\\wo-user-2min-malware.csv',
    
    'ad9093633b9ecaeea7bff69ab8d8781213fec82db6c7f2e963a40d2e0ee0e9ce\\w-user-2min-malware.csv',
    'ad9093633b9ecaeea7bff69ab8d8781213fec82db6c7f2e963a40d2e0ee0e9ce\\wo-user-2min-malware.csv',
    
    'def98259bba7c128a22dbb9100a3e9512911d9775ec82175f8a8a3c26b993dbf\\w-user-2min-malware.csv',
    'def98259bba7c128a22dbb9100a3e9512911d9775ec82175f8a8a3c26b993dbf\\wo-user-2min-malware.csv',
    
    'e6dff8475541ebddc1f0db47a311eb2c25581b7d5e62af8066d59c283114c2d3\\w-user-2min-malware.csv',
    'e6dff8475541ebddc1f0db47a311eb2c25581b7d5e62af8066d59c283114c2d3\\wo-user-2min-malware.csv',
]

search_strings = [
    '0fad4bfb7d4697eaf966b88c9e623f454f843700195bb192cfcadcd256a4c431',
    '0fad4bfb7d4697eaf966b88c9e623f454f843700195bb192cfcadcd256a4c431',
    
    '63e9c56ece51abcf78da3653ed4b03355f36982fdca931043a4bcfca7caf145e',
    '63e9c56ece51abcf78da3653ed4b03355f36982fdca931043a4bcfca7caf145e',
    
    '85a7d0f873bde902bb66bc481fce0fb3e377113fcf483f8e6d5aca7ca1196141',
    '85a7d0f873bde902bb66bc481fce0fb3e377113fcf483f8e6d5aca7ca1196141',
    
    '989b2568cd870052182bd637a9a5a1bf68461e2fa59763678f88f634bfd1d51c',
    '989b2568cd870052182bd637a9a5a1bf68461e2fa59763678f88f634bfd1d51c',
    
    '2820b924640b10eb028a6fb30a77db1d3c53077e368fdf204fa914306e1ecc3a',
    '2820b924640b10eb028a6fb30a77db1d3c53077e368fdf204fa914306e1ecc3a',
    
    '83681a0a9a4e5ffcaa56761f4cd6881de408e3408ae579da2a4d006bec0f39c9',
    '83681a0a9a4e5ffcaa56761f4cd6881de408e3408ae579da2a4d006bec0f39c9',
    
    '491960fbf4a3f4f085b1a318c88d7f39208150b2fd11d1251bf251e912691a0b',
    '491960fbf4a3f4f085b1a318c88d7f39208150b2fd11d1251bf251e912691a0b',
    
    'ad9093633b9ecaeea7bff69ab8d8781213fec82db6c7f2e963a40d2e0ee0e9ce',
    'ad9093633b9ecaeea7bff69ab8d8781213fec82db6c7f2e963a40d2e0ee0e9ce',
    
    'def98259bba7c128a22dbb9100a3e9512911d9775ec82175f8a8a3c26b993dbf',
    'def98259bba7c128a22dbb9100a3e9512911d9775ec82175f8a8a3c26b993dbf',
    
    'e6dff8475541ebddc1f0db47a311eb2c25581b7d5e62af8066d59c283114c2d3',
    'e6dff8475541ebddc1f0db47a311eb2c25581b7d5e62af8066d59c283114c2d3',
]

columns_to_keep = [
    'EventId', 'MapDescription', 'UserName', 'PayloadData1', 'PayloadData2', 'PayloadData3',
    'PayloadData4', 'PayloadData5', 'PayloadData6', 'ExecutableInfo',
    'HiddenRecord', 'SourceFile', 'Keywords', 'ExtraDataOffset'
]

pid_regex = re.compile(r"ProcessID: (\d+)")

def extract_pids(df):
    event1_rows = df[df['EventId'] == 1]
    pid_record_map = {}
    for _, row in event1_rows.iterrows():
        val = row['PayloadData1']
        if pd.notna(val):
            match = pid_regex.search(val)
            if match:
                pid = match.group(1)
                record_number = row.get('RecordNumber', float('inf'))
                if pid not in pid_record_map or record_number < pid_record_map[pid]:
                    pid_record_map[pid] = record_number

    return pid_record_map

for input_file, output_file, search_string in zip(input_files, output_files, search_strings):
    df = pd.read_csv(input_file)

    initial_df = df[df.apply(lambda row: row.astype(str).str.contains(search_string).any(), axis=1)]
    all_related_rows = initial_df.copy()

    new_pids_map = extract_pids(initial_df)

    if new_pids_map:
        while True:
            new_rows = df[df.apply(
                lambda row: any(
                    (
                        (f"SourceProcessID: {pid}" in str(row) or
                         f"ParentProcessID: {pid}" in str(row) or
                         (f"ProcessID: {pid}" in str(row) and row.get('EventId') != 1))
                        and row.get('RecordNumber', float('-inf')) > rec_num
                    )
                    for pid, rec_num in new_pids_map.items()
                ),
                axis=1
            )]
            new_rows = new_rows[~new_rows.index.isin(all_related_rows.index)]
            if new_rows.empty:
                break

            all_related_rows = pd.concat([all_related_rows, new_rows], ignore_index=True)

            new_pids_map = extract_pids(new_rows)
            if not new_pids_map:
                break

    all_related_rows = all_related_rows.sort_values(by='TimeCreated')
    selected_df = all_related_rows[columns_to_keep]
    selected_df.to_csv(output_file, index=False)
    print(f"malware filtered data saved to {output_file}")
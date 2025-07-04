import pandas as pd

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
    '0fad4bfb7d4697eaf966b88c9e623f454f843700195bb192cfcadcd256a4c431\\w-user-2min-eventid.csv',
    '0fad4bfb7d4697eaf966b88c9e623f454f843700195bb192cfcadcd256a4c431\\wo-user-2min-eventid.csv',
    
    '63e9c56ece51abcf78da3653ed4b03355f36982fdca931043a4bcfca7caf145e\\w-user-2min-eventid.csv',
    '63e9c56ece51abcf78da3653ed4b03355f36982fdca931043a4bcfca7caf145e\\wo-user-2min-eventid.csv',
    
    '85a7d0f873bde902bb66bc481fce0fb3e377113fcf483f8e6d5aca7ca1196141\\w-user-2min-eventid.csv',
    '85a7d0f873bde902bb66bc481fce0fb3e377113fcf483f8e6d5aca7ca1196141\\wo-user-2min-eventid.csv',
    
    '989b2568cd870052182bd637a9a5a1bf68461e2fa59763678f88f634bfd1d51c\\w-user-2min-eventid.csv',
    '989b2568cd870052182bd637a9a5a1bf68461e2fa59763678f88f634bfd1d51c\\wo-user-2min-eventid.csv',

    '2820b924640b10eb028a6fb30a77db1d3c53077e368fdf204fa914306e1ecc3a\\w-user-2min-eventid.csv',
    '2820b924640b10eb028a6fb30a77db1d3c53077e368fdf204fa914306e1ecc3a\\wo-user-2min-eventid.csv',
    
    '83681a0a9a4e5ffcaa56761f4cd6881de408e3408ae579da2a4d006bec0f39c9\\w-user-2min-eventid.csv',
    '83681a0a9a4e5ffcaa56761f4cd6881de408e3408ae579da2a4d006bec0f39c9\\wo-user-2min-eventid.csv',
    
    '491960fbf4a3f4f085b1a318c88d7f39208150b2fd11d1251bf251e912691a0b\\w-user-2min-eventid.csv',
    '491960fbf4a3f4f085b1a318c88d7f39208150b2fd11d1251bf251e912691a0b\\wo-user-2min-eventid.csv',
    
    'ad9093633b9ecaeea7bff69ab8d8781213fec82db6c7f2e963a40d2e0ee0e9ce\\w-user-2min-eventid.csv',
    'ad9093633b9ecaeea7bff69ab8d8781213fec82db6c7f2e963a40d2e0ee0e9ce\\wo-user-2min-eventid.csv',
    
    'def98259bba7c128a22dbb9100a3e9512911d9775ec82175f8a8a3c26b993dbf\\w-user-2min-eventid.csv',
    'def98259bba7c128a22dbb9100a3e9512911d9775ec82175f8a8a3c26b993dbf\\wo-user-2min-eventid.csv',
    
    'e6dff8475541ebddc1f0db47a311eb2c25581b7d5e62af8066d59c283114c2d3\\w-user-2min-eventid.csv',
    'e6dff8475541ebddc1f0db47a311eb2c25581b7d5e62af8066d59c283114c2d3\\wo-user-2min-eventid.csv',
]

for input_file, output_file in zip(input_files, output_files):
    df = pd.read_csv(input_file)

    eventid_counts = df['EventId'].value_counts().reset_index()
    eventid_counts.columns = ['eventid', 'count']
    eventid_counts = eventid_counts.sort_values(by='eventid')
    eventid_counts.to_csv(output_file, index=False)

    print(f"Sorted EventId counts saved to {output_file}")

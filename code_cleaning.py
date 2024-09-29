import pandas as pd


def excel_to_csv(excel_path, output_path, sheet_name=0):
    df = pd.read_excel(excel_path, sheet_name=sheet_name)
    df.to_csv(output_path, index=False)


def Initial_dropping_of_duplicates_and_calculating_their_Frequencies(input_filepath, output_filepath):
    # Reading the file
    df = pd.read_csv(input_filepath, low_memory=False)

    # Compute the frequency for each value in the 'Details' column
    frequency_of_each_value = df['Details'].value_counts().to_dict()

    df['Frequency'] = df['Details'].map(frequency_of_each_value)
    df = df.drop_duplicates(subset='Details', keep='first')
    df.to_csv(output_filepath, index=False)


def further_minimizing_the_data_taking_first_two_words(input_filepath, output_filepath):
    data = pd.read_csv(input_filepath)
    # Removing specific column for a time being
    data = data[~data['Issue'].str.contains('Lightning & Earthing System', case=False, na=False)]

    data['first_two_words'] = data['Details'].str.split().str[:2].str.join(' ')

    # Group by the 'first_two_words' column and sum the frequencies
    summing_frequency = data.groupby('first_two_words')['Frequency'].sum()
    merged_df = data.merge(summing_frequency, on='first_two_words', suffixes=('', '_summing_frequency'))
    merged_df['Frequency'] = merged_df['Frequency_summing_frequency']
    merged_df.drop('Frequency_summing_frequency', axis=1, inplace=True)


def Final_frequency_count(input_filepath, output_filepath):
    data = pd.read_csv(input_filepath)
    details_frequency = data['Details'].value_counts().to_dict()
    data['Frequently'] = data['Details'].map(details_frequency)
    data.drop_duplicates(subset='Details')
    data['Final_frequency'] = data['Frequency'] + data['Frequently']
    data.to_csv(output_filepath)


def extract_and_save_matching_strings(input_filepath, column_name, word, output_csv):
    df = pd.read_excel(input_filepath)
    if column_name not in df.columns:
        raise ValueError(f"Column '{column_name}' not found in the DataFrame")

    # Filter rows where the column contains the specified word
    matching_rows = df[df[column_name].str.contains(word, case=False, na=False)]

    # Extract the specific column and save to CSV
    matching_rows.to_csv(output_csv, index=False)


if __name__ == "__main__":
    input_csv_path = '/home/bilalahmad/Desktop/analysis_task/TowerPreventiveMaintenanceDetails.xlsx'
    output_csv_path = '/home/bilalahmad/Desktop/analysis_task/Detailed_Analysis_Dirt.csv'
    extract_and_save_matching_strings(input_csv_path, 'Details', 'Dirt', output_csv_path)

import pandas as pd

# define file paths
filepath_original_csv = ('cle/Tutoring Transcript Sample copy.csv')
filepath_final_csv = ('cle/Tutoring Transcript Sample Processed.csv')
final_excel_file = ('cle/TutoringTranscriptSampleProcessed.xlsx')

# read from the csv file
df = pd.read_csv(filepath_original_csv)

# filter data leaving only the ID and the course title
new_df = df.loc[(df['TRANSACTION_STS'] == 'H') & (df['CAREER_GPA'] >= 3.0) & (df['GRADE_CDE'].str.contains('A'))]
new_df.rename(columns={'Merged.1' : 'Course_Title'}, inplace=True)
new_df = new_df.drop(columns=['YR_CDE', 'TRANSACTION_STS', 'CAREER_GPA', 'Expr1', 'TRANSACTION_STS', 'TRM_CDE', 'TRANSCRIPT_DIV', 'GRADE_CDE', 'CREDIT_HRS'])

# drop duplicates
new_df = new_df.drop_duplicates()

# save to a new csv file
new_df.to_csv(filepath_final_csv, index=False)

# Create a Pandas Excel writer
writer = pd.ExcelWriter(final_excel_file)

# Loop over unique IDs
for id_value in new_df['ID'].unique():
    # Filter the DataFrame for the current ID
    id_df = new_df[new_df['ID'] == id_value]
    rows = id_df.shape[0]
    if rows >= 5:
        # Write the DataFrame to a sheet with the ID as the sheet name
        id_df.to_excel(writer, sheet_name=str(id_value), index=False)

# Save the Excel file
writer.save()

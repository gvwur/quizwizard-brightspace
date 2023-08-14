import pandas as pd
from datetime import datetime
import os, random, string


def transform_dataframe_mc(df):
    
    df = df.rename(columns=lambda col: 'Option' if col.startswith('Content') or col.startswith('Unnamed') else col)
    transformed_data = []
   
    # Iterate through each row in the original DataFrame
    for _, row in df.iterrows():
        question_text = row['Question']
        answer = str(row['Answer'])

    # Check if the answer contains a comma, indicating multiple answer indices
        if ',' in answer:
            # answer_indices = [int(idx) for idx in answer.split(',')]
            answer_indices = [int(idx) if idx != 'nan' else 0 for idx in answer.split(',')]
        else:
            answer_index = int(answer) if answer != 'nan' else 0

    # Check if the answer contains a comma, if yes change to MS question, else MC question
        if ',' in answer:
            transformed_data.append(['NewQuestion', 'MS', '', ''])
            transformed_data.append(['QuestionText', question_text, '', ''])
        else:
            transformed_data.append(['NewQuestion', 'MC', '', ''])
            transformed_data.append(['QuestionText', question_text, '', ''])

    # Set the correct answers
        for i, option in enumerate(row['Option':'Option'], start=1):
            if ',' in answer:
                is_answer = '100' if i in answer_indices else '0'
            else:
                is_answer = '100' if i == answer_index else '0'
            transformed_data.append(['Option', is_answer, option, ''])

    # Create a new DataFrame from the transformed data
    df = pd.DataFrame(transformed_data, columns=['1','2','3',''])

    # Drop a rows if the third column has no value (NaN).
    df = df.dropna(subset=['3'])

    return df
def transform_dataframe_sa(df):
    
    transformed_data = []

    # Iterate through each row in the original DataFrame
    for _, row in df.iterrows():
        question_text = row['Question']
        answer = str(row['Answer'])
        transformed_data.append(['NewQuestion', 'SA', '', ''])
        transformed_data.append(['QuestionText', question_text, '', ''])
        transformed_data.append(['Answer', '100', answer, ''])
    
    df = pd.DataFrame(transformed_data, columns=['','','',''])
    return df
def transform_dataframe_wr(df):
    
    transformed_data = []

    # Iterate through each row in the original DataFrame
    for _, row in df.iterrows():
        question_text = row['Question']
        transformed_data.append(['NewQuestion', 'WR', '', ''])
        transformed_data.append(['QuestionText', question_text, '', ''])
    
    df = pd.DataFrame(transformed_data, columns=['','','',''])
    return df
def export_dataframe(df, question_type):

    if question_type == 'Multiple Choice Question':
        question_type = 'MC'
    elif question_type == 'Question/Answer':
        question_type = 'SA'
    elif question_type == 'Open Question':
        question_type = 'WR'
    else:
        print("Error with question types")

    now = datetime.now()
    time = now.strftime("%m%d%H%M")
    randomstring = ''.join(random.choices(string.ascii_uppercase + string.digits, k=5))
    filename = f'QuizWizardBS_{question_type}_{time}_{randomstring}.csv'
    df.to_csv(filename, encoding='UTF-8', index=False)
    print (df)
    print("\nTransformation of Excel with: '" + question_type + "' complete.")

def main():
    
    current_directory = os.getcwd()
    all_files = os.listdir(current_directory)
    excel_files = [file for file in all_files if file.endswith('.xlsx')]

    print(excel_files)

    for excel_file in excel_files:
        file_path = os.path.join(current_directory, excel_file)
        df = pd.read_excel(file_path, 'Sheet1', index_col=None)

        question_type_col = 'Type of question'
        mc_question_type = 'Multiple Choice Question'
        sa_question_type = 'Question/Answer'
        wr_question_type = 'Open Question'
    
        if question_type_col in df.columns and mc_question_type in df[question_type_col].values:
            df = transform_dataframe_mc(df)
            export_dataframe(df, mc_question_type)
        elif question_type_col in df.columns and sa_question_type in df[question_type_col].values:
            df = transform_dataframe_sa(df)
            export_dataframe(df, sa_question_type)
        elif question_type_col in df.columns and wr_question_type in df[question_type_col].values:
            df = transform_dataframe_wr(df)
            export_dataframe(df, wr_question_type)
        else:
            print("No Questions found in the DataFrame. Possible Excel formatting incorrect.")

if __name__ == "__main__":
    main()
    input("\nPress Enter to continue...")
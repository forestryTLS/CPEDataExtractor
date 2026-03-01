from pathlib import Path
import pandas as pd

BASE_DIR = Path(__file__).parent

if __name__ == "__main__":
    enrollments_file = BASE_DIR / "all-enrollments.xlsx"
    users_file = BASE_DIR / "all-users.xlsx"

    df_enrollments = pd.read_excel(enrollments_file)
    df_users = pd.read_excel(users_file)

    unique_programs = df_enrollments['Program'].unique()
    unique_sessions = df_enrollments['Session'].unique()

    for program in unique_programs:
        df_program_all_sesisons: pd.DataFrame = df_enrollments[df_enrollments['Program'] == program].copy()

        for session in unique_sessions:
            df_program_single_session: pd.DataFrame = df_program_all_sesisons[df_program_all_sesisons['Session'] == session].copy()

            def fill_null_columns_from_first(s1: pd.Series):
                df2 = df_users[df_users['Student Catalog ID'] == s1.at['Student Catalog ID']]

                if len(df2) > 0:
                    s2: pd.Series = df2.iloc[0]

                    for index, value in s1.items():
                        if value == '' or pd.isna(value) or value is None:
                            new_value = s2.get(index)

                            if new_value != '' and not pd.isna(new_value) and new_value is not None:
                                s1.at[index] = new_value

                return s1

            df_merged = df_program_single_session.apply(fill_null_columns_from_first, axis=1, result_type='reduce')
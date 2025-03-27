import pandas as pd


file_path = "C:\\Users\\Admin\\Documents\\Data Engineering\\data - sample.xlsx"

def read_excel_file(file_path):
    """Read the Excel file and return DataFrames."""
    xls = pd.ExcelFile(file_path)
    attendance_df = pd.read_excel(xls, sheet_name="Attendance_data")
    student_df = pd.read_excel(xls, sheet_name="Student_data")
    return attendance_df, student_df

def find_absence_streaks(attendance_df):
    """Find students absent for more than three consecutive days."""
    attendance_df = attendance_df.sort_values(by=['student_id', 'attendance_date'])
    absence_streaks = []
    
    for student_id, group in attendance_df.groupby('student_id'):
        group = group.reset_index(drop=True)
        streak_start, streak_end, streak_count = None, None, 0
        
        for i, row in group.iterrows():
            if row['status'] == 'Absent':
                if streak_count == 0:
                    streak_start = row['attendance_date']
                streak_end = row['attendance_date']
                streak_count += 1
            else:
                if streak_count > 3:
                    absence_streaks.append([student_id, streak_start, streak_end, streak_count])
                streak_start, streak_end, streak_count = None, None, 0
        
        if streak_count > 3:
            absence_streaks.append([student_id, streak_start, streak_end, streak_count])
    
    return pd.DataFrame(absence_streaks, columns=['student_id', 'absence_start_date', 'absence_end_date', 'total_absent_days'])

def is_valid_email(email):
    """Validate email addresses."""
    return isinstance(email, str) and email.endswith(".com") and "@" in email

def generate_parent_messages(absence_df, student_df):
    """Generate messages for parents based on absence records."""
    merged_df = absence_df.merge(student_df, on='student_id', how='left')
    merged_df['email'] = merged_df['parent_email'].apply(lambda x: x if is_valid_email(str(x)) else None)
    
    merged_df['msg'] = merged_df.apply(lambda row: f"Dear Parent, your child {row['student_name']} was absent from {row['absence_start_date']} to {row['absence_end_date']} for {row['total_absent_days']} days. Please ensure their attendance improves." if row['email'] else None, axis=1)
    
    return merged_df[['student_id', 'absence_start_date', 'absence_end_date', 'total_absent_days', 'email', 'msg']]

def run():
    """Main function to execute the steps."""
    attendance_df, student_df = read_excel_file(file_path)
    absence_df = find_absence_streaks(attendance_df)
    final_df = generate_parent_messages(absence_df, student_df)
    
    
    final_df.to_csv("output.csv", index=False)
    
    return final_df


if __name__ == "__main__":
    result = run()
    print(result.head())

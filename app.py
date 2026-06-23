import streamlit as st
import pandas as pd
import os

FILE_NAME = "schedule.xlsx"

def load_data():
    if not os.path.exists(FILE_NAME):
        st.error(f"Error: Could not find '{FILE_NAME}' in this repository.")
        st.stop()
    # Read the 3 tabs by their sequential order (0, 1, 2)
    current_sched = pd.read_excel(FILE_NAME, sheet_name=0) # Tab 1: Current Schedule
    class_rules = pd.read_excel(FILE_NAME, sheet_name=1)   # Tab 2: Class Rules/Designated Times
    student_info = pd.read_excel(FILE_NAME, sheet_name=2)  # Tab 3: Student Profiles
    return current_sched, class_rules, student_info

# --- STREAMLIT UI SETUP ---
st.set_page_config(page_title="Tutorial Time Changer", layout="wide")
st.title("🕒 Tutorial Class Schedule Portal")
st.write("Parents: Select your child's name below. The system will instantly show available designated times for their specific subject and grade.")

current_sched, class_rules, student_info = load_data()

# Clean up column headers by removing hidden spaces
current_sched.columns = current_sched.columns.str.strip()
class_rules.columns = class_rules.columns.str.strip()
student_info.columns = student_info.columns.str.strip()

# Show the live spreadsheet view at the very top
st.subheader("📊 Live Schedule Spreadsheet")
st.dataframe(current_sched, use_container_width=True, hide_index=True)

st.write("---")
st.subheader("🔄 Change Your Child's Time Slot")

# 1. Parent selects Student Name
student_list = student_info["Student Name"].dropna().unique().tolist()
selected_student = st.selectbox("1. Select Your Child's Name:", ["-- Choose Name --"] + student_list)

if selected_student != "-- Choose Name --":
    # Locate the student's current row
    student_row = current_sched[current_sched["Student Name"] == selected_student]
    info_row = student_info[student_info["Student Name"] == selected_student]
    
    if student_row.empty or info_row.empty:
        st.error("Student name profile could not be found completely across the tabs. Please notify reception.")
    else:
        # Dynamically find column names that contain key terms to avoid spelling errors
        time_col = [col for col in current_sched.columns if "time" in col.lower() or "slot" in col.lower()][0]
        subject_col = [col for col in current_sched.columns if "subject" in col.lower()][0]
        grade_col = [col for col in student_info.columns if "grade" in col.lower()][0]
        
        # Extract individual text entries cleanly
        student_subject = str(student_row[subject_col].values[0]).strip()
        current_time = str(student_row[time_col].values[0]).strip()
        student_grade = str(info_row[grade_col].values[0]).strip()
        
        st.info(f"📋 **Detected Profile:** {selected_student} is in **Grade {student_grade}** taking **{student_subject}** (Current Slot: {current_time})")
        
        # Dynamically match column names in Tab 2 Rules
        rule_time_col = [col for col in class_rules.columns if "time" in col.lower() or "slot" in col.lower()][0]
        rule_subject_col = [col for col in class_rules.columns if "subject" in col.lower()][0]
        rule_grade_col = [col for col in class_rules.columns if "grade" in col.lower()][0]
        rule_seats_col = [col for col in class_rules.columns if "seat" in col.lower()][0]
        
        # 2. Filter Tab 2 Rules: Show options matching this Subject AND Grade Level
        matched_classes = class_rules[
            (class_rules[rule_subject_col].astype(str).str.lower() == student_subject.lower()) & 
            (class_rules[rule_grade_col].astype(str).str.strip() == student_grade)
        ]
        
        available_times = matched_classes[rule_time_col].dropna().unique().tolist()
        
        if not available_times:
            st.warning(f"No alternative designated slots found in master rules for Grade {student_grade} {student_subject}.")
        else:
            with st.form("submit_swap_form"):
                requested_time = st.selectbox("2. Select From Available Designated Times for This Subject:", available_times)
                submit_button = st.form_submit_button("Confirm New Time Slot")
                
            if submit_button:
                if requested_time == current_time:
                    st.warning("Your child is already assigned to this time slot!")
                else:
                    # Check seat limits for the targeted class choice
                    target_class_rule = matched_classes[matched_classes[rule_time_col] == requested_time]
                    max_seats = int(target_class_rule[rule_seats_col].values[0])
                    
                    # Count seats currently taken in Tab 1
                    seats_taken = len(current_sched[current_sched[time_col] == requested_time])
                    
                    if seats_taken >= max_seats:
                        st.error(f"❌ Sorry, the {requested_time} class is completely FULL ({seats_taken}/{max_seats} seats taken).")
                    else:
                        # Update Tab 1 in memory
                        current_sched.loc[current_sched["Student Name"] == selected_student, time_col] = requested_time
                        
                        # Save the updated data directly back into the original 3 Excel tabs
                        with pd.ExcelWriter(FILE_NAME, engine="openpyxl") as writer:
                            current_sched.to_excel(writer, sheet_name="Sheet1", index=False)
                            class_rules.to_excel(writer, sheet_name="Sheet2", index=False)
                            student_info.to_excel(writer, sheet_name="Sheet3", index=False)
                            
                        st.success(f"🎉 Success! {selected_student} has been updated to {requested_time} in real-time.")
                        st.rerun()

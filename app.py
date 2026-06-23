import streamlit as st
import pandas as pd
import os

FILE_NAME = "schedule.xlsx"

def load_data():
    if not os.path.exists(FILE_NAME):
        st.error(f"Error: Could not find '{FILE_NAME}' in this folder.")
        st.stop()
    # Read the 3 tabs cleanly
    current_sched = pd.read_excel(FILE_NAME, sheet_name=0) # Tab 1
    class_rules = pd.read_excel(FILE_NAME, sheet_name=1)   # Tab 2
    student_info = pd.read_excel(FILE_NAME, sheet_name=2)  # Tab 3
    return current_sched, class_rules, student_info

# --- STREAMLIT UI ---
st.set_page_config(page_title="Tutorial Time Changer", layout="wide")
st.title("🕒 Tutorial Class Schedule Portal")
st.write("Parents: Select your child's name. The system will automatically show available times for their subject.")

current_sched, class_rules, student_info = load_data()

# Clean up column names by removing hidden spaces
current_sched.columns = current_sched.columns.str.strip()
class_rules.columns = class_rules.columns.str.strip()
student_info.columns = student_info.columns.str.strip()

# Show Live Spreadsheet
st.subheader("📊 Live Schedule Spreadsheet")
st.dataframe(current_sched, use_container_width=True, hide_index=True)

st.write("---")
st.subheader("🔄 Change Your Child's Time Slot")

# 1. Parent selects Student Name
student_list = student_info["Student Name"].dropna().unique().tolist()
selected_student = st.selectbox("1. Select Your Child's Name:", ["-- Choose Name --"] + student_list)

if selected_student != "-- Choose Name --":
    # Automatically look up the student's assigned subject from Tab 1
    student_row = current_sched[current_sched["Student Name"] == selected_student]
    
    if student_row.empty:
        st.error("Student name not found in the current schedule sheet.")
    else:
        # Detect current subject and current time slot
        # Adjust these column headers if your Excel uses different names (e.g., 'Class Time' or 'Time Slot')
        time_col = [col for col in current_sched.columns if "time" in col.lower() or "slot" in col.lower()][0]
        subject_col = [col for col in current_sched.columns if "subject" in col.lower()][0]
        
        student_subject = student_row[subject_col].values[0]
        current_time = student_row[time_col].values[0]
        
        # Get student's grade level from Tab 3
        student_grade = student_info[student_info["Student Name"] == selected_student]["Grade Level"].values[0]
        
        st.info(f"📋 **Detected Profile:** {selected_student} is in **Grade {student_grade}** studying **{student_subject}** (Current Slot: {current_time})")
        
        # 2. Filter Tab 2 Rules: Only show slots matching this Subject AND Grade Level
        rule_time_col = [col for col in class_rules.columns if "time" in col.lower() or "slot" in col.lower()][0]
        rule_subject_col = [col for col in class_rules.columns if "subject" in col.lower()][0]
        rule_grade_col = [col for col in class_rules.columns if "grade" in col.lower()][0]
        rule_seats_col = [col for col in class_rules.columns if "seat" in col.lower()][0]
        
        matched_classes = class_rules[
            (class_rules[rule_subject_col].str.lower() == student_subject.lower()) & 
            (class_rules[rule_grade_col].astype(str) == str(student_grade))
        ]
        
        available_times = matched_classes[rule_time_col].unique().tolist()
        
        if not available_times:
            st.warning(f"No alternative time slots found in master rules for Grade {student_grade} {student_subject}.")
        else:
            with st.form("submit_swap_form"):
                requested_time = st.selectbox("2. Select From Available Designated Times for This Subject:", available_times)
                submit_button = st.form_submit_button("Confirm New Time Slot")
                
            if submit_button:
                if requested_time == current_time:
                    st.warning("Your child is already assigned to this time slot!")
                else:
                    # Check seat limit for the newly selected class
                    target_class_rule = matched_classes[matched_classes[rule_time_col] == requested_time]
                    max_seats = target_class_rule[rule_seats_col].values[0]
                    
                    # Count how many kids are currently in that target slot
                    seats_taken = len(current_sched[current_sched[time_col] == requested_time])
                    
                    if seats_taken >= max_seats:
                        st.error(f"❌ Sorry, the {requested_time} class is completely FULL ({seats_taken}/{max_seats} seats taken).")
                    else:
                        # Success! Modify the time cell safely
                        current_sched.loc[current_sched["Student Name"] == selected_student, time_col] = requested_time
                        
                        # Save the updated data directly back to Excel across all 3 tabs
                        with pd.ExcelWriter(FILE_NAME, engine="openpyxl") as writer:
                            current_sched.to_excel(writer, sheet_name=writer.handles if hasattr(writer, 'handles') else "Sheet1", index=False)
                            class_rules.to_excel(writer, sheet_name="Sheet2", index=False)
                            student_info.to_excel(writer, sheet_name="Sheet3", index=False)
                            
                        st.success(f"🎉 Success! {selected_student} has been moved to {requested_time}.")
                        st.rerun()

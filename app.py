import streamlit as st
import pandas as pd
import os

FILE_NAME = "schedule.xlsx"

# 1. Load the 3 tabs exactly as they are structured
def load_data():
    if not os.path.exists(FILE_NAME):
        st.error(f"Error: Could not find '{FILE_NAME}' in this folder. Please place your Excel file here.")
        st.stop()
    current_sched = pd.read_excel(FILE_NAME, sheet_name=0) # Tab 1
    class_rules = pd.read_excel(FILE_NAME, sheet_name=1)   # Tab 2
    student_info = pd.read_excel(FILE_NAME, sheet_name=2)  # Tab 3
    return current_sched, class_rules, student_info

# --- STREAMLIT WEB INTERFACE ---
st.set_page_config(page_title="Parent Schedule Changer", layout="wide")
st.title("🕒 Tutorial Class Schedule Changer")
st.write("Parents: Use this portal to change your child's time. The system validates seats and grades instantly.")

# Load live sheets
current_sched, class_rules, student_info = load_data()

# Show the live current schedule to parents
st.subheader("📊 Live Spreadsheet View")
st.dataframe(current_sched, use_container_width=True, hide_index=True)

st.write("---")
st.subheader("🔄 Change Your Child's Class Time")

# Create simple dropdown inputs for the parents
with st.form("change_time_form"):
    col1, col2 = st.columns(2)
    
    with col1:
        student_list = student_info["Student Name"].tolist()
        selected_student = st.selectbox("1. Select Your Child's Name:", student_list)
        
    with col2:
        # Get list of unique available class times from Tab 2
        available_times = class_rules["Date & Time"].unique().tolist()
        requested_time = st.selectbox("2. Select Your Desired New Time Slot:", available_times)
        
    submit_button = st.form_submit_button("Confirm and Update Spreadsheet")

# When a parent clicks the button, Python instantly calculates and edits the file
if submit_button:
    # Get student's current row data
    student_row = current_sched[current_sched["Student Name"] == selected_student]
    
    if student_row.empty:
        st.error("Student not found in the current schedule.")
    else:
        current_time = student_row["Date & Time"].values[0]
        
        if current_time == requested_time:
            st.warning("Your child is already assigned to this time slot!")
        else:
            # 🛑 CRITICAL RULE CHECKS (Foolproof verification)
            # Check Grade Level Match
            student_grade = student_info[student_info["Student Name"] == selected_student]["Grade Level"].values[0]
            target_class = class_rules[class_rules["Date & Time"] == requested_time]
            
            if target_class.empty:
                st.error("This target class does not exist in the master rules.")
            else:
                target_grade = target_class["Grade Level Match"].values[0]
                max_seats = target_class["Max Seats"].values[0]
                
                # Check Subject Match
                student_subject = student_row["Subject"].values[0]
                target_subject = target_class["Subject"].values[0]
                
                # Rule 1: Subject must match
                if student_subject != target_subject:
                    st.error(f"❌ Cannot change! {selected_student} studies {student_subject}, but the requested slot is for {target_subject}.")
                
                # Rule 2: Grade level must match
                elif student_grade != target_grade:
                    st.error(f"❌ Cannot change! This slot is for Grade {target_grade}, but your child is Grade {student_grade}.")
                
                else:
                    # Rule 3: Check seat capacity
                    seats_taken = len(current_sched[current_sched["Date & Time"] == requested_time])
                    
                    if seats_taken >= max_seats:
                        st.error(f"❌ Sorry, the {requested_time} class is completely FULL ({seats_taken}/{max_seats} seats taken).")
                    else:
                        # Success! Update the DataFrame index safely
                        current_sched.loc[current_sched["Student Name"] == selected_student, "Date & Time"] = requested_time
                        
                        # Save back into the exact original Excel structure across the 3 tabs
                        with pd.ExcelWriter(FILE_NAME, engine="openpyxl") as writer:
                            current_sched.to_excel(writer, sheet_name="Sheet1", index=False)
                            class_rules.to_excel(writer, sheet_name="Sheet2", index=False)
                            student_info.to_excel(writer, sheet_name="Sheet3", index=False)
                        
                        st.success(f"🎉 Success! {selected_student} has been moved to {requested_time}. The spreadsheet is updated.")
                        st.rerun()

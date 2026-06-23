import streamlit as st
import pandas as pd
import os

FILE_NAME = "schedule.xlsx"

def load_data():
    if not os.path.exists(FILE_NAME):
        st.error(f"Error: Could not find '{FILE_NAME}' in this repository.")
        st.stop()
    # Load tabs exactly in order (Tab 1, Tab 2, Tab 3)
    student_info = pd.read_excel(FILE_NAME, sheet_name=0)  # Tab 1: Students
    class_rules = pd.read_excel(FILE_NAME, sheet_name=1)   # Tab 2: Classes
    current_sched = pd.read_excel(FILE_NAME, sheet_name=2) # Tab 3: Schedule
    return student_info, class_rules, current_sched

# --- STREAMLIT UI SETUP ---
st.set_page_config(page_title="Tutorial Time Changer", layout="wide")
st.title("🕒 Tutorial Class Schedule Portal")
st.write("Parents: Select your child's name below to change their class time slot for their designated subject.")

student_info, class_rules, current_sched = load_data()

# Clean up column headers by removing hidden leading/trailing spaces
student_info.columns = student_info.columns.str.strip()
class_rules.columns = class_rules.columns.str.strip()
current_sched.columns = current_sched.columns.str.strip()

# Show the live schedule view at the very top (Tab 3 data)
st.subheader("📊 Live Schedule Spreadsheet")
st.dataframe(current_sched, use_container_width=True, hide_index=True)

st.write("---")
st.subheader("🔄 Change Your Child's Time Slot")

# 1. Parent selects Student Name from Tab 1
student_list = student_info["Student Name"].dropna().unique().tolist()
selected_student = st.selectbox("1. Select Your Child's Name:", ["-- Choose Name --"] + student_list)

if selected_student != "-- Choose Name --":
    # Locate student entries across sheets safely
    info_row = student_info[student_info["Student Name"] == selected_student]
    sched_row = current_sched[current_sched["Student Name"] == selected_student]
    
    if sched_row.empty or info_row.empty:
        st.error(f"Could not find a current class schedule entry for '{selected_student}'. Please contact reception.")
    else:
        # Pull exact data matching your tab structure cleanly
        student_grade = str(info_row["Grade Level"].iloc[0]).strip()
        student_subject = str(sched_row["Subject"].iloc[0]).strip()
        current_time = str(sched_row["Date & Time"].iloc[0]).strip()
        current_room = str(sched_row["Room Name"].iloc[0]).strip()
        
        st.info(f"📋 **Profile Found:** {selected_student} | **Grade {student_grade}** | **{student_subject}** (Current Room: {current_room} at {current_time})")
        
        # 2. Filter Tab 2 (Classes) to find alternative slots for this exact Subject and Grade Match
        # Find the column that matches 'Grade Level Match' even if cut off in naming
        grade_match_col = [col for col in class_rules.columns if "grade" in col.lower()][0]
        
        matched_classes = class_rules[
            (class_rules["Subject"].astype(str).str.strip().str.lower() == student_subject.lower()) & 
            (class_rules[grade_match_col].astype(str).str.strip() == student_grade)
        ]
        
        available_times = matched_classes["Date & Time"].dropna().unique().tolist()
        
        if not available_times:
            st.warning(f"No alternative time slots found in master rules for Grade {student_grade} {student_subject}.")
        else:
            with st.form("submit_swap_form"):
                requested_time = st.selectbox("2. Select From Available Designated Times for This Subject:", available_times)
                submit_button = st.form_submit_button("Confirm New Time Slot")
                
            if submit_button:
                if str(requested_time).strip() == current_time:
                    st.warning("Your child is already assigned to this time slot!")
                else:
                    # Get the rules for the newly chosen class time slot
                    target_class_rule = matched_classes[matched_classes["Date & Time"] == requested_time]
                    max_seats = int(target_class_rule["Max Seats"].iloc[0])
                    new_room_name = str(target_class_rule["Room Name"].iloc[0]).strip()
                    
                    # Compute total seats currently taken in the Live Schedule (Tab 3) for that time slot
                    seats_taken = len(current_sched[current_sched["Date & Time"] == requested_time])
                    
                    if seats_taken >= max_seats:
                        st.error(f"❌ Sorry, the class on {requested_time} is completely FULL ({seats_taken}/{max_seats} seats taken).")
                    else:
                        # Success! Update both Date & Time and Room Name columns in Tab 3 array row
                        current_sched.loc[current_sched["Student Name"] == selected_student, "Date & Time"] = requested_time
                        current_sched.loc[current_sched["Student Name"] == selected_student, "Room Name"] = new_room_name
                        
                        # Save the dataframes back to Excel sheets matching your exact tab index order
                        with pd.ExcelWriter(FILE_NAME, engine="openpyxl") as writer:
                            student_info.to_excel(writer, sheet_name="Sheet1", index=False)
                            class_rules.to_excel(writer, sheet_name="Sheet2", index=False)
                            current_sched.to_excel(writer, sheet_name="Sheet3", index=False)
                            
                        st.success(f"🎉 Success! {selected_student} has been updated to {requested_time} in {new_room_name}.")
                        st.rerun()

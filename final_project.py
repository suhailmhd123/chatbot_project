#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import uuid  # Import uuid to generate unique IDs


# Function to save responses to Excel
def save_to_excel(data, file_name='bot_excel.xlsx'):
    try:
        # Try to load the existing Excel file and sheet using 'openpyxl' engine
        existing_df = pd.read_excel(file_name, sheet_name='Sheet1', engine='openpyxl')
        # Append the new data
        df = pd.concat([existing_df, pd.DataFrame([data])], ignore_index=True)
    except FileNotFoundError:
        # If file doesn't exist, create a new DataFrame
        df = pd.DataFrame([data])

    # Save the DataFrame to Excel (create a new file or overwrite if it already exists)
    with pd.ExcelWriter(file_name, engine='openpyxl', mode='w') as writer:
        df.to_excel(writer, sheet_name='Sheet1', index=False)


def health_chatbot():
    # Start of conversation
    st.title("Welcome to Happy Couple Solution - Online Health Guide")
    
    # Initialize dictionary to store responses
    responses = {}
    
    # Generate a unique ID for this session
    unique_id = str(uuid.uuid4())  # Create a unique ID
    responses["Unique ID"] = unique_id  # Add unique ID to responses

    user_input = st.text_input("User: Say 'hi' to start the conversation.").lower()

    if user_input == "hi":
        st.write("Chatbot: Hi! I am your Online Health guide From Happy Couple Solution. Welcome!")
        st.write("Chatbot: Kindly complete a confidential interaction to find your ideal health product.")

        customer_name = st.text_input("Please enter your name:")
        responses["Customer Name"] = customer_name

        # Ask for health concerns
        st.write("Chatbot: What are your prior concerns about health?")
        concern = st.radio("Select a concern:", ["Sexual Wellness", "Stress Management"])
        responses["Health Concern"] = concern

        if concern == "Sexual Wellness":
            st.write("Chatbot: Share your main concern you need a complete solution for?")
            main_concern = st.radio("Select your main concern:", ["Erectile Dysfunction (ED)", "Premature Ejaculation (PE)", "Low sexual interest / Poor Satisfaction"])
            responses["main_concern"] = main_concern
            
            if main_concern == "Erectile Dysfunction (ED)":
                st.write("Chatbot: Which of the following issues do you identify with?")
                erection_issue = st.radio("Select an issue:", [
                    "Occasionally, my erection is not hard enough to penetrate",
                    "I usually find it difficult to maintain an erection",
                    "I haven't had an erection for some days/months",
                    "I am very interested in sex, but my erection level is zero",
                    "I have no issues in maintaining my erection"
                ])
                
                responses["erection_issue"] = erection_issue

                pe_issue = st.radio("Chatbot: Do you have issues with Premature Ejaculation?", ["Yes", "No"])
                responses["pe_issue"] = pe_issue

                # Basic physical fitness questions
                age = st.text_input("Chatbot: What is your age?")
                sex = st.selectbox("Chatbot: What is your sex?", ["Male", "Female", "Others"])
                height = st.text_input("Chatbot: What is your height?")
                weight = st.text_input("Chatbot: What is your weight?")
                
                # Store basic information
                responses["Age"] = age
                responses["Sex"] = sex
                responses["Height"] = height
                responses["Weight"] = weight
                
                st.write("Chatbot: Do you have any of the following medical conditions?")
                medical_condition = st.radio("Select a medical condition:", [
                    "Diabetes", "High BP", "Low BP", "High Cholesterol", 
                    "Thyroid issues", "Creatinine issues", "Cardiac issues", "None of the above"
                ])
                
                responses["medical_condition"] = medical_condition

                sexual_frequency = st.radio("Chatbot: What is your sexual activity frequency?", [
                    "Daily", "Once a week", "More than once a week", "Less than once a week", 
                    "Don't keep track", "It's been a while since I have been active"
                ])
                responses["sexual_frequency"] = sexual_frequency

                Physical_Activity = st.radio("Physical Activity Levels, How would you describe your activity levels?", [
                    "Sedentary", "Moderately active", "Very active"
                ])
                responses["Physical_Activity"] = Physical_Activity

                Diet_eat = st.radio("Diet and Eating habits, Which of the following do you identify with?", [
                    "I eat junk food regularly", "I usually skip meals", "I usually eat home-cooked meals",
                    "Balanced food", "None of the above"
                ])
                responses["Diet_eat"] = Diet_eat

                Digestive_issues = st.radio("Digestive issues, Do you experience constipation and gas regularly?", ["Yes", "No"])
                responses["Digestive_issues"] = Digestive_issues
                
                Stress_levels = st.radio("Do you experience stress regularly?", [
                    "I am under a lot of stress", "I’ve been more stressed lately", "The normal stress of regular life", "No stress"
                ])
                responses["Stress_levels"] = Stress_levels

                Sleep_patterns = st.radio("For how many hours do you usually sleep at night?", [
                    "Less than 6 hours", "6 to 8 hours", "8 hours+"
                ])
                responses["Sleep_patterns"] = Sleep_patterns

                Life_habits = st.radio("Life habits, Which of the following apply to you?", [
                    "I smoke regularly", "I drink alcohol regularly", "I drink more than 3 caffeinated drinks per day", "None of the above"
                ])
                responses["Life_habits"] = Life_habits

                Medication_concerns = st.radio("Are you concerned about the side effects of erectile dysfunction medications, and would you prefer an alternative treatment that doesn't involve medication?", ["Yes", "No"])
                responses["Medication_concerns"] = Medication_concerns
                
                if erection_issue == "I usually find it difficult to maintain an erection":
                    st.write("Based on your response, we recommend considering the Erectaid Vacuum Therapy Device.")
                else:
                    st.write("Based on your responses, further suggestions will be made by our experts.")

                any_treatment = st.radio("Are you taking any treatment?", ["Yes", "No"])
                responses["any_treatment"] = any_treatment

                if any_treatment == "Yes":
                    uploaded_file = st.file_uploader("Upload doctor prescription", type=["pdf", "doc", "docx"])
                    responses["doctor_prescription"] = uploaded_file
                    st.write("We will evaluate your prescription and give further recommendations.")
                else:
                    st.write("Our team will provide conclusions and recommendations based on your responses.")

                # Save the data to Excel
                if st.button("Submit ED", key="submit_ed"):
                    save_to_excel(responses)
                    st.success("Responses saved successfully!")

            elif main_concern in ["Premature Ejaculation (PE)", "Low sexual interest / Poor Satisfaction"]:
                # Same flow as ED but adapted for these concerns
                st.write("Chatbot: Let's gather some basic information about your physical fitness.")
                age = st.text_input("Chatbot: What is your age?")
                sex = st.selectbox("Chatbot: What is your sex?", ["Male", "Female", "Others"])
                height = st.text_input("Chatbot: What is your height?")
                weight = st.text_input("Chatbot: What is your weight?")
                
                # Store basic information
                responses["Age"] = age
                responses["Sex"] = sex
                responses["Height"] = height
                responses["Weight"] = weight

                # More specific questions for PE or Low interest / Poor satisfaction
                medical_condition = st.radio("Chatbot: Do you have any of the following medical conditions?", [
                    "Diabetes", "High BP", "Low BP", "High Cholesterol", 
                    "Thyroid issues", "Creatinine issues", "Cardiac issues", "None of the above"
                ])
                responses["medical_condition"] = medical_condition
                penetation_issue = st.radio("Chatbot: Which of the following issues do you identify with?", [
                    "I ejaculate before penetration", "I find myself ejaculating too early during sex", "I have no issues with ejaculation"
                ])
                responses["penetation_issue"] = penetation_issue
                
                Medication_concerns = st.radio("Have you issue with erectile dysfunction?", [
                    "Yes", "No", 
                ])
                responses["Medication_concerns"] = Medication_concerns
                
                sexual_frequency = st.radio("Chatbot: What is your sexual activity frequency?", [
                    "Daily", "Once a week", "More than once a week", "Less than once a week", 
                    "Don't keep track", "It's been a while since I have been active"
                ])
                responses["sexual_frequency"] = sexual_frequency
                
                Physical_Activity = st.radio("Physical Activity Levels, How would you describe your activity levels?", [
                    "Sedentary", "Moderately active", "Very active"
                ])
                responses["Physical_Activity"] = Physical_Activity
                Diet_eat = st.radio("Diet and Eating habits, Which of the following do you identify with?", [
                    "I eat junk food regularly", "I usually skip meals", "I usually eat home-cooked meals",
                    "Balanced food", "None of the above"
                ])
                responses["Diet_eat"] = Diet_eat
                
                Stress_levels = st.radio("Do you experience stress regularly?", [
                    "I am under a lot of stress", "I’ve been more stressed lately", "The normal stress of regular life", "No stress"
                ])
                responses["Stress_levels"] = Stress_levels
                Sleep_patterns = st.radio("For how many hours do you usually sleep at night?", [
                    "Less than 6 hours", "6 to 8 hours", "8 hours+"
                ])
                responses["Sleep_patterns"] = Sleep_patterns
                Life_habits = st.radio("Life habits, Which of the following apply to you?", [
                    "I smoke regularly", "I drink alcohol regularly", "I drink more than 3 caffeinated drinks per day", "None of the above"
                ])
                responses["Life_habits"] = Life_habits
                taking_medicine = st.radio("Are you taking any medicine continuously?", ["Yes", "No"])
                responses["taking_medicine"] = taking_medicine
                
                if taking_medicine == "Yes":
                    uploaded_file = st.file_uploader("Upload your doctor prescription", type=["pdf", "doc", "docx"])
                    responses["doctor_prescription"] = uploaded_file
                    st.write("We will review your prescription and provide recommendations.")

                else:
                    st.write("Our team will evaluate your responses and give suggestions.")

                if st.button("Submit PE", key="submit_pe"):
                    save_to_excel(responses)
                    st.success("Responses saved successfully!")


        elif concern == "Stress Management":
            st.write("Chatbot: Let's gather some basic information about your physical fitness.")
            age = st.text_input("Chatbot: What is your age?")
            sex = st.selectbox("Chatbot: What is your sex?", ["Male", "Female", "Others"])
            height = st.text_input("Chatbot: What is your height?")
            weight = st.text_input("Chatbot: What is your weight?")
            
            # Store basic information
            responses["Age"] = age
            responses["Sex"] = sex
            responses["Height"] = height
            responses["Weight"] = weight           
            
            st.write("Chatbot: Do you have any of the following medical conditions?")
            medical_condition = st.radio("Select a medical condition:", [
                "Diabetes", "High BP", "Low BP", "High Cholesterol", 
                "Thyroid issues", "Creatinine issues", "Cardiac issues", "None of the above"
            ])
            responses["medical_condition"] = medical_condition

            stress_frequency = st.radio("How often do you feel stressed or anxious?", ["Rarely", "Sometimes", "Often", "Almost always"])
            responses["stress_frequency"] = stress_frequency
            sleeping_trouble = st.radio("Do you have trouble sleeping due to stress or anxiety?", ["Yes", "No"])
            responses["sleeping_trouble"] = sleeping_trouble
            life_changes = st.radio("Have you experienced any major life changes recently?", ["Yes", "No"])
            responses["life_changes"] = life_changes
            stress_coping = st.radio("How do you usually cope with stress?", [
                "Exercise", "Meditation", "Talking to friends/family", "Hobbies", "None of the above"
            ])
            responses["stress_coping"] = stress_coping

            Activity_level = st.radio("How would you describe your activity levels?", ["Sedentary", "Moderately active", "Very active"])
            responses["Activity_level"] = Activity_level
            sleep_duration = st.radio("For how many hours do you usually sleep at night?", ["Less than 6 hours", "6 to 8 hours", "8 hours+"])
            responses["sleep_duration"] = sleep_duration
            Diet_eat = st.radio("Diet and Eating habits, Which of the following do you identify with?", [
                "I eat junk food regularly", "I usually skip meals", "I usually eat home-cooked meals", "Balanced food", "None of the above"
            ])
            responses["Diet_eat"] = Diet_eat

            Digestive_issues = st.radio("Do you experience constipation and gas regularly?", ["Yes", "No"])
            responses["Digestive_issues"] = Digestive_issues
            Life_habits = st.radio("Which of the following apply to you?", [
                "I smoke regularly", "I drink alcohol regularly", "I drink more than 3 caffeinated drinks per day", "None of the above"
            ])
            responses["Life_habits"] = Life_habits

            taking_medicine = st.radio("Are you taking any medicine continuously?", ["Yes", "No"])
            responses["taking_medicine"] = taking_medicine

            if taking_medicine == "Yes":
                uploaded_file = st.file_uploader("Upload your doctor prescription", type=["pdf", "doc", "docx"])
                responses["doctor_prescription"] = uploaded_file
                st.write("We will review your prescription and provide recommendations.")
                
            else:
                st.write("Our team will evaluate your responses and give suggestions.")

            st.write("Chatbot: Thank you for sharing your information! We will suggest appropriate stress management solutions for you.")
    
            if st.button("Submit Stress", key="submit_stress"):
            # Save the data to Excel
                save_to_excel(responses)
                st.success("Responses saved successfully!")
        else:
            st.error("Please fill all Details.")                   

    else:
        st.write("Chatbot: Please say 'hi' to start the conversation.")

if __name__ == "__main__":
    health_chatbot()


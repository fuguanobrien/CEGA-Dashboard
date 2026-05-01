#!/usr/bin/env python3
"""
Combined Survey Data Processor (2020-2024) — Syrian Refugee Life Study (S-RLS)
Cleans, standardizes, combines all survey years → Excel workbook for Tableau.
"""
import pandas as pd
import numpy as np
import re
import warnings
warnings.filterwarnings('ignore')

DATA_DIR = "/Users/fuguan/Desktop/survey data"

# Sentinel values used as "don't know" / "refused" — exclude from calculations
SENTINEL_VALUES = {-99, -98, -96, -88, -77, -66}

def mask_sentinels(series):
    """Replace sentinel values with NaN so they're excluded from calculations."""
    num = pd.to_numeric(series, errors='coerce')
    return num.where(~num.isin(SENTINEL_VALUES))


###############################################################################
# COLUMN MAPPINGS: (original_col → (std_name, topic))
###############################################################################

def m2020():
    m = {}
    # Demographics
    m['ID'] = ('Respondent_ID', 'Demographics')
    m['Year of birth'] = ('Year_of_Birth', 'Demographics')
    m['Age (in years)'] = ('Age', 'Demographics')
    m['Gender'] = ('Gender', 'Demographics')
    m['Marital status'] = ('Marital_Status', 'Demographics')
    m['Registered with UNHCR (Y/N)'] = ('Registered_UNHCR', 'Demographics')
    m['Reasons not registered in the UNHCR?'] = ('Reasons_Not_Registered_UNHCR', 'Demographics')
    m['Has a MOI card?'] = ('Has_MOI_Card', 'Demographics')
    m['Has a passport?'] = ('Has_Passport', 'Demographics')
    m['Has a residency permit?'] = ('Has_Residency_Permit', 'Demographics')
    m['Has a UNHCR file?'] = ('Has_UNHCR_File', 'Demographics')
    m['Has a work permit?'] = ('Has_Work_Permit', 'Demographics')
    m['Has a household family Book?'] = ('Has_Family_Book', 'Demographics')
    m['Has a Syrian ID?'] = ('Has_Syrian_ID', 'Demographics')
    m['Country'] = ('Country', 'Demographics')
    m['Governorate'] = ('Governorate', 'Demographics')
    m['District'] = ('District', 'Demographics')
    m['Sub-district'] = ('Sub_District', 'Demographics')
    m['Lives in a formal refugee camp?'] = ('Lives_In_Camp', 'Demographics')
    m['Location of interview'] = ('Interview_Location', 'Demographics')

    # Household
    m['Number of people in the household (other than FR)'] = ('HH_Size', 'Household')
    m['Head of household (Y/N)'] = ('Head_of_HH', 'Household')
    m['Currently breastfeeding (Y/N)'] = ('Currently_Breastfeeding', 'Household')
    m['Household chores hours (indiv, week)'] = ('HH_Chores_Hours_Week', 'Household')
    m['Childcare hours (indiv, week)'] = ('Childcare_Hours_Week', 'Household')
    m['Other people doing HH chores/childcare'] = ('Others_Doing_Chores', 'Household')
    m['Total hours HH chors/childcare (HH, week)'] = ('Total_Chores_Childcare_Hours_Week', 'Household')

    # Education
    m['Attending school (Y/N)'] = ('Currently_Attending_School', 'Education')
    m['Type of school last attended (1/3)'] = ('Type_School_Last_Attended', 'Education')
    m['Year last attended school'] = ('Year_Last_Attended_School', 'Education')
    m['Type of school last attended'] = ('School_Public_Private', 'Education')
    m['Country of school last attended'] = ('Country_School_Last_Attended', 'Education')
    m['Length of schooling, years'] = ('Length_Schooling_Years', 'Education')
    m['Length of schooling, months'] = ('Length_Schooling_Months', 'Education')
    m['Completed any other type of education (Y/N)'] = ('Completed_Other_Education', 'Education')

    # Dwelling
    m['Number of rooms in dwelling'] = ('Number_Rooms_Dwelling', 'Dwelling')
    m['Occupies entire dwelling (Y/N)'] = ('Occupies_Entire_Dwelling', 'Dwelling')
    m['Rooms that HH occupies'] = ('Rooms_HH_Occupies', 'Dwelling')
    m['Another family living in dwelling (Y/N)'] = ('Another_Family_In_Dwelling', 'Dwelling')
    m['Number of families in dwelling'] = ('Number_Families_In_Dwelling', 'Dwelling')
    m['Made of temporary materials (Y/N)'] = ('Temporary_Materials', 'Dwelling')
    m['Prefabricated housing unit /caravan (Y/N)'] = ('Prefab_Housing', 'Dwelling')
    m['Floors materials'] = ('Floor_Material', 'Dwelling')
    m['Roof material'] = ('Roof_Material', 'Dwelling')
    m['Acces to electricity (Y/N)'] = ('Has_Electricity', 'Dwelling')
    m['Source of your electricity'] = ('Electricity_Source', 'Dwelling')
    m['Most used toilet facility'] = ('Toilet_Facility', 'Dwelling')
    m['Main water source (week)'] = ('Main_Water_Source', 'Dwelling')
    m['Source of drinking water (last week)'] = ('Drinking_Water_Source', 'Dwelling')
    m['Treat water in last 7 days?'] = ('Treats_Water', 'Dwelling')
    m['Access safe water for drinking (last week)'] = ('Safe_Water_Access', 'Dwelling')
    m['Water easy to access'] = ('Water_Easy_Access', 'Dwelling')

    # Economics
    m['Own or rent living place'] = ('Own_Or_Rent', 'Economics')
    m['Monthly rent, dinar (agreed)'] = ('Monthly_Rent_Agreed_JOD', 'Economics')
    m['Monthly rent, dinar (paid)'] = ('Monthly_Rent_Paid_JOD', 'Economics')
    m['Land owned in Syria'] = ('Land_Owned_Syria', 'Economics')
    m['Land owned in Jordan'] = ('Land_Owned_Jordan', 'Economics')
    m['Self-employed (Y/N, year)'] = ('Self_Employed', 'Economics')
    m['Employed for pay (Y/N)'] = ('Employed_For_Pay', 'Economics')
    m['Looking for job (Y/N)'] = ('Looking_For_Job', 'Economics')
    m['Why not looking for job'] = ('Why_Not_Looking_Job', 'Economics')
    m['Hours searching for job (week)'] = ('Hours_Job_Search_Week', 'Economics')
    m['Were you employed? (Jan 2011)'] = ('Employed_Jan_2011', 'Economics')
    m['Change in economic situation (2 years)'] = ('Economic_Change_2yr', 'Economics')
    m['Future change in ec. situation (2 years)'] = ('Future_Economic_Change_2yr', 'Economics')
    m['Apply for a loan? (Y/N)'] = ('Applied_For_Loan', 'Economics')
    m['Taken a loan? (Y/N)'] = ('Taken_Loan', 'Economics')
    m['Amount loans taken (total, year, dinar)'] = ('Total_Loans_JOD', 'Economics')
    m['Purpose of loans; T1 codes'] = ('Loan_Purpose', 'Economics')
    m['Received gift/assistance (HH, year)'] = ('Received_Assistance', 'Economics')
    m['Any cash assistance from a govt program'] = ('Cash_Assistance_Govt', 'Economics')
    m['Any cash assistance from a UNHCR'] = ('Cash_Assistance_UNHCR', 'Economics')
    m['Any cash assistance from a WFP'] = ('Cash_Assistance_WFP', 'Economics')
    m['Learned new languages (7 years)'] = ('Languages_Learned', 'Social')

    # Food Security
    m['Consume food produced at home'] = ('Consumes_Home_Food', 'Food_Security')
    m['Consumed Cereals and cereal products (Y/N, year)'] = ('Consumed_Cereals', 'Food_Security')
    m['Consumed Meat, etc (Y/N, year)'] = ('Consumed_Meat', 'Food_Security')
    m['Consumed Fish and other seafood (Y/N, year)'] = ('Consumed_Fish', 'Food_Security')
    m['Consumed Dairy (Y/N, year)'] = ('Consumed_Dairy', 'Food_Security')
    m['Consumed Oils and fats (Y/N, year)'] = ('Consumed_Oils_Fats', 'Food_Security')
    m['Consumed Fruits and nuts (Y/N, year)'] = ('Consumed_Fruits_Nuts', 'Food_Security')
    m['Consumed Vegetables, tubers, pulses (Y/N, year)'] = ('Consumed_Vegetables', 'Food_Security')
    m['Consumed Sugar and desserts (Y/N, year)'] = ('Consumed_Sugar_Desserts', 'Food_Security')
    m['Money spend on Cereals and cereal products consumed (week)'] = ('Spend_Cereals_Week_JOD', 'Food_Security')
    m['Money spend on Meat, etc consumed (week)'] = ('Spend_Meat_Week_JOD', 'Food_Security')
    m['Money spend on Dairy consumed (week)'] = ('Spend_Dairy_Week_JOD', 'Food_Security')
    m['Meals eaten yesterday'] = ('Meals_Yesterday', 'Food_Security')
    m['Slept hungry (indiv, week)'] = ('Slept_Hungry_Indiv', 'Food_Security')
    m['Other aduls slept hungry (week)'] = ('Other_Adults_Slept_Hungry', 'Food_Security')
    m['Children slept hungry (week)'] = ('Children_Slept_Hungry', 'Food_Security')
    m['Rely on less preferred foods last week'] = ('Rely_Less_Preferred_Foods', 'Food_Security')
    m['Have no food of any kind in your household last week'] = ('No_Food_In_HH', 'Food_Security')
    m['Limit portion size at meal-times last week'] = ('Limit_Portion_Size', 'Food_Security')
    m['Reduce number of meals eaten in a day last week'] = ('Reduce_Meals', 'Food_Security')
    m['Borrow food, or reply on help from a friend or relative last week'] = ('Borrow_Food', 'Food_Security')

    # Migration
    m['Return to Syria within 2 years if conflict unresolved'] = ('Return_Syria_2yr_If_Conflict', 'Migration')
    m['Conflict will end in 2 years'] = ('Conflict_End_2yr', 'Migration')
    m['Stay in Jordan when conflict ends (Y/N)'] = ('Stay_Jordan_After_Conflict', 'Migration')
    m['Return within 1 year when conflict ends'] = ('Return_1yr_When_Conflict_Ends', 'Migration')
    m['Taken steps to move (Y/N)'] = ('Taken_Steps_To_Move', 'Migration')
    m['No. times returned to Syria, even for a day'] = ('Times_Returned_Syria', 'Migration')
    m['Lives in refugee camp?'] = ('Lives_In_Camp_Current', 'Migration')
    m['Why left Syria'] = ('Why_Left_Syria', 'Migration')
    m['Why moved'] = ('Why_Moved', 'Migration')
    m['Traveled with when moved to Jordan'] = ('Traveled_With_To_Jordan', 'Migration')
    m['When moved, who lived with'] = ('Moved_With_Whom', 'Migration')
    m['q1611. Who was already living in the neighborhood to which you moved when you fi'] = ('Neighborhood_When_Moved', 'Migration')

    # Health
    m['Fever (Y/N)'] = ('Symptom_Fever', 'Health')
    m['Persistent cough  (Y/N)'] = ('Symptom_Persistent_Cough', 'Health')
    m['Always feeling tired  (Y/N)'] = ('Symptom_Tired', 'Health')
    m['Stomach pain  (Y/N)'] = ('Symptom_Stomach_Pain', 'Health')
    m['Blood in stool  (Y/N)'] = ('Symptom_Blood_Stool', 'Health')
    m['Rapid weight loss  (Y/N)'] = ('Symptom_Weight_Loss', 'Health')
    m['Frequent diarrhea  (Y/N)'] = ('Symptom_Diarrhea', 'Health')
    m['Skin rash or irritation  (Y/N)'] = ('Symptom_Skin_Rash', 'Health')
    m['Open sores / boils  (Y/N)'] = ('Symptom_Open_Sores', 'Health')
    m['Difficulty Swallowing  (Y/N)'] = ('Symptom_Difficulty_Swallowing', 'Health')
    m['Serious wound or injury  (Y/N)'] = ('Symptom_Wound_Injury', 'Health')
    m['Typhoid  (Y/N)'] = ('Symptom_Typhoid', 'Health')
    m['Tuberculosis  (Y/N)'] = ('Symptom_Tuberculosis', 'Health')
    m['Pneumonia  (Y/N)'] = ('Symptom_Pneumonia', 'Health')
    m['Asthma / breathlessness at night  (Y/N)'] = ('Symptom_Asthma', 'Health')
    m['Frequent and excessive urination  (Y/N)'] = ('Symptom_Frequent_Urination', 'Health')
    m['Constant thirst (Y/N)'] = ('Symptom_Constant_Thirst', 'Health')
    m['Diabetes  (Y/N)'] = ('Symptom_Diabetes', 'Health')
    m['Cancer  (Y/N)'] = ('Symptom_Cancer', 'Health')
    m['Visits to hospital (month)'] = ('Hospital_Visits_Month', 'Health')
    m['Type of hospital'] = ('Hospital_Type', 'Health')
    m['Medical care spending (month)'] = ('Medical_Spending_Month_JOD', 'Health')
    m['Medicine spending (month)'] = ('Medicine_Spending_Month_JOD', 'Health')
    m['Days missed because illness (month)'] = ('Days_Missed_Illness_Month', 'Health')
    m['Days needed doctor but skipped (month)'] = ('Days_Skipped_Doctor_Month', 'Health')
    m['Why skipped doctor'] = ('Reason_Skipped_Doctor', 'Health')
    m['General health perception'] = ('General_Health', 'Health')
    m['Major health problems since Jan 2011 (Y/N)'] = ('Major_Health_Problems', 'Health')
    m['Difficulty seeing'] = ('Difficulty_Seeing', 'Health')
    m['Difficulty hearing'] = ('Difficulty_Hearing', 'Health')
    m['Difficulty walking or climbing steps'] = ('Difficulty_Walking', 'Health')
    m['Difficulty remembering or concentrating?'] = ('Difficulty_Remembering', 'Health')
    m['Difficulty with self-care'] = ('Difficulty_Self_Care', 'Health')
    m['Difficulty communicating'] = ('Difficulty_Communicating', 'Health')

    # Mental Health
    m['Unusually bothered (week)'] = ('CESD_Bothered', 'Mental_Health')
    m['Unusually unconcetrated (week)'] = ('CESD_Unconcentrated', 'Mental_Health')
    m['Depressed (week)'] = ('CESD_Depressed', 'Mental_Health')
    m['Without energy (week)'] = ('CESD_Without_Energy', 'Mental_Health')
    m['Hopeful about the future (week)'] = ('CESD_Hopeful', 'Mental_Health')
    m['Afraid (week)'] = ('CESD_Afraid', 'Mental_Health')
    m['Trouble sleeping (week)'] = ('CESD_Trouble_Sleeping', 'Mental_Health')
    m['Happy (week)'] = ('CESD_Happy', 'Mental_Health')
    m['Felt lonely (week)'] = ('CESD_Lonely', 'Mental_Health')
    m['Lacked motivation (week)'] = ('CESD_Lacked_Motivation', 'Mental_Health')
    m['Unable control things (month)'] = ('PSS_Unable_Control', 'Mental_Health')
    m['Confidence overcoming problems (month)'] = ('PSS_Confidence', 'Mental_Health')
    m['Things going their way (month)'] = ('PSS_Going_Well', 'Mental_Health')
    m['Problems were too much (month)'] = ('PSS_Problems_Too_Much', 'Mental_Health')
    m['Life satisfaction (scale)'] = ('Life_Satisfaction_scale', 'Mental_Health')
    m['Sleeping quality (month)'] = ('Sleep_Quality', 'Mental_Health')
    m['Nightmares (Y/N, month)'] = ('Nightmares', 'Mental_Health')
    m['Hard worker'] = ('Grit_Hard_Worker', 'Mental_Health')
    m['Change goals'] = ('Grit_Change_Goals', 'Mental_Health')

    # Social / Community / Political
    m['Religion'] = ('Religion', 'Social')
    m['Sect'] = ('Sect', 'Social')
    m['Importance of religion'] = ('Religion_Importance', 'Social')
    m['Community attends mosque (Y/N)'] = ('Attends_Mosque', 'Social')
    m['Children have Jordanian friends (Y/N)'] = ('Children_Jordanian_Friends', 'Social')
    m['Children share spaces with Jordanian children (Y/N)'] = ('Children_Share_Spaces_Jordanians', 'Social')
    m['Regularly attends place of worship with Jordanians'] = ('Worships_With_Jordanians', 'Social')
    m['Most identified group'] = ('Most_Identified_Group', 'Social')
    m['Opinion democracy'] = ('Opinion_Democracy', 'Social')
    m['Female work (agree)'] = ('Agree_Female_Work', 'Social')
    m['Men important decisions (agree)'] = ('Agree_Men_Decisions', 'Social')
    m['Women as leaders (agree)'] = ('Agree_Women_Leaders', 'Social')
    m['Violence in politics'] = ('Violence_In_Politics', 'Social')
    m['Women rights'] = ('Women_Rights', 'Social')
    m['Beating wives'] = ('Beating_Wives', 'Social')

    # Safety
    m['Corruption in community (Y/N)'] = ('Community_Corruption', 'Safety')
    m['Forced eviction in community (Y/N)'] = ('Community_Forced_Eviction', 'Safety')
    m['Gender violence in community (Y/N)'] = ('Community_Gender_Violence', 'Safety')
    m['discrimination in community (Y/N)'] = ('Community_Discrimination', 'Safety')
    m['Anything stolen or attempted stealing? (Y/N, year)'] = ('Been_Stolen_From', 'Safety')
    m['Times been stolen (year)'] = ('Times_Stolen_From', 'Safety')
    m['Been assaulted (Y/N, year)'] = ('Been_Assaulted', 'Safety')
    m['Times been assaulted (year)'] = ('Times_Assaulted', 'Safety')
    m['Safety walking outside'] = ('Safety_Walking_Outside', 'Safety')
    m['Compared security Jan 2011'] = ('Security_Compared_2011', 'Safety')
    m['Arrested in Syria?'] = ('Arrested_Syria', 'Safety')
    m['Imprisoned in Syria?'] = ('Imprisoned_Syria', 'Safety')

    # Marriage/Fertility
    m['Ever been married (Y/N)'] = ('Ever_Married', 'Marriage_Fertility')
    m['Age first marriage'] = ('Age_First_Marriage', 'Marriage_Fertility')
    m['Times been married'] = ('Times_Married', 'Marriage_Fertility')
    m['Been pregnant (Y/N)'] = ('Been_Pregnant', 'Marriage_Fertility')
    m['Times being pregnant'] = ('Times_Pregnant', 'Marriage_Fertility')
    return m


def m2021():
    m = {}
    m['ID'] = ('Respondent_ID', 'Demographics')
    m['Age (in years)'] = ('Age', 'Demographics')
    m['Gender'] = ('Gender', 'Demographics')
    m['Marital status'] = ('Marital_Status', 'Demographics')
    m['Has a MOI card?'] = ('Has_MOI_Card', 'Demographics')
    m['Has a passport?'] = ('Has_Passport', 'Demographics')
    m['Has a residency permit?'] = ('Has_Residency_Permit', 'Demographics')
    m['Has a work permit?'] = ('Has_Work_Permit', 'Demographics')
    m['Has a household family Book?'] = ('Has_Family_Book', 'Demographics')
    m['Has a Syrian ID?'] = ('Has_Syrian_ID', 'Demographics')
    m['Country'] = ('Country', 'Demographics')
    m['Governorate'] = ('Governorate', 'Demographics')

    m['Number of people in the household (other than FR)'] = ('HH_Size', 'Household')
    m['Head of household (Y/N)'] = ('Head_of_HH', 'Household')
    m['Currently Breastfeeding? (Y/N)'] = ('Currently_Breastfeeding', 'Household')
    m['Household chores hours (indiv, week)'] = ('HH_Chores_Hours_Week', 'Household')
    m['Childcare hours (indiv, week)'] = ('Childcare_Hours_Week', 'Household')
    m['Other people doing HH chores/childcare'] = ('Others_Doing_Chores', 'Household')
    m['Total hours HH chors/childcare (HH, week)'] = ('Total_Chores_Childcare_Hours_Week', 'Household')

    m['Number of rooms in dwelling'] = ('Number_Rooms_Dwelling', 'Dwelling')
    m['Occupies entire dwelling (Y/N)'] = ('Occupies_Entire_Dwelling', 'Dwelling')
    m['Rooms that HH occupies'] = ('Rooms_HH_Occupies', 'Dwelling')
    m['Another family living in dwelling (Y/N)'] = ('Another_Family_In_Dwelling', 'Dwelling')
    m['Number of families in dwelling'] = ('Number_Families_In_Dwelling', 'Dwelling')
    m['Made of temporary materials (Y/N)'] = ('Temporary_Materials', 'Dwelling')
    m['Prefabricated housing unit /caravan (Y/N)'] = ('Prefab_Housing', 'Dwelling')
    m['Floors materials'] = ('Floor_Material', 'Dwelling')
    m['Roof material'] = ('Roof_Material', 'Dwelling')
    m['Acces to electricity (Y/N)'] = ('Has_Electricity', 'Dwelling')
    m['Source of your electricity'] = ('Electricity_Source', 'Dwelling')
    m['Most used toilet facility'] = ('Toilet_Facility', 'Dwelling')
    m['Main water source (week)'] = ('Main_Water_Source', 'Dwelling')
    m['Source of drinking water (last week)'] = ('Drinking_Water_Source', 'Dwelling')
    m['Treat water in last 7 days?'] = ('Treats_Water', 'Dwelling')

    m['Own or rent living place'] = ('Own_Or_Rent', 'Economics')
    m['Monthly rent, dinar (agreed)'] = ('Monthly_Rent_Agreed_JOD', 'Economics')
    m['Monthly rent, dinar (paid)'] = ('Monthly_Rent_Paid_JOD', 'Economics')
    m['Jobs lost during or after the lockdown in hh'] = ('Jobs_Lost_Lockdown', 'Economics')
    m['Any loans in 2020?'] = ('Any_Loans_2020', 'Economics')
    m['Amount consumed of food gifted (in JD)'] = ('Food_Gifted_JOD', 'Economics')

    m['Attending school (Y/N)'] = ('Currently_Attending_School', 'Education')
    m['Type of school last attended (1/3)'] = ('Type_School_Last_Attended', 'Education')
    m['Year last attended school'] = ('Year_Last_Attended_School', 'Education')
    m['Type of school last attended'] = ('School_Public_Private', 'Education')
    m['Country of school last attended'] = ('Country_School_Last_Attended', 'Education')
    m['Completed any other type of education (Y/N)'] = ('Completed_Other_Education', 'Education')

    m['Meals eaten yesterday'] = ('Meals_Yesterday', 'Food_Security')
    m['Slept hungry (indiv, week)'] = ('Slept_Hungry_Indiv', 'Food_Security')
    m['Other aduls slept hungry (week)'] = ('Other_Adults_Slept_Hungry', 'Food_Security')
    m['Children slept hungry (week)'] = ('Children_Slept_Hungry', 'Food_Security')
    m['Change in food consumption (month)'] = ('Change_Food_Consumption', 'Food_Security')
    m['Change in cereal consumption (month)'] = ('Change_Cereal_Consumption', 'Food_Security')
    m['Change in meat consumption (month)'] = ('Change_Meat_Consumption', 'Food_Security')

    m['Q701a. Where were you living in in January 2011?'] = ('Location_Jan_2011', 'Migration')
    m['Have you moved since October 2019?'] = ('Moved_Since_Oct_2019', 'Migration')
    m['Plans change residence within 6 months'] = ('Plans_Move_6_Months', 'Migration')
    m[' Any plans moving next year? (Y/N)'] = ('Plans_Move_Next_Year', 'Migration')

    m['Fever (Y/N)'] = ('Symptom_Fever', 'Health')
    m['Persistent cough  (Y/N)'] = ('Symptom_Persistent_Cough', 'Health')
    m['Always feeling tired  (Y/N)'] = ('Symptom_Tired', 'Health')
    m['Muscle Pain (myalgia)  (Y/N)'] = ('Symptom_Muscle_Pain', 'Health')
    m['Headache  (Y/N)'] = ('Symptom_Headache', 'Health')
    m['Diarrhea/Nausea/Vomiting  (Y/N)'] = ('Symptom_Diarrhea_Nausea', 'Health')
    m['Difficulty breathing/ Chest tightness  (Y/N)'] = ('Symptom_Breathing_Difficulty', 'Health')
    m['Runny nose  (Y/N)'] = ('Symptom_Runny_Nose', 'Health')
    m['Pneumonia  (Y/N)'] = ('Symptom_Pneumonia', 'Health')
    m['Loss of sense of smell / not being able to taste food  (Y/N)'] = ('Symptom_Loss_Smell_Taste', 'Health')
    m['Would you describe your general health as good, fair, poor, or very poor?'] = ('General_Health', 'Health')
    m['Health problems that seriously affected your life or work, since October 2019'] = ('Major_Health_Problems', 'Health')
    m['Difficulty seeing'] = ('Difficulty_Seeing', 'Health')
    m['Difficulty hearing'] = ('Difficulty_Hearing', 'Health')
    m['Difficulty walking or climbing steps'] = ('Difficulty_Walking', 'Health')
    m['Difficulty remembering or concentrating?'] = ('Difficulty_Remembering', 'Health')
    m['Difficulty with self-care'] = ('Difficulty_Self_Care', 'Health')
    m['Difficulty communicating'] = ('Difficulty_Communicating', 'Health')

    m['Unusually bothered (week)'] = ('CESD_Bothered', 'Mental_Health')
    m['Unusually unconcetrated (week)'] = ('CESD_Unconcentrated', 'Mental_Health')
    m['Depressed (week)'] = ('CESD_Depressed', 'Mental_Health')
    m['Without energy (week)'] = ('CESD_Without_Energy', 'Mental_Health')
    m['Hopeful about the future (week)'] = ('CESD_Hopeful', 'Mental_Health')
    m['Afraid (week)'] = ('CESD_Afraid', 'Mental_Health')
    m['Trouble sleeping (week)'] = ('CESD_Trouble_Sleeping', 'Mental_Health')
    m['Happy (week)'] = ('CESD_Happy', 'Mental_Health')
    m['Felt lonely (week)'] = ('CESD_Lonely', 'Mental_Health')
    m['Lacked motivation (week)'] = ('CESD_Lacked_Motivation', 'Mental_Health')
    m['Unable control things (month)'] = ('PSS_Unable_Control', 'Mental_Health')
    m['Confidence overcoming problems (month)'] = ('PSS_Confidence', 'Mental_Health')
    m['Things going their way (month)'] = ('PSS_Going_Well', 'Mental_Health')
    m['Problems were too much (month)'] = ('PSS_Problems_Too_Much', 'Mental_Health')

    m['Stayed alone at home (week)'] = ('COVID_Stayed_Home', 'Health')
    m['Attended social gatherings (week)'] = ('COVID_Social_Gatherings', 'Health')
    m['Kept distance to outside hh (week)'] = ('COVID_Kept_Distance', 'Health')

    m['a. Gender based violence'] = ('Gender_based_violence', 'Safety')
    m['b. Child protection'] = ('Knows_Child_Protection_Hotline', 'Safety')
    return m


def m2022():
    m = {}
    m['ID'] = ('Respondent_ID', 'Demographics')
    m['FR Age'] = ('Age', 'Demographics')
    m['q202c. To interviewer: please note FR gender.'] = ('Gender', 'Demographics')
    m['Q202b. What is your marital status?'] = ('Marital_Status', 'Demographics')
    m['Q205a. Do you have MOI card?'] = ('Has_MOI_Card', 'Demographics')
    m['Q205c. Do you have residency permit?'] = ('Has_Residency_Permit', 'Demographics')
    m['Q205d. Do you have work permit?'] = ('Has_Work_Permit', 'Demographics')
    m['Q205e. Do you have family Book?'] = ('Has_Family_Book', 'Demographics')
    m['Q205f. Do you have Syrian ID?'] = ('Has_Syrian_ID', 'Demographics')
    m['Q205g. Do you have UNHCR file?'] = ('Has_UNHCR_File', 'Demographics')
    m['Q204b. Of what country(ies) are you a citizen?: Syria?'] = ('Citizenship_Syrian', 'Demographics')
    m['Q204b. Of what country(ies) are you a citizen?: Jordan?'] = ('Citizenship_Jordanian', 'Demographics')
    m['Interviewee alone (Y/N)'] = ('Interviewee_Alone', 'Demographics')

    m['Q401.Are you currently attending school?'] = ('Currently_Attending_School', 'Education')
    m['Type of school last attended (1/3)'] = ('Type_School_Last_Attended', 'Education')
    m['q402a2. In what year did you last attend this type of school?'] = ('Year_Last_Attended_School', 'Education')
    m['Q402a. In your last year of study (${q402}), were you at a public or private sch'] = ('School_Public_Private', 'Education')
    m['Q403. In what country ${do_did} this type of school?'] = ('Country_School_Last_Attended', 'Education')
    m['Completed any other type of education (Y/N)'] = ('Completed_Other_Education', 'Education')
    m['q404a. Years:'] = ('Length_Schooling_Years', 'Education')
    m['q404b.Months:'] = ('Length_Schooling_Months', 'Education')

    m['Sum the hh size with the members who still living in the hh'] = ('HH_Size', 'Household')
    m['Head of the household age'] = ('Head_HH_Age', 'Household')
    m['Household size in the last round'] = ('HH_Size_Last_Round', 'Household')
    return m


def m2023():
    m = {}
    m['ID'] = ('Respondent_ID', 'Demographics')
    m['Q202A . How old are you (in years)?'] = ('Age', 'Demographics')
    m['Q203 . To interviewer: please note FR gender.'] = ('Gender', 'Demographics')
    m['Q204 . What is your marital status?'] = ('Marital_Status', 'Demographics')
    m['Highest level of education completed'] = ('Highest_Education', 'Education')
    m['Q206 . What is your citizenship?: Syrian?'] = ('Citizenship_Syrian', 'Demographics')
    m['Q206 . What is your citizenship?: Jordanian?'] = ('Citizenship_Jordanian', 'Demographics')
    m['Country_last_round'] = ('Country', 'Demographics')
    m['Governorat_last_round'] = ('Governorate', 'Demographics')

    m['total_family_member'] = ('Total_Family_Members', 'Household')
    m['household_roster_count'] = ('HH_Roster_Count', 'Household')

    m['Any work to generate income, even for an hour (Month)'] = ('Worked_Last_Month', 'Economics')
    m['Q501ai. Are you currently attending school?'] = ('Currently_Attending_School', 'Education')
    m['Have you ever done work for pay to generate income'] = ('Ever_Worked_For_Pay', 'Economics')
    m['Occupation, most important primary position'] = ('Primary_Occupation', 'Economics')
    m['Was this work self-employment, informal or formal with a contract?'] = ('Work_Type_Contract', 'Economics')
    m['Hours worked in a typical week(Month)'] = ('Hours_Worked_Week', 'Economics')
    m['Q504. In the past 30 days, what were your total earnings from all forms of work?'] = ('Total_Earnings_30Days_JOD', 'Economics')
    m['Hours spend actively searching for a job(Week)'] = ('Hours_Job_Search_Week', 'Economics')
    m['Cash assistance hh received (Year)'] = ('Cash_Assistance_Received_JOD', 'Economics')
    m['In-kind cash assistance (Year)'] = ('InKind_Assistance_JOD', 'Economics')

    m['Q509a. In the past 30 days, how many times did you host a neighbor for a meal or'] = ('Times_Hosted_Neighbor', 'Social')
    m['Q509b. How many of these times were with a Jordanian neighbor?'] = ('Times_Hosted_Jordanian', 'Social')

    m['Q601A. Where were you (), living in in January 2011?'] = ('Location_Jan_2011', 'Migration')
    return m


def m2024():
    m = {}
    m['Unique ID'] = ('Respondent_ID', 'Demographics')
    m['q203. To interviewer: please note FR gender.'] = ('Gender', 'Demographics')
    m['q204. What is your marital status?'] = ('Marital_Status', 'Demographics')
    m['Q205. What is the highest level of education you successfully completed?'] = ('Highest_Education', 'Education')
    m['Are you currently attending school?'] = ('Currently_Attending_School', 'Education')
    m[' Are you able to read and write?'] = ('Can_Read_Write', 'Education')
    m['In what year did you last attend this type of school?'] = ('Year_Last_Attended_School', 'Education')
    m['For how long did you attend this type of schooling? (years)'] = ('Length_Schooling_Years', 'Education')
    m['For how long did you attend this type of schooling? (months)'] = ('Length_Schooling_Months', 'Education')

    m['Q301A. Which country?'] = ('Country', 'Demographics')
    m['Q301B. Which governorate?'] = ('Governorate', 'Demographics')
    m['Lives in refugee camp?'] = ('Lives_In_Camp_Current', 'Migration')

    # Household
    m['Number of people in the household (other than FR)'] = ('HH_Size', 'Household')
    m['Head of household (Y/N)'] = ('Head_of_HH', 'Household')
    m['Currently breastfeeding (Y/N)'] = ('Currently_Breastfeeding', 'Household')
    m['Household chores hours (indiv, week)'] = ('HH_Chores_Hours_Week', 'Household')
    m['Childcare hours (indiv, week)'] = ('Childcare_Hours_Week', 'Household')
    m['Other people doing HH chores/childcare'] = ('Others_Doing_Chores', 'Household')
    m['Total hours HH chors/childcare (HH, week)'] = ('Total_Chores_Childcare_Hours_Week', 'Household')

    # Dwelling
    m['Number of rooms in dwelling'] = ('Number_Rooms_Dwelling', 'Dwelling')
    m['Occupies entire dwelling (Y/N)'] = ('Occupies_Entire_Dwelling', 'Dwelling')
    m['Rooms that HH occupies'] = ('Rooms_HH_Occupies', 'Dwelling')
    m['Another family living in dwelling (Y/N)'] = ('Another_Family_In_Dwelling', 'Dwelling')
    m['Number of families in dwelling'] = ('Number_Families_In_Dwelling', 'Dwelling')
    m['Made of temporary materials (Y/N)'] = ('Temporary_Materials', 'Dwelling')
    m['Prefabricated housing unit /caravan (Y/N)'] = ('Prefab_Housing', 'Dwelling')
    m['Floors materials'] = ('Floor_Material', 'Dwelling')
    m['Roof material'] = ('Roof_Material', 'Dwelling')
    m['Acces to electricity (Y/N)'] = ('Has_Electricity', 'Dwelling')
    m['Most used toilet facility'] = ('Toilet_Facility', 'Dwelling')
    m['Main water source (week)'] = ('Main_Water_Source', 'Dwelling')
    m['Source of drinking water (last week)'] = ('Drinking_Water_Source', 'Dwelling')
    m['Source of your electricity'] = ('Electricity_Source', 'Dwelling')
    m['Water bill last month (Dinars)'] = ('Water_Bill_Month_JOD', 'Dwelling')

    # Economics
    m['Own or rent living place'] = ('Own_Or_Rent', 'Economics')
    m['Monthly rent, dinar (agreed)'] = ('Monthly_Rent_Agreed_JOD', 'Economics')
    m['Monthly rent, dinar (paid)'] = ('Monthly_Rent_Paid_JOD', 'Economics')
    m['Land owned in Syria'] = ('Land_Owned_Syria', 'Economics')
    m['Land owned in Jordan'] = ('Land_Owned_Jordan', 'Economics')
    m['Self-employed (Y/N, year)'] = ('Self_Employed', 'Economics')
    m['Employed for pay (Y/N)'] = ('Employed_For_Pay', 'Economics')
    m['Worked or volunteered (Y/N, year)'] = ('Worked_Or_Volunteered', 'Economics')
    m['Employment status, position 1'] = ('Employment_Status', 'Economics')
    m['Working patterns, position 1'] = ('Working_Patterns', 'Economics')
    m['Hours worked (last week), position 1'] = ('Hours_Worked_Week', 'Economics')
    m['Pre-tax income (year), position 1'] = ('Pretax_Income_Year_JOD', 'Economics')
    m['Looking for job (Y/N)'] = ('Looking_For_Job', 'Economics')
    m['Why not looking for job'] = ('Why_Not_Looking_Job', 'Economics')
    m['Hours searching for job (week)'] = ('Hours_Job_Search_Week', 'Economics')
    m['Were you employed? (Jan 2011)'] = ('Employed_Jan_2011', 'Economics')
    m['Adults working in the hh (hours,month)'] = ('Adults_Working_Hours_Month', 'Economics')
    m['Adults total income in the hh (hours,month)'] = ('Adults_Total_Income_Month_JOD', 'Economics')
    m['Received gift/assistance (HH, year)'] = ('Received_Assistance', 'Economics')
    m['Any cash assistance from a govt program'] = ('Cash_Assistance_Govt', 'Economics')
    m['Any cash assistance from a UNHCR'] = ('Cash_Assistance_UNHCR', 'Economics')
    m['Any cash assistance from a WFP'] = ('Cash_Assistance_WFP', 'Economics')
    m['Apply for a loan? (Y/N)'] = ('Applied_For_Loan', 'Economics')
    m['Taken a loan? (Y/N)'] = ('Taken_Loan', 'Economics')
    m['Amount loans taken (total, year, dinar)'] = ('Total_Loans_JOD', 'Economics')
    m['Aware of Jordanian work permits?'] = ('Aware_Work_Permits', 'Economics')
    m['Currently have a valid Jordanian work permit?'] = ('Has_Valid_Work_Permit', 'Economics')
    m['Is mobility important for Syrian refugees?'] = ('Mobility_Important', 'Economics')
    m['Purpose of loans; T1 codes'] = ('Loan_Purpose', 'Economics')
    m['Workforce barriers to entry for Syrian women'] = ('Workforce_Barriers_Women', 'Economics')

    # Food Security
    m['Consumed Cereals and cereal products (Y/N, year)'] = ('Consumed_Cereals', 'Food_Security')
    m['Consumed Meat, etc (Y/N, year)'] = ('Consumed_Meat', 'Food_Security')
    m['Consumed Fish and other seafood (Y/N, year)'] = ('Consumed_Fish', 'Food_Security')
    m['Consumed Dairy (Y/N, year)'] = ('Consumed_Dairy', 'Food_Security')
    m['Consumed Oils and fats (Y/N, year)'] = ('Consumed_Oils_Fats', 'Food_Security')
    m['Consumed Fruits and nuts (Y/N, year)'] = ('Consumed_Fruits_Nuts', 'Food_Security')
    m['Consumed Vegetables, tubers, pulses (Y/N, year)'] = ('Consumed_Vegetables', 'Food_Security')
    m['Consumed Sugar and desserts (Y/N, year)'] = ('Consumed_Sugar_Desserts', 'Food_Security')
    m['Money spend on Cereals and cereal products consumed (week)'] = ('Spend_Cereals_Week_JOD', 'Food_Security')
    m['Money spend on Meat, etc consumed (week)'] = ('Spend_Meat_Week_JOD', 'Food_Security')
    m['Money spend on Dairy consumed (week)'] = ('Spend_Dairy_Week_JOD', 'Food_Security')
    m['Money spend on Oils and fats consumed (week)'] = ('Spend_Oils_Week_JOD', 'Food_Security')
    m['Money spend on Fruits and nuts consumed (week)'] = ('Spend_Fruits_Week_JOD', 'Food_Security')
    m['Money spend on Vegetables, tubers, pulses consumed (week)'] = ('Spend_Vegetables_Week_JOD', 'Food_Security')
    m['Money spend on Sugar and desserts consumed (week)'] = ('Spend_Sugar_Week_JOD', 'Food_Security')
    m['Meals eaten yesterday'] = ('Meals_Yesterday', 'Food_Security')
    m['Slept hungry (indiv, week)'] = ('Slept_Hungry_Indiv', 'Food_Security')
    m['Other aduls slept hungry (week)'] = ('Other_Adults_Slept_Hungry', 'Food_Security')
    m['Children slept hungry (week)'] = ('Children_Slept_Hungry', 'Food_Security')
    m['Rely on less preferred foods last week'] = ('Rely_Less_Preferred_Foods', 'Food_Security')
    m['Have no food of any kind in your household last week'] = ('No_Food_In_HH', 'Food_Security')
    m['Limit portion size at meal-times last week'] = ('Limit_Portion_Size', 'Food_Security')
    m['Reduce number of meals eaten in a day last week'] = ('Reduce_Meals', 'Food_Security')
    m['Borrow food, or reply on help from a friend or relative last week'] = ('Borrow_Food', 'Food_Security')

    # Health
    m['Health problem: Fever'] = ('Symptom_Fever', 'Health')
    m['Health problem: Persistent cough'] = ('Symptom_Persistent_Cough', 'Health')
    m['Health problem: Always feeling tired'] = ('Symptom_Tired', 'Health')
    m['Health problem: Muscle pain (myalgia)'] = ('Symptom_Muscle_Pain', 'Health')
    m['Health problem: Headache/migraine'] = ('Symptom_Headache', 'Health')
    m['Health problem: Stomach pain'] = ('Symptom_Stomach_Pain', 'Health')
    m['Health problem: Blood in stool'] = ('Symptom_Blood_Stool', 'Health')
    m['Health problem: Rapid weight loss'] = ('Symptom_Weight_Loss', 'Health')
    m['Health problem: Open sores / boils'] = ('Symptom_Open_Sores', 'Health')
    m['Health problem: Skin rash or irritation'] = ('Symptom_Skin_Rash', 'Health')
    m['Health problem: Difficulty swallowing'] = ('Symptom_Difficulty_Swallowing', 'Health')
    m['Health problem: Pneumonia'] = ('Symptom_Pneumonia', 'Health')
    m['Health problem: Frequent and excessive urination'] = ('Symptom_Frequent_Urination', 'Health')
    m['Health problem: Constant thirst / increased drinking of fluids'] = ('Symptom_Constant_Thirst', 'Health')
    m['Health problem: Difficulty breathing / chest tightness'] = ('Symptom_Breathing_Difficulty', 'Health')
    m['Health problem: Diarrhea / nausea / vomiting'] = ('Symptom_Diarrhea_Nausea', 'Health')
    m['Health problem: Loss of sense of smell / not being able to taste food'] = ('Symptom_Loss_Smell_Taste', 'Health')
    m['Health problem: Back pain or other muscle pain'] = ('Symptom_Back_Pain', 'Health')
    m['Health problem: Runny nose'] = ('Symptom_Runny_Nose', 'Health')
    m['Health problem: Sore throat'] = ('Symptom_Sore_Throat', 'Health')
    m['Health problem: Fast or irregular heartbeat '] = ('Symptom_Heart', 'Health')
    m['Illiness: Diabetes'] = ('Symptom_Diabetes', 'Health')
    m['Illiness: Cancer'] = ('Symptom_Cancer', 'Health')
    m['Visits to hospital (month)'] = ('Hospital_Visits_Month', 'Health')
    m['Hospital / clinic medical care spending (month)'] = ('Medical_Spending_Month_JOD', 'Health')
    m['Medicine spending (month)'] = ('Medicine_Spending_Month_JOD', 'Health')
    m['Days of work missed due to poor health (month)'] = ('Days_Missed_Illness_Month', 'Health')
    m['Days felt should but didn\'t visit doctor (month)'] = ('Days_Skipped_Doctor_Month', 'Health')
    m['General health'] = ('General_Health', 'Health')
    m['Major health problems since Jan 2011 (Y/N)'] = ('Major_Health_Problems', 'Health')
    m['Difficulty seeing'] = ('Difficulty_Seeing', 'Health')
    m['Difficulty hearing'] = ('Difficulty_Hearing', 'Health')
    m['Difficulty walking or climbing steps'] = ('Difficulty_Walking', 'Health')
    m['Difficulty remembering or concentrating?'] = ('Difficulty_Remembering', 'Health')
    m['Difficulty with self-care'] = ('Difficulty_Self_Care', 'Health')
    m['Difficulty communicating'] = ('Difficulty_Communicating', 'Health')
    m['List health problems'] = ('Health_Problems_List', 'Health')

    # Mental Health
    m['Q17.3.03. How would you characterize your sleep quality in the last 30 days?'] = ('Sleep_Quality', 'Mental_Health')
    m['q15.6.1. Taking everything together, would you say you are somewhat happy, very '] = ('Happiness', 'Mental_Health')
    m['q15.6.02a. On a scale of 1-7, with 1 being sad and 7 being happy, how do you fee'] = ('Happiness_Scale_1_7', 'Mental_Health')

    # Migration
    m['Taken steps to move (Y/N)'] = ('Taken_Steps_To_Move', 'Migration')
    m['No. times returned to Syria, even for a day'] = ('Times_Returned_Syria', 'Migration')
    m['Lives in refugee camp?'] = ('Lives_In_Camp_Current', 'Migration')
    m['Why left Syria'] = ('Why_Left_Syria', 'Migration')
    m['Why moved'] = ('Why_Moved', 'Migration')
    m['Traveled with when moved to Jordan (who)'] = ('Traveled_With_To_Jordan', 'Migration')
    m['When moved, who lived with'] = ('Moved_With_Whom', 'Migration')

    # Social
    m['Religion'] = ('Religion', 'Social')
    m['Sect'] = ('Sect', 'Social')
    m['Importance of religion'] = ('Religion_Importance', 'Social')
    m['Children have Jordanian friends (Y/N)'] = ('Children_Jordanian_Friends', 'Social')
    m['Children share spaces with Jordanian children (Y/N)'] = ('Children_Share_Spaces_Jordanians', 'Social')
    m['A married woman can work outside the home if she wishes'] = ('Agree_Female_Work', 'Social')
    m['Languages spoker'] = ('Languages_Spoken', 'Social')

    # Safety
    m['Anything stolen or attempted stealing? (Y/N, year)'] = ('Been_Stolen_From', 'Safety')
    m['Times been stolen (year)'] = ('Times_Stolen_From', 'Safety')
    m['Been assaulted (Y/N, year)'] = ('Been_Assaulted', 'Safety')
    m['Safety walking outside'] = ('Safety_Walking_Outside', 'Safety')
    m['Arrested in Syria?'] = ('Arrested_Syria', 'Safety')

    # Marriage/Fertility
    m['Ever been married (Y/N)'] = ('Ever_Married', 'Marriage_Fertility')
    m['Age first marriage'] = ('Age_First_Marriage', 'Marriage_Fertility')
    m['Times been married'] = ('Times_Married', 'Marriage_Fertility')
    m['Been pregnant (Y/N)'] = ('Been_Pregnant', 'Marriage_Fertility')
    m['Times being pregnant'] = ('Times_Pregnant', 'Marriage_Fertility')
    return m


###############################################################################
# DATA LOADING
###############################################################################

def _fuzzy_find_col(orig, df_columns):
    """Find best matching column name, preferring exact → strip → case → partial."""
    if orig in df_columns:
        return orig
    # Strip whitespace
    for c in df_columns:
        if c.strip() == orig.strip():
            return c
    # Case-insensitive
    for c in df_columns:
        if c.strip().lower() == orig.strip().lower():
            return c
    # Partial match for long column names that got truncated
    # Must match at least 30 chars AND be the longest match
    if len(orig) > 30:
        best = None
        best_len = 0
        prefix = orig[:30]
        for c in df_columns:
            if c.strip().startswith(prefix):
                # Check how much of orig matches c
                match_len = 0
                cs = c.strip()
                for i in range(min(len(orig), len(cs))):
                    if orig[i] == cs[i]:
                        match_len = i + 1
                    else:
                        break
                if match_len > best_len:
                    best = c
                    best_len = match_len
        return best
    return None


def load_year(filepath, year, mapping):
    print(f"  Loading {year}...")
    df_raw = pd.read_excel(filepath)
    print(f"    Raw: {df_raw.shape[0]} rows x {df_raw.shape[1]} cols")

    available = {}
    for orig, (std, topic) in mapping.items():
        match = _fuzzy_find_col(orig, df_raw.columns)
        if match is not None:
            available[match] = (std, topic)

    rename_map = {o: s for o, (s, _) in available.items()}
    result = df_raw[list(rename_map.keys())].rename(columns=rename_map)
    result['Survey_Year'] = year
    topic_map = {s: t for _, (s, t) in available.items()}

    # Build code→label maps from adjacent binary columns
    code_maps = build_code_label_map(df_raw, mapping)

    # Resolve "Other (specify)" values from adjacent raw columns
    resolve_other_specify(result, df_raw, mapping, code_maps)

    # Aggregate roster/repeating columns (positions, businesses)
    agg = aggregate_roster_columns(df_raw, year)
    for col_name, series in agg.items():
        result[col_name] = series.values
        topic_map[col_name] = 'Economics'  # roster aggregates are economics-related

    # Nullify sentinel values in all numeric columns
    nullify_sentinels(result)

    print(f"    Selected: {result.shape[1]-1} data cols + Survey_Year")
    return result, topic_map, code_maps


###############################################################################
# COMBINE YEAR+MONTH → DECIMAL YEARS
###############################################################################

YEAR_MONTH_PAIRS = [
    ('Length_Schooling_Years', 'Length_Schooling_Months', 'Length_Schooling_Total_Years'),
]

def combine_year_month(df):
    for yr_col, mo_col, combined_col in YEAR_MONTH_PAIRS:
        if yr_col in df.columns and mo_col in df.columns:
            y = mask_sentinels(df[yr_col]).fillna(0)
            mo = mask_sentinels(df[mo_col]).fillna(0)
            df[combined_col] = y + mo / 12.0
            df[combined_col] = df[combined_col].round(2)
            # If both original values were NaN/sentinel, result should be NaN
            both_missing = df[yr_col].isna() | pd.to_numeric(df[yr_col], errors='coerce').isin(SENTINEL_VALUES)
            both_missing = both_missing & (df[mo_col].isna() | pd.to_numeric(df[mo_col], errors='coerce').isin(SENTINEL_VALUES))
            df.loc[both_missing, combined_col] = np.nan
            print(f"    Combined {yr_col} + {mo_col} → {combined_col}")
    return df


###############################################################################
# NORMALIZE CURRENCY TO JOD
###############################################################################

# Approximate exchange rate for Syrian Lira → JOD (varies by year)
SYP_TO_JOD = {2020: 0.00093, 2021: 0.00047, 2022: 0.00028,
               2023: 0.00011, 2024: 0.00011}

def normalize_currency(df):
    """Already most values are in JOD. Flag any that aren't."""
    # For 2020 Jan 2011 income (was in SYP), convert
    if 'Monthly_Income_Jan_2011' in df.columns:
        # Jan 2011 rate ~47 SYP/JOD
        mask = df['Survey_Year'] == 2020
        df.loc[mask, 'Monthly_Income_Jan_2011_JOD'] = pd.to_numeric(
            df.loc[mask, 'Monthly_Income_Jan_2011'], errors='coerce') / 47.0
    # Most columns are already in JOD/Dinar per column naming
    # Columns ending in _JOD are already normalized
    print("    Currency normalized (all monetary columns in JOD)")
    return df


###############################################################################
# STANDARDIZE TIMESTAMPS TO 24HR
###############################################################################

def standardize_timestamps(df):
    """Convert any timestamp columns to 24hr format strings."""
    for col in df.columns:
        if df[col].dtype == 'datetime64[ns]':
            df[col] = df[col].dt.strftime('%Y-%m-%d %H:%M')
        elif 'time' in col.lower() and col not in ('Survey_Year',):
            # Check if numeric (Excel serial time)
            sample = df[col].dropna().head(5)
            if len(sample) > 0:
                if pd.api.types.is_numeric_dtype(sample):
                    # Excel serial time: fraction of day
                    numeric_vals = pd.to_numeric(df[col], errors='coerce')
                    mask = numeric_vals.notna() & (numeric_vals > 0) & (numeric_vals < 1)
                    if mask.sum() > 0:
                        total_seconds = numeric_vals[mask] * 86400
                        hours = (total_seconds // 3600).astype(int)
                        minutes = ((total_seconds % 3600) // 60).astype(int)
                        df.loc[mask, col] = hours.astype(str).str.zfill(2) + ':' + minutes.astype(str).str.zfill(2)
    print("    Timestamps standardized to 24hr format")
    return df


###############################################################################
# CLEAN ENCODING ARTIFACTS
###############################################################################

def clean_encoding(df):
    for col in df.select_dtypes(include=['object']).columns:
        df[col] = df[col].apply(lambda x: _clean(x) if isinstance(x, str) else x)
    return df

def _clean(text):
    if not isinstance(text, str):
        return text
    text = text.replace('\u00c2', '').replace('\u00c3', '')
    text = text.replace('Â', '').replace('Ã', '')
    text = text.replace('\u200b', '').replace('\u200c', '').replace('\u200d', '')
    text = text.replace('\ufeff', '')
    text = re.sub(r'\x80[\x90-\x9f]', '', text)  # Windows-1252 artifacts
    text = re.sub(r'_x[0-9a-fA-F]{4}_', '', text)  # Excel escape sequences like _x0080_
    text = re.sub(r'  +', ' ', text).strip()
    return text if text else np.nan


###############################################################################
# NULLIFY SENTINEL VALUES IN ALL NUMERIC COLUMNS
###############################################################################

def nullify_sentinels(df):
    """Replace sentinel values (-99, -88, etc.) with NaN across all numeric cols."""
    count = 0
    for col in df.columns:
        if pd.api.types.is_numeric_dtype(df[col]):
            mask = df[col].isin(SENTINEL_VALUES)
            n = mask.sum()
            if n > 0:
                df.loc[mask, col] = np.nan
                count += n
    if count > 0:
        print(f"    Nullified {count} sentinel values across numeric columns")


###############################################################################
# RESOLVE "OTHER (SPECIFY)" VALUES FROM ADJACENT RAW COLUMNS
###############################################################################

def _is_arabic(text):
    """Check if text is primarily Arabic/non-Latin characters."""
    if not isinstance(text, str) or not text.strip():
        return False
    latin_count = sum(1 for c in text if c.isascii() and c.isalpha())
    total_alpha = sum(1 for c in text if c.isalpha())
    if total_alpha == 0:
        return False
    return latin_count / total_alpha < 0.5


def resolve_other_specify(result, df_raw, mapping, code_maps):
    """For coded columns where a code maps to 'other', find the actual
    specify text from adjacent raw columns.

    Patterns found in the data:
      1) Binary col "Source of electricity: other" → next col "Other source, specify"
      2) Arabic specify col → next col "... Code 1" with English translation
      3) Main question col → next col "Other (specify)" or "other. Other (specify)"
    """
    raw_cols = list(df_raw.columns)
    other_resolved = 0

    for orig_col, (std_name, _topic) in mapping.items():
        real_col = _fuzzy_find_col(orig_col, raw_cols)
        if real_col is None or std_name not in result.columns:
            continue

        idx = raw_cols.index(real_col)
        cmap = code_maps.get(std_name, {})

        # Find the "other" code number (if this is a coded column)
        other_code = None
        for code_num, label in cmap.items():
            if 'other' in label.lower():
                other_code = code_num
                break

        # Strategy 1: Look for "Other (specify)" or "other. Other (specify)" columns
        # near the main question column
        specify_col = None
        english_col = None
        search_range = min(idx + 50, len(raw_cols))

        for j in range(idx + 1, search_range):
            cname = raw_cols[j].lower()
            # Stop if we hit the next mapped question
            if j > idx + 2 and any(_fuzzy_find_col(o, [raw_cols[j]]) is not None
                                     for o in mapping.keys() if o != orig_col):
                break

            if 'other' in cname and ('specify' in cname or 'specif' in cname):
                if 'code 1' in cname or 'coded' in cname.replace('code 1', ''):
                    english_col = raw_cols[j]
                elif specify_col is None:
                    specify_col = raw_cols[j]
                    # Check if next col is "Code 1" (English translation)
                    if j + 1 < len(raw_cols) and 'code 1' in raw_cols[j + 1].lower():
                        english_col = raw_cols[j + 1]

        # If we found a specify column, use it to fill in "other" values
        if specify_col is None and english_col is None:
            continue

        # Prefer English "Code 1" column, fall back to specify column
        source_col = english_col if english_col is not None else specify_col

        for row_idx in range(len(result)):
            val = result.at[row_idx, std_name]
            specify_val = df_raw.at[row_idx, source_col] if row_idx < len(df_raw) else None

            if pd.isna(specify_val) or (isinstance(specify_val, str) and not specify_val.strip()):
                continue
            if isinstance(specify_val, str) and _is_arabic(specify_val):
                # Try the English column if we haven't already
                if source_col != english_col and english_col is not None:
                    eng_val = df_raw.at[row_idx, english_col] if row_idx < len(df_raw) else None
                    if pd.notna(eng_val) and isinstance(eng_val, str) and eng_val.strip():
                        specify_val = eng_val
                    else:
                        continue  # Skip Arabic-only values
                else:
                    continue  # Skip Arabic-only values

            # Clean the specify value
            specify_val = _clean(str(specify_val).strip()) if isinstance(specify_val, str) else str(specify_val)
            if not specify_val or specify_val == 'nan':
                continue

            # For coded columns: if value matches "other" code, replace with specify text
            if other_code is not None:
                raw_val = str(val).strip() if pd.notna(val) else ''
                if raw_val == str(other_code) or raw_val.lower() in ('other', 'other (specify)'):
                    result.at[row_idx, std_name] = specify_val
                    other_resolved += 1
                elif ' ' in raw_val:
                    # Multi-code: replace the "other" code within the string
                    parts = raw_val.split()
                    parts = [specify_val if p == str(other_code) else p for p in parts]
                    result.at[row_idx, std_name] = ' '.join(parts)
                    other_resolved += 1
            else:
                # Non-coded column: if current value is "other" / "Other", replace
                if pd.notna(val) and str(val).strip().lower() in ('other', 'other (specify)', 'other, specify'):
                    result.at[row_idx, std_name] = specify_val
                    other_resolved += 1

    if other_resolved > 0:
        print(f"    Resolved {other_resolved} 'Other (specify)' values with actual text")


###############################################################################
# AUTO-DECODE MULTI-RESPONSE CODES  →  HUMAN-READABLE LABELS
###############################################################################

# Map: (raw_col_in_source, std_col_name) → ordered list of binary column
# name prefixes.  The index in the list = code number (1-based).
# Built automatically by build_code_label_map() at load time.

def build_code_label_map(df_raw, mapping):
    """Auto-detect coded multi-response columns and build code→label dicts.

    For a coded column like 'Source of your electricity' followed by
    binary columns 'Source of electricity: Connection to grid',
    'Source of electricity: Generator', etc., we extract:
      {1: 'Connection to grid', 2: 'Generator', ...}

    Returns {std_col_name: {int_code: label_str, ...}}
    """
    cols = list(df_raw.columns)
    space_num = re.compile(r'^-?\d+(\s+-?\d+)*$')
    code_maps = {}

    for orig_col, (std_name, _topic) in mapping.items():
        real_col = _fuzzy_find_col(orig_col, cols)
        if real_col is None:
            continue

        # Check if column has space-separated numeric codes
        sample = df_raw[real_col].dropna().astype(str).head(20)
        coded_count = sum(1 for v in sample if space_num.match(str(v).strip()))
        single_num = sum(1 for v in sample if re.match(r'^-?\d+$', str(v).strip()))
        if coded_count == 0 and single_num < 3:
            continue

        # Look at adjacent binary (Yes/No) columns
        idx = cols.index(real_col)
        binary_labels = []
        for j in range(idx + 1, min(idx + 40, len(cols))):
            next_col = cols[j]
            vals = df_raw[next_col].dropna().astype(str).head(10).tolist()
            has_yn = any(v in ('Yes', 'No', 'yes', 'no') for v in vals)
            if has_yn:
                binary_labels.append(next_col)
            else:
                break

        if len(binary_labels) < 2:
            continue  # Not a coded multi-response column

        # Extract labels from binary column names
        # Pattern: "Toilet most used: Latrine" → "Latrine"
        # Pattern: "Loan for: School fees" → "School fees"
        # Pattern: "Felt unsafe in home/area was dangerous" → full name
        labels = {}
        for code_num, bcol in enumerate(binary_labels, start=1):
            # Strip .N suffix for repeated roster cols
            label = re.sub(r'\.\d+$', '', bcol).strip()
            # Try to extract label after colon
            if ':' in label:
                label = label.split(':', 1)[1].strip()
            # Remove common prefixes
            for prefix in ['Lived with ', 'Traveled with when moved to Jordan: ',
                           'Languages spoken: ', 'Languages learned in last 7 yrs: ',
                           'Health problem: ', 'Loan for: ',
                           'in neighborhood when moved to Jordan: ']:
                if label.startswith(prefix):
                    label = label[len(prefix):]
                    break
            labels[code_num] = label

        code_maps[std_name] = labels
        # print(f"    Code map for '{std_name}': {len(labels)} codes")

    return code_maps


def decode_multi_response(df, code_maps):
    """Decode space-separated numeric codes to human-readable labels.

    All coded columns are decoded in-place as comma-separated labels.
    No row splitting — one row per respondent (wide format).
    """

    DECODE_COLS = [
        'Why_Left_Syria', 'Why_Moved', 'Reasons_Not_Registered_UNHCR',
        'Reason_Skipped_Doctor',
        'Electricity_Source', 'Toilet_Facility', 'Loan_Purpose',
        'Traveled_With_To_Jordan', 'Moved_With_Whom',
        'Neighborhood_When_Moved', 'Health_Problems_List',
        'Languages_Spoken', 'Languages_Learned',
        'Workforce_Barriers_Women',
    ]

    space_num = re.compile(r'^-?\d+(\s+-?\d+)+$')
    single_num = re.compile(r'^-?\d+$')

    def _decode_codes_inplace(df, col, cmap):
        """Replace numeric codes with decoded labels (comma-separated for multi)."""
        if not cmap:
            return df
        def _decode(x):
            s = str(x).strip()
            if single_num.match(s):
                code = int(s)
                return cmap.get(code, s)
            if space_num.match(s):
                parts = s.split()
                labels = []
                for p in parts:
                    code = int(p) if p.lstrip('-').isdigit() else p
                    labels.append(str(cmap.get(code, p)))
                return ', '.join(labels)
            return x
        mask = df[col].notna()
        if mask.sum() > 0:
            df.loc[mask, col] = df.loc[mask, col].apply(_decode)
        return df

    for col in DECODE_COLS:
        if col not in df.columns:
            continue
        cmap = code_maps.get(col, {})
        df = _decode_codes_inplace(df, col, cmap)
        decoded = df[col].notna().sum()
        if cmap:
            print(f"    Decoded '{col}' in-place ({decoded} values, {len(cmap)} codes)")

    return df


###############################################################################
# AGGREGATE ROSTER / REPEATING COLUMNS (positions, businesses, etc.)
###############################################################################

def aggregate_roster_columns(df_raw, year):
    """For repeating question groups (position 1/2/3, business 1/2/3),
    compute aggregates: averages, totals, counts.
    Returns a dict of {new_col_name: Series} to add to the loaded dataframe.
    """
    cols = list(df_raw.columns)
    aggregated = {}

    # --- Employment Positions (position 1, 2, 3) ---
    pos_numeric_fields = {
        'Hours worked (last week)': 'Avg_Hours_Worked_Week',
        'Salary (last month, dinar)': 'Avg_Salary_Month_JOD',
        'Pre-tax income (year)': 'Total_Pretax_Income_Year_JOD',
    }
    for field_prefix, agg_name in pos_numeric_fields.items():
        pos_cols = []
        for p in range(1, 5):
            suffix = f', position {p}'
            matches = [c for c in cols if c.startswith(field_prefix) and suffix in c]
            if matches:
                pos_cols.append(matches[0])
        if len(pos_cols) >= 2:
            numeric_data = df_raw[pos_cols].apply(lambda s: mask_sentinels(s))
            if 'Total' in agg_name or 'Sum' in agg_name:
                aggregated[agg_name] = numeric_data.sum(axis=1, min_count=1)
            else:
                aggregated[agg_name] = numeric_data.mean(axis=1)
            print(f"    Aggregated {len(pos_cols)} position cols → {agg_name}")

    # Count of positions held
    occ_cols = [c for c in cols if c.startswith('Occupation, position')]
    if len(occ_cols) >= 2:
        aggregated['Num_Positions_Held'] = df_raw[occ_cols].notna().sum(axis=1)
        print(f"    Counted positions held → Num_Positions_Held")

    # --- Self-employment Businesses (Business No. 1, 2, 3) ---
    biz_numeric_fields = {
        'Business earnings last 30 days': 'Avg_Business_Earnings_30Days_JOD',
        'Business profit last 30 days': 'Avg_Business_Profit_30Days_JOD',
        'Hours worked (typical week)': 'Avg_Business_Hours_Week',
        'People employed (month)': 'Total_Business_Employees',
        'Business expenses (excluding rent) last 30 days': 'Avg_Business_Expenses_30Days_JOD',
    }
    for field_prefix, agg_name in biz_numeric_fields.items():
        biz_cols = []
        for b in range(1, 5):
            suffix = f', Business No. {b}'
            matches = [c for c in cols if c.startswith(field_prefix) and suffix in c]
            if matches:
                biz_cols.append(matches[0])
        if len(biz_cols) >= 2:
            numeric_data = df_raw[biz_cols].apply(lambda s: mask_sentinels(s))
            if 'Total' in agg_name:
                aggregated[agg_name] = numeric_data.sum(axis=1, min_count=1)
            else:
                aggregated[agg_name] = numeric_data.mean(axis=1)
            print(f"    Aggregated {len(biz_cols)} business cols → {agg_name}")

    # Count of businesses
    biz_ind_cols = [c for c in cols if c.startswith('Industry, Business No.')]
    if len(biz_ind_cols) >= 2:
        aggregated['Num_Businesses'] = df_raw[biz_ind_cols].notna().sum(axis=1)
        print(f"    Counted businesses → Num_Businesses")

    return aggregated


###############################################################################
# HANDLE "OTHER" COLUMNS
###############################################################################

def handle_other_columns(df):
    """Merge _Other columns into their base column in-place (no row duplication).

    If base value is 'Other' or similar, replace it with the specify text.
    If base already has a real value, append the Other text as comma-separated.
    """
    other_cols = [c for c in df.columns if '_Other' in c or '_other' in c]
    for oc in other_cols:
        base = oc.replace('_Other', '').replace('_other', '').rstrip('_')
        if base not in df.columns:
            continue
        has_val = df[oc].notna() & (df[oc].astype(str).str.strip() != '')
        if has_val.sum() == 0:
            df = df.drop(columns=[oc])
            continue
        for idx in df.index[has_val]:
            other_text = str(df.at[idx, oc]).strip()
            base_val = str(df.at[idx, base]).strip() if pd.notna(df.at[idx, base]) else ''
            if not base_val or base_val.lower() in ('other', 'other (specify)', 'nan', ''):
                df.at[idx, base] = other_text
            else:
                df.at[idx, base] = f"{base_val}, {other_text}"
        df = df.drop(columns=[oc])
        print(f"    Merged '{oc}' → '{base}' in-place ({has_val.sum()} values)")
    return df


###############################################################################
# LIGHT BINNING
###############################################################################

def add_bins(df):
    if 'Age' in df.columns:
        df['Age'] = mask_sentinels(df['Age'])
        df['Age_Group'] = pd.cut(df['Age'], bins=[0,17,25,35,50,65,200],
                                  labels=['0-17','18-25','26-35','36-50','51-65','65+'], right=True)
    if 'HH_Size' in df.columns:
        df['HH_Size'] = mask_sentinels(df['HH_Size'])
        df['HH_Size_Group'] = pd.cut(df['HH_Size'], bins=[-1,0,1,3,5,8,100],
                                      labels=['0','1','2-3','4-5','6-8','9+'], right=True)
    if 'Length_Schooling_Total_Years' in df.columns:
        df['Education_Level_Binned'] = pd.cut(
            mask_sentinels(df['Length_Schooling_Total_Years']),
            bins=[-1,0,6,9,12,100],
            labels=['None','1-6 years','7-9 years','10-12 years','12+ years'], right=True)
    return df


###############################################################################
# WRITE EXCEL WORKBOOK
###############################################################################

def write_workbook(df, topic_map, output_path):
    topics = ['Demographics', 'Household', 'Dwelling', 'Education', 'Economics',
              'Food_Security', 'Health', 'Mental_Health', 'Migration', 'Social',
              'Safety', 'Marriage_Fertility', 'Other']

    topic_map['Age_Group'] = 'Demographics'
    topic_map['HH_Size_Group'] = 'Household'
    topic_map['Length_Schooling_Total_Years'] = 'Education'
    topic_map['Education_Level_Binned'] = 'Education'

    all_data_cols = [c for c in df.columns if c not in ('Respondent_ID', 'Survey_Year')]
    for c in all_data_cols:
        if c not in topic_map:
            topic_map[c] = 'Other'

    # Use xlsxwriter for fast writing
    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        wb = writer.book
        hdr_fmt = wb.add_format({
            'bold': True, 'font_name': 'Arial', 'font_size': 11,
            'font_color': 'white', 'bg_color': '#2F5496',
            'text_wrap': True, 'align': 'center', 'valign': 'vcenter',
            'border': 1
        })

        for topic in topics:
            cols = ['Respondent_ID', 'Survey_Year'] + [c for c in all_data_cols if topic_map.get(c) == topic]
            existing = [c for c in cols if c in df.columns]
            if len(existing) <= 2:
                continue
            sheet_df = df[existing].copy()
            sheet_name = topic[:31]
            sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)

            ws = writer.sheets[sheet_name]
            for i, col_name in enumerate(existing):
                ws.write(0, i, col_name, hdr_fmt)
                ws.set_column(i, i, min(len(col_name) + 4, 40))
            ws.freeze_panes(1, 2)  # Freeze header row + ID/Year cols
            ws.autofilter(0, 0, sheet_df.shape[0], len(existing) - 1)
            print(f"  Sheet '{sheet_name}': {len(existing)-2} cols, {sheet_df.shape[0]} rows")

    print(f"\nSaved: {output_path}")


###############################################################################
# MAIN
###############################################################################

def main():
    print("=" * 70)
    print("COMBINED SURVEY DATA PROCESSOR (2020-2024)")
    print("=" * 70)

    configs = [
        (2020, f"{DATA_DIR}/2020.xlsx", m2020()),
        (2021, f"{DATA_DIR}/2021.xlsx", m2021()),
        (2022, f"{DATA_DIR}/2022.xlsx", m2022()),
        (2023, f"{DATA_DIR}/2023.xlsx", m2023()),
        (2024, f"{DATA_DIR}/2024.xlsx", m2024()),
    ]

    all_dfs = []
    all_topics = {}
    all_code_maps = {}

    for year, path, mapping in configs:
        print(f"\n--- {year} ---")
        df, tmap, cmaps = load_year(path, year, mapping)
        df = clean_encoding(df)
        df = combine_year_month(df)
        df = normalize_currency(df)
        df = standardize_timestamps(df)
        all_dfs.append((df, cmaps))
        all_topics.update(tmap)
        print(f"  Done: {df.shape}")

    # Decode multi-response codes per year (before combining, so code maps are year-specific)
    print(f"\n--- Decoding multi-response codes ---")
    decoded_dfs = []
    for df, cmaps in all_dfs:
        df = decode_multi_response(df, cmaps)
        decoded_dfs.append(df)

    print(f"\n--- Combining all years ---")
    combined = pd.concat(decoded_dfs, ignore_index=True, sort=False)
    print(f"  Combined: {combined.shape}")

    print(f"\n--- Adding bins ---")
    combined = add_bins(combined)

    print(f"\n--- Handling Other columns ---")
    combined = handle_other_columns(combined)
    print(f"  After Other merge: {combined.shape}")

    print(f"\n--- Writing Excel ---")
    import shutil
    tmp_output = "/tmp/Combined_survey_03_17.xlsx"
    output = f"{DATA_DIR}/Combined_survey_03_17.xlsx"
    write_workbook(combined, all_topics, tmp_output)
    shutil.copy2(tmp_output, output)
    print(f"Copied to: {output}")

    print(f"\n{'='*70}")
    print("SUMMARY")
    print(f"{'='*70}")
    print(f"Total rows: {combined.shape[0]}")
    print(f"Total columns: {combined.shape[1]}")
    for y in sorted(combined['Survey_Year'].unique()):
        n = combined[combined['Survey_Year']==y]['Respondent_ID'].nunique()
        print(f"  {y}: {n} unique respondents")


if __name__ == '__main__':
    main()

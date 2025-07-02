# =================================== IMPORTS ================================= #

import pandas as pd 
import plotly.graph_objects as go
import plotly.express as px
from datetime import datetime, timedelta
import os
import dash
from dash import dcc, html

# Google Web Credentials
import json
import base64
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# 'data/~$bmhc_data_2024_cleaned.xlsx'
# print('System Version:', sys.version)

# ------ Pandas Display Options ------ #
pd.set_option('display.max_rows', None)  # Show all rows
pd.set_option('display.max_columns', None)  # Show all columns (if needed)
pd.set_option('display.width', 1000)  # Adjust the width to prevent line wrapping
# -------------------------------------- DATA ------------------------------------------- #

current_dir = os.getcwd()
current_file = os.path.basename(__file__)
script_dir = os.path.dirname(os.path.abspath(__file__))
# data_path = 'data/Engagement_Responses.xlsx'
# file_path = os.path.join(script_dir, data_path)
# data = pd.read_excel(file_path)
# df = data.copy()

# Define the Google Sheets URL
sheet_url = "https://docs.google.com/spreadsheets/d/1D0oOioAfJyNCHhJhqFuhxxcx3GskP9L-CIL1DcOyhug/edit?resourcekey=&gid=1261604285#gid=1261604285"

# Define the scope
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

# Load credentials
encoded_key = os.getenv("GOOGLE_CREDENTIALS")

if encoded_key:
    json_key = json.loads(base64.b64decode(encoded_key).decode("utf-8"))
    creds = ServiceAccountCredentials.from_json_keyfile_dict(json_key, scope)
else:
    creds_path = r"C:\Users\CxLos\OneDrive\Documents\BMHC\Data\bmhc-timesheet-4808d1347240.json"
    if os.path.exists(creds_path):
        creds = ServiceAccountCredentials.from_json_keyfile_name(creds_path, scope)
    else:
        raise FileNotFoundError("Service account JSON file not found and GOOGLE_CREDENTIALS is not set.")

# Authorize and load the sheet
client = gspread.authorize(creds)
sheet = client.open_by_url(sheet_url)
worksheet = sheet.get_worksheet(0)  # ✅ This grabs the first worksheet
data = pd.DataFrame(worksheet.get_all_records())
# data = pd.DataFrame(client.open_by_url(sheet_url).get_all_records())
df = data.copy()

# Get the reporting month:
current_month = datetime(2025, 3, 1).strftime("%B")

# Trim leading and trailing whitespaces from column names
df.columns = df.columns.str.strip()

# Define a discrete color sequence
# color_sequence = px.colors.qualitative.Plotly

# Filtered df where 'Date of Activity:' is between Ocotber to December:
df['Date of Activity'] = pd.to_datetime(df['Date of Activity'], errors='coerce')
df = df[(df['Date of Activity'].dt.month >= 1) & (df['Date of Activity'].dt.month <= 3)]
df['Month'] = df['Date of Activity'].dt.month_name()

df_1 = df[df['Month'] == 'January']
df_2 = df[df['Month'] == 'February']
df_3 = df[df['Month'] == 'March']

# print(df.head(10))
# print('Column Names: \n', df.columns)
# print('DF Shape:', df.shape)
# print('Dtypes: \n', df.dtypes)
# print('Info:', df.info())
# print("Amount of duplicate rows:", df.duplicated().sum())

# print('Current Directory:', current_dir)
# print('Script Directory:', script_dir)
# print('Path to data:',file_path)

# ================================= Columns ================================= #

columns =[
    'Timestamp', 'Date of Activity', 'Person submitting this form:', 'Activity Duration (minutes):', 'Care Network Activity:', 'Entity name:', 'Brief Description:', 'Activity Status:', 'BMHC Administrative Activity:', 'Total travel time (minutes):', 'Community Outreach Activity:', 'Number engaged at Community Outreach Activity:', 'Any recent or planned changes to BMHC lead services or programs?', 'Email Address', 'Month'
]

# =============================== Missing Values ============================ #

# missing = df.isnull().sum()
# print('Columns with missing values before fillna: \n', missing[missing > 0])

# ============================== Data Preprocessing ========================== #

# Check for duplicate columns
# duplicate_columns = df.columns[df.columns.duplicated()].tolist()
# print(f"Duplicate columns found: {duplicate_columns}")
# if duplicate_columns:
#     print(f"Duplicate columns found: {duplicate_columns}")

df.rename(
    columns={
        "Activity Duration (minutes):": "Minutes",
        "Person submitting this form:": "Person",
        "Total travel time (minutes):": "Travel Time",
        "BMHC Administrative Activity:": "Admin Activity",
        "Care Network Activity:": "Care Activity",
        "Community Outreach Activity:": "Outreach Activity",
        "Activity Status:": "Activity Status",
    }, 
inplace=True)

# Get the reporting quarter:
def get_custom_quarter(date_obj):
    month = date_obj.month
    if month in [10, 11, 12]:
        return "Q1"  # October–December
    elif month in [1, 2, 3]:
        return "Q2"  # January–March
    elif month in [4, 5, 6]:
        return "Q3"  # April–June
    elif month in [7, 8, 9]:
        return "Q4"  # July–September

# Reporting Quarter (use last month of the quarter)
report_date = datetime(2025, 3, 1)  # Example report date for Q2 (Jan–Mar)
month = report_date.month
report_year = report_date.year
current_quarter = get_custom_quarter(report_date)
# print(f"Reporting Quarter: {current_quarter}")

# Adjust the quarter calculation for custom quarters
if month in [10, 11, 12]:
    quarter = 1  # Q1: October–December
elif month in [1, 2, 3]:
    quarter = 2  # Q2: January–March
elif month in [4, 5, 6]:
    quarter = 3  # Q3: April–June
elif month in [7, 8, 9]:
    quarter = 4  # Q4: July–September

# Define a mapping for months to their corresponding quarter
quarter_months = {
    1: ['October', 'November', 'December'],  # Q1
    2: ['January', 'February', 'March'],    # Q2
    3: ['April', 'May', 'June'],            # Q3
    4: ['July', 'August', 'September']      # Q4
}

# Get the months for the current quarter
months_in_quarter = quarter_months[quarter]

# Calculate start and end month indices for the quarter
# all_months = [
#     'January', 'February', 'March', 
#     'April', 'May', 'June',
#     'July', 'August', 'September', 
#     'October', 'November', 'December'
# ]
# start_month_idx = (quarter - 1) * 3
# month_order = all_months[start_month_idx:start_month_idx + 3]

# ------------------------ Total Engagements DF ---------------------------- #

total_engagements = len(df)
# print('Total Engagements:', total_engagements)

# ------------------------ Engagement Hours DF ---------------------------- #

# print("Activity Duration Unique Before: \n", df['Activity Duration (minutes):'].unique().tolist())
# print(df['Activity Duration (minutes):'])

activity_unique = [120, 
                   300, 
                   180, 
                   60,
                   0,
                   5,
                   30, 
                   15, 
                   10, 
                   90, 
                   45, 
                   '2400 minutes', 
                   '1680 minutes( 28 hours) over 2 week period', 
                   240, 
                   1320, 
                   360,
                   '450 mins', 
                   720, 
                   105, 
                   420, 
                   210, 
                   150, 
                   280, 
                   'Onboarding Activities (Jordan Calbert)', 
                   '75 minutes',
                   20, 
                   16, 
                   75]

df['Minutes'] = (
    df['Minutes']
    .astype(str)
    .str.strip()               
    .replace({
        "6 hrs": 360,
        "5 hrs": 300,
        "nan": 0,
        "2400 minutes": 2400,
        "1680 minutes( 28 hours) over 2 week period": 1680,
        "450 mins": 450,
        "75 minutes": 75,
        "Onboarding Activities (Jordan Calbert)": 0,
    })
)

df['Minutes'] = pd.to_numeric(df['Minutes'], errors='coerce')
df['Minutes'] = df['Minutes'].fillna(0)

# print("Activity Duration Unique After: \n", df['Activity Duration (minutes):'].unique().tolist())

# Calculate total hours for each month in the current quarter
hours = []
for month in months_in_quarter:
    hours_in_month = df[df['Month'] == month]['Minutes'].sum()/60
    hours_in_month = round(hours_in_month)
    hours.append(hours_in_month)
    # print(f'Engagement hours in {month}:', hours_in_month, 'hours')
    
eng_hours = df.groupby('Minutes').size().reset_index(name='Count')
eng_hours = df['Minutes'].sum()/60
eng_hours = round(eng_hours)

df_hours = pd.DataFrame({
    'Month': months_in_quarter,
    'Hours': hours
})

# Engagment Hours Bar Chart:
hours_fig = px.bar(
    df_hours,
    x='Month',
    y='Hours',
    color = 'Month',
    text='Hours',
    labels={
        'Hours': 'Hours',
        'Month': 'Month'
    }
).update_layout(
    title_x=0.5,
    xaxis_title='Month',
    yaxis_title='Engagement Hours',
    height=600,  # Adjust graph height
    title=dict(
        text= f'{current_quarter} Engagement Hours by Month',
        x=0.5, 
        font=dict(
            size=35,
            family='Calibri',
            color='black',
            )
    ),
    font=dict(
        family='Calibri',
        size=17,
        color='black'
    ),
    xaxis=dict(
        title=dict(
            text=None,
            # text="Month",
            font=dict(size=20),  # Font size for the title
        ),
        tickmode='array',
        tickvals=df_hours['Month'].unique(),
        tickangle=0  # Rotate x-axis labels for better readability
    ),
).update_traces(
    texttemplate='%{text}',  # Display the count value above bars
    textfont=dict(size=20),  # Increase text size in each bar
    textposition='auto',  # Automatically position text above bars
    textangle=0, # Ensure text labels are horizontal
    hovertemplate=(  # Custom hover template
        '<b>Name</b>: %{label}<br><b>Count</b>: %{y}<extra></extra>'  
    ),
)

hours_pie = px.pie(
    df_hours,
    names='Month',
    values='Hours',
    color='Month',
    height=550
).update_layout(
    title=dict(
        x=0.5,
        text=f'{current_quarter} Ratio Engagement Hours by Month',  # Title text
        font=dict(
            size=35,  # Increase this value to make the title bigger
            family='Calibri',  # Optional: specify font family
            color='black'  # Optional: specify font color
        ),
    ),  # Center-align the title
    margin=dict(
        l=0,  # Left margin
        r=0,  # Right margin
        t=100,  # Top margin
        b=0   # Bottom margin
    )  # Add margins around the chart
).update_traces(
    rotation=180,  # Rotate pie chart 90 degrees counterclockwise
    textfont=dict(size=19),  # Increase text size in each bar
    textinfo='value+percent',
    # texttemplate='<br>%{percent:.0%}',  # Format percentage as whole numbers
    hovertemplate='<b>%{label}</b>: %{value}<extra></extra>'
)

# ------------------------ Total Travel Time DF ---------------------------- #

# print("Travel Time Unique Before: \n", df['Travel Time'].unique().tolist())

travel_unique =  [
    0, 
    45,
    60,
    30,
    300,
    15,
    90,
    'End of Week 1 to 1 Performance Review',
    240,
    'nan',
    'Sustainable Food Center + APH Health Education Strategy Meeting & Planning Activities',
    480,
    120,
    'Community First Village Huddle',
 ]

# Clean travel time values
df['Travel Time'] = (
    df['Travel Time']
    .astype(str)
    .str.strip()
    .replace({
        "End of Week 1 to 1 Performance Review": 0,
        "Sustainable Food Center + APH Health Education Strategy Meeting & Planning Activities": 0,
        "Community First Village Huddle": 0,
        "nan": 0,
    })
)

df['Travel Time'] = pd.to_numeric(df['Travel Time'], errors='coerce')
df['Travel Time'] = df['Travel Time'].fillna(0)

# print("Travel Time Unique After: \n", df['Total travel time (minutes):'].unique().tolist())
# print(['Travel Time Value Counts: \n', df['Travel Time'].value_counts()])

total_travel_time = df['Travel Time'].sum()
total_travel_time = round(total_travel_time)
# print("Total travel time:",total_travel_time)

# Calculate total travel time per month
travel_hours = []
for month in months_in_quarter:
    hours_in_month = df[df['Month'] == month]['Travel Time'].sum() / 60
    hours_in_month = round(hours_in_month)
    travel_hours.append(hours_in_month)

df_travel = pd.DataFrame({
    'Month': months_in_quarter,
    'Travel Time': travel_hours
})

# Bar chart
travel_fig = px.bar(
    df_travel,
    x='Month',
    y='Travel Time',
    color='Month',
    text='Travel Time',
    labels={
        'Travel Time': 'Travel Time (hours)',
        'Month': 'Month'
    }
).update_layout(
    title_x=0.5,
    xaxis_title='Month',
    yaxis_title='Travel Time (hours)',
    height=600,
    title=dict(
        text=f'{current_quarter} Travel Time by Month',
        x=0.5, 
        font=dict(
            size=35,
            family='Calibri',
            color='black',
            )
    ),
    font=dict(
        family='Calibri',
        size=17,
        color='black'
    ),
    xaxis=dict(
        title=dict(
            text=None,
            font=dict(size=20),
        ),
        tickmode='array',
        tickvals=df_travel['Month'].unique(),
        tickangle=0
    ),
).update_traces(
    texttemplate='%{text}',
    textfont=dict(size=20),
    textposition='auto',
    textangle=0,
    hovertemplate='<b>Month</b>: %{label}<br><b>Travel Time</b>: %{y} hours<extra></extra>',
)

# Pie chart
travel_pie = px.pie(
    df_travel,
    names='Month',
    values='Travel Time',
    color='Month',
    height=550
).update_layout(
    title=dict(
        x=0.5,
        text=f'{current_quarter} Travel Time Ratio by Month',
        font=dict(
            size=35,
            family='Calibri',
            color='black'
        ),
    ),
    margin=dict(l=0, r=0, t=100, b=0)
).update_traces(
    rotation=180,
    textfont=dict(size=19),
    textinfo='value+percent',
    hovertemplate='<b>%{label}</b>: %{value} hours<extra></extra>'
)

# --------------------------------- Activity Status DF -------------------------------- #

# Group by 'Activity Status:' dataframe
df_activity_status = df.groupby('Activity Status').size().reset_index(name='Count')

# print("Activity Status Unique before: \n", df['Activity Status'].unique().tolist())

mode = df['Activity Status'].mode()[0]
df['Activity Status'] = df['Activity Status'].fillna(mode)

df_activity_status['Activity Status'] = (
    df_activity_status['Activity Status']
    .str.strip()
    .replace({
        '': mode,
    })
)

df_activity_status_counts = (
    df.groupby(['Month', 'Activity Status'], sort=False)
    .size()
    .reset_index(name='Count')
)

df_activity_status_counts['Month'] = pd.Categorical(
    df_activity_status_counts['Month'],
    categories=months_in_quarter,
    ordered=True
)

# print("Activity Status Unique after: \n", df['Activity Status'].unique().tolist())

status_fig = px.bar(
    df_activity_status_counts,
    x='Month',
    y='Count',
    color='Activity Status',
    barmode='group',
    text='Count',
    labels={
        'Count': 'Number of Submissions',
        'Month': 'Month',
        'Activity Status': 'Activity Status'
    }
).update_layout(
    title_x=0.5,
    xaxis_title='Month',
    yaxis_title='Count',
    height=550,
    font=dict(
        family='Calibri',
        size=17,
        color='black'
    ),
    title=dict(
        text= f'{current_quarter} Activity Status by Month',
        x=0.5, 
        font=dict(
            size=22,
            family='Calibri',
            color='black',
            )
    ),
    xaxis=dict(
        tickmode='array',
        tickvals=df_activity_status_counts['Month'].unique(),
        tickangle=-35
    ),
    legend=dict(
        title='',
        orientation="v",
        x=1.05,
        xanchor="left",
        y=1,
        yanchor="top"
    ),
    hovermode='x unified'
).update_traces(
    textfont=dict(size=17),  # Increase text size in each bar
    textposition='outside',
    hovertemplate='<br><b>Count: </b>%{y}<br>',
    customdata=df_activity_status_counts['Activity Status'].values.tolist()
)

status_pie = px.pie(
    df_activity_status,
    names='Activity Status',
    values='Count',
).update_layout(
    title= f'{current_quarter} Activity Status',
    title_x=0.5,
    height=550,
    font=dict(
        family='Calibri',
        size=17,
        color='black'
    )
).update_traces(
    textposition='auto',
    # textinfo='label+percent',
    texttemplate='%{value}<br>%{percent:.0%}',  # Format percentage as whole numbers
    hovertemplate='<b>Status</b>: %{label}<br><b>Count</b>: %{value}<extra></extra>'
)

# --------------------- Administrative Activity DF ------------------------ # 

admin_value_counts = df['Admin Activity'].value_counts()

# Convert the Series to a DataFrame
admin_value_counts_df = admin_value_counts.reset_index()
admin_value_counts_df.columns = ['Admin Activity', 'Count']  # Rename columns

# Save the DataFrame to an Excel file
output_path = os.path.join(script_dir, 'admin_activity_counts.xlsx')

# admin_value_counts_df.to_excel(output_path, index=False)
# print(f"Admin activity counts saved to {output_path}")

# print("Administrative Activity Unique Before:", admin_value_counts)
# print("Admin Activity value counts:", df['Admin Activity'].value_counts())

admin_unique = [
    '', 'Communication & Correspondence', '(4) Outreach 1 to 1 Strategy Meetings', 'Outreach Team Meeting', "St. David's + Kazi 88.7FM Strategic Partnership Meeting & Strategy Planning Discussion/Activities", 'Travis County Judge Andy Brown & Travis County Commissioner Ann Howard BMHC Tour & Discussion', 'Key Leaders Huddle', '2025 Calendar Year Outreach Preparation & Strategic Planning Activities', 'BMHC Quarterly Team Meeting', 'Events Planning Meeting', 'Gudlife 2025 Strategic Planning Session', 'Community First Village Huddle', 'Community First Village Onsite Outreach', 'Record Keeping & Documentation', "Men's Mental Health 1st Saturdays", 'Financial & Budgetary Management', 'Office Management', 'Meeting With Frost Bank', 'HR Support', 'Compliance & Policy Enforcement', 'BMHC Team', 'Special Events Team Meeting', 'Weekly team meeting', 'National Kidney Foundation Strategy Meeting (Know Your Numbers Campaign Program)', 'Healthy Cuts/Know Your Numbers Event at Community First Village', 'IT', 'Meeting with Cameron', 'Implementation Studios Planning & Strategy Meeting', 'Outreach & Navigation Leads 1 to 1 Strategy Meeting', 'BMHC + Community First Village Onsite Outreach Strategy Planning Huddle', 'BMHC + Gudlife Strategy Huddle', 'BMHC + Community First Village Onsite Outreach Strategy Huddle', 'Downtown Austin Community Court Onsite Outreach', 'Outreach Onboarding (Jordan Calbert)', 'BMHC + Gudlife Outreach Strategy Huddle', 'End of Week 1 to 1 Performance Review', 'BMHC + KAZI Basketball Tournament', 'BMHC Gudlife Meeting', 'BMHC Pflugerville Asset Mapping Activities', '100 Black Men of Austin Quarterly Partnership Review (QPR)', 'Onboarding', 'Outreach 1 to 1 Strategy Meetings', 'Impact Forms Follow Up Meeting', 'Community First Village Outreach Strategy Huddle', 'Any Baby Can Tour & Partnership Meeting', 'Housing Authority of Travis County (Self-Care Day) Outreach Event', 'psh support call with Dr Wallace', 'BMHC Tour (Austin Mayor Kirk Watson & Austin City Council Member District 4 "Chito" Vela)', 'PSH Audit for ECHO', 'BMHC + Community First Village Neighborhood Care Team Planning Meeting', 'Biweekly PSH staffing with ECHO', 'PSH file updates and case staffing', 'Child Inc Travis County HeadStart Program (Fatherhood Program Event)', 'BMHC + Breakthrough of Central Texas Partnership Discussion', 'Housing Authority of Travis County Quarterly Partnership Review (QPR)', 'PSH', 'Meeting', 'Training', 'BMHC & GUD LFE Huddle Meeting', 'BMHC Internal & External Emails and Phone Calls Performed', 'Manor 5K Planning Meeting & Follow Up Activities', 'HSO stakeholder meeting', 'outreach coordination meeting', 'Outreach & Navigation Team Leads Huddle', 'Implementation Studios Planning Meeting', 'homeless advocacy meeting', 'Central Health Virtual Lunch', 'Community First Village Onsite Outreach & Healthy Cuts Preventative Screenings', 'MOU conversation with Extended Stay America', 'PSH iPilot', 'End of Week Outreach Performance Reviews', 'Outreach Onboarding Activities (Jordan Calbert)', 'BMHC Gudlife Huddle', 'BMHC & GUD LIFE Weekly Huddle', 'Bi-Partner Neighbor Partner Engagement Meeting', 'BOLO list and placement', 'In-Person Key Leaders Huddle', 'weekly HMIS updates and phone calls for clients on BOLO list', 'HMIS monthly reports submission to ECHO', 'timesheet completion and submit to Dr. Wallace', 'client referrals/community partnership'
]

admin_categories = [
    '1 to 1 Outreach Strategy Meetings',
    'BMHC & GUD LIFE Huddle Meetings',
    'Administrative & Communications',
    'Research & Planning',
    'Reports & Documentation',
    'Financial & Budgeting',
    'Human Resources (HR) & Office Management',
    'Training & Onboarding',
    'PSH & Client Support',
    'Outreach & Engagement',
    'Stakeholder & Key Leader Meetings',
    'Performance & Reviews'
]

# print("Administrative Activity Unique Before:", df['Admin Activity'].unique().tolist())

df['Admin Activity'] = (
    df['Admin Activity']
        .astype(str)
        .str.strip()
        .replace({
        
        "" : pd.NA,
        
        # 1 to 1 Outreach Strategy Meetings
        '(4) Outreach 1 to 1 Strategy Meetings': '1 to 1 Outreach Strategy Meetings',
        'Outreach Team Meeting': '1 to 1 Outreach Strategy Meetings',
        'Outreach & Navigation Leads 1 to 1 Strategy Meeting': '1 to 1 Outreach Strategy Meetings',
        'Outreach 1 to 1 Strategy Meetings': '1 to 1 Outreach Strategy Meetings',

        # BMHC & GUD LIFE Huddle Meetings
        'BMHC & GUD LFE Huddle Meeting': 'BMHC & GUD LIFE Huddle Meetings',
        'BMHC Gudlife Huddle': 'BMHC & GUD LIFE Huddle Meetings',
        'BMHC & GUD LIFE Weekly Huddle': 'BMHC & GUD LIFE Huddle Meetings',
        'Key Leaders Huddle': 'BMHC & GUD LIFE Huddle Meetings',
        'BMHC + Gudlife Strategy Huddle': 'BMHC & GUD LIFE Huddle Meetings',
        'BMHC + Gudlife Outreach Strategy Huddle': 'BMHC & GUD LIFE Huddle Meetings',

        # Administrative & Communications
        'Communication & Correspondence': 'Administrative & Communications',
        'BMHC Quarterly Team Meeting': 'Administrative & Communications',
        'BMHC Team': 'Administrative & Communications',
        'Weekly team meeting': 'Administrative & Communications',
        'IT': 'Administrative & Communications',
        'BMHC Internal & External Emails and Phone Calls Performed': 'Administrative & Communications',
        'Meeting With Frost Bank': 'Administrative & Communications',
        'Outreach Onboarding Activities (Jordan Calbert)': 'Administrative & Communications',

        # Research & Planning
        '2025 Calendar Year Outreach Preparation & Strategic Planning Activities': 'Research & Planning',
        'Gudlife 2025 Strategic Planning Session': 'Research & Planning',
        'Events Planning Meeting': 'Research & Planning',
        'Implementation Studios Planning & Strategy Meeting': 'Research & Planning',
        'Impact Forms Follow Up Meeting': 'Research & Planning',
        'MOU conversation with Extended Stay America': 'Research & Planning',
        'Implementation Studios Planning Meeting': 'Research & Planning',
        'BMHC Pflugerville Asset Mapping Activities': 'Research & Planning',
        'Housing Authority of Travis County Quarterly Partnership Review (QPR)': 'Research & Planning',

        # Reports & Documentation
        'Record Keeping & Documentation': 'Reports & Documentation',
        'HMIS monthly reports submission to ECHO': 'Reports & Documentation',
        'weekly HMIS updates and phone calls for clients on BOLO list': 'Reports & Documentation',

        # Financial & Budgeting
        'Financial & Budgetary Management': 'Financial & Budgeting',

        # Human Resources (HR) & Office Management
        'Office Management': 'Human Resources (HR) & Office Management',
        'HR Support': 'Human Resources (HR) & Office Management',
        'Compliance & Policy Enforcement': 'Human Resources (HR) & Office Management',
        'timesheet completion and submit to Dr. Wallace': 'Human Resources (HR) & Office Management',

        # Training & Onboarding
        'Onboarding': 'Training & Onboarding',
        'Outreach Onboarding (Jordan Calbert)': 'Training & Onboarding',
        'Training': 'Training & Onboarding',

        # PSH & Client Support
        'psh support call with Dr Wallace': 'PSH & Client Support',
        'PSH Audit for ECHO': 'PSH & Client Support',
        'PSH': 'PSH & Client Support',
        'PSH iPilot': 'PSH & Client Support',
        'Biweekly PSH staffing with ECHO': 'PSH & Client Support',
        'PSH file updates and case staffing': 'PSH & Client Support',
        'client referrals/community partnership': 'PSH & Client Support',
        'BMHC + Community First Village Neighborhood Care Team Planning Meeting': 'PSH & Client Support',

        # Outreach & Engagement
        'Community First Village Onsite Outreach': 'Outreach & Engagement',
        'Healthy Cuts/Know Your Numbers Event at Community First Village': 'Outreach & Engagement',
        'Community First Village Huddle': 'Outreach & Engagement',
        'Outreach & Navigation Team Leads Huddle': 'Outreach & Engagement',
        'Downtown Austin Community Court Onsite Outreach': 'Outreach & Engagement',
        'BMHC + Community First Village Onsite Outreach Strategy Planning Huddle': 'Outreach & Engagement',
        'BMHC + Community First Village Onsite Outreach Strategy Huddle': 'Outreach & Engagement',
        'Outreach & Navigation Leads 1 to 1 Strategy Meeting': 'Outreach & Engagement',
        'Community First Village Outreach Strategy Huddle': 'Outreach & Engagement',
        'Outreach & Engagement': 'Outreach & Engagement',
        'Outreach Team Meeting': 'Outreach & Engagement',
        'Any Baby Can Tour & Partnership Meeting': 'Outreach & Engagement',
        'Housing Authority of Travis County (Self-Care Day) Outreach Event': 'Outreach & Engagement',
        'Outreach Onboarding Activities (Jordan Calbert)': 'Outreach & Engagement',
        'Outreach Onboarding (Jordan Calbert)': 'Outreach & Engagement',
        'homeless advocacy meeting': 'Outreach & Engagement',
        'Community First Village Onsite Outreach & Healthy Cuts Preventative Screenings': 'Outreach & Engagement',
        'BOLO list and placement': 'Outreach & Engagement',

        # Stakeholder & Key Leader Meetings
        'St. David\'s + Kazi 88.7FM Strategic Partnership Meeting & Strategy Planning Discussion/Activities': 'Stakeholder & Key Leader Meetings',
        'Travis County Judge Andy Brown & Travis County Commissioner Ann Howard BMHC Tour & Discussion': 'Stakeholder & Key Leader Meetings',
        'Key Leaders Huddle': 'Stakeholder & Key Leader Meetings',
        'BMHC Gudlife Meeting': 'Stakeholder & Key Leader Meetings',
        '100 Black Men of Austin Quarterly Partnership Review (QPR)': 'Stakeholder & Key Leader Meetings',
        'National Kidney Foundation Strategy Meeting (Know Your Numbers Campaign Program)': 'Stakeholder & Key Leader Meetings',
        'Meeting with Cameron': 'Stakeholder & Key Leader Meetings',
        'BMHC + Gudlife Strategy Huddle': 'Stakeholder & Key Leader Meetings',
        'BMHC Gudlife Huddle': 'Stakeholder & Key Leader Meetings',
        'BMHC & Gudlife Strategy Huddle': 'Stakeholder & Key Leader Meetings',
        'BMHC + Breakthrough of Central Texas Partnership Discussion': 'Stakeholder & Key Leader Meetings',
        'Housing Authority of Travis County Quarterly Partnership Review (QPR)': 'Stakeholder & Key Leader Meetings',
        'Bi-Partner Neighbor Partner Engagement Meeting': 'Stakeholder & Key Leader Meetings',
        'In-Person Key Leaders Huddle': 'Stakeholder & Key Leader Meetings',
        'Any Baby Can Tour & Partnership Meeting': 'Stakeholder & Key Leader Meetings',
        'PSH Audit for ECHO': 'Stakeholder & Key Leader Meetings',
        'Meeting': 'Stakeholder & Key Leader Meetings',

        # Performance & Reviews
        'End of Week 1 to 1 Performance Review': 'Performance & Reviews',
        'End of Week Outreach Performance Reviews': 'Performance & Reviews',

        # Special Event Support
        "Men's Mental Health 1st Saturdays": 'Special Event Support',
        'Special Events Team Meeting': 'Special Event Support',
        'BMHC + KAZI Basketball Tournament': 'Special Event Support',
        'BMHC Tour (Austin Mayor Kirk Watson & Austin City Council Member District 4 "Chito" Vela)': 'Special Event Support',
        'Child Inc Travis County HeadStart Program (Fatherhood Program Event)': 'Special Event Support',
        'Manor 5K Planning Meeting & Follow Up Activities': 'Special Event Support',
        'Special Event Support': 'Special Event Support',

        # Outreach & Engagement
        'HSO stakeholder meeting': 'Outreach & Engagement',
        'outreach coordination meeting': 'Outreach & Engagement',
        'Central Health Virtual Lunch': 'Stakeholder & Key Leader Meetings'
    })
)

df_admin = df[df['Admin Activity'].notna()]

# admin_mode = df_admin['Admin Activity'].mode()[0]
# print("Admin Mode:", admin_mode)
# df['Admin Activity'] = df['Admin Activity'].fillna(admin_mode)

# Check the changes
# print("Administrative Activity Unique After Replacement:", df['Admin Activity'].unique().tolist())
# print("Admin value counts:", df_admin['Admin Activity'].value_counts())

# Find any remaining unmatched purposes
unmatched_admin = df_admin[~df_admin['Admin Activity'].isin(admin_categories)]['Admin Activity'].unique().tolist()
# print("Unmatched Administrative Activities:", unmatched_admin)

# Group the data by 'Month' and 'Admin Activity' and count occurrences
df_admin_counts = (
    df_admin.groupby(['Month', 'Admin Activity'], sort=True)
    .size()
    .reset_index(name='Count')
)

# Assign categorical ordering to the 'Month' column
df_admin_counts['Month'] = pd.Categorical(
    df_admin_counts['Month'],
    categories=months_in_quarter,
    ordered=True
)

# Sort df:
df_admin_counts = df_admin_counts.sort_values(by=['Month', 'Admin Activity'])

# Create the grouped bar chart
admin_fig = px.bar(
    df_admin_counts,
    x='Month',
    y='Count',
    color='Admin Activity',
    barmode='group',
    text='Count',
    labels={
        'Count': 'Number of Activities',
        'Month': 'Month',
        'Admin Activity': 'Administrative Activity'
    }
).update_layout(
    title_x=0.5,
    xaxis_title='Month',
    yaxis_title='Count',
    height=900,  # Adjust graph height
    title=dict(
        text= f'{current_quarter } Administrative Activities by Month',
        x=0.5, 
        font=dict(
            size=35,
            family='Calibri',
            color='black',
            )
    ),
    font=dict(
        family='Calibri',
        size=17,
        color='black'
    ),
    xaxis=dict(
        tickmode='array',
        tickvals=df_admin_counts['Month'].unique(),
        tickangle=-35  # Rotate x-axis labels for better readability
    ),
    legend=dict(
        # title='Administrative Activity',
        title=None,
        orientation="v",  # Vertical legend
        x=1.05,  # Position legend to the right
        xanchor="left",  # Anchor legend to the left
        y=1,  # Position legend at the top
        yanchor="top"  # Anchor legend at the top
    ),
    margin=dict(l=0, r=0, t=0, b=0),
    hovermode='x unified'  # Display unified hover info
).update_traces(
    textposition='outside',  # Display text above bars
    textfont=dict(size=30),  # Increase text size in each bar
    hovertemplate=(
        '<br>'
        '<b>Count: </b>%{y}<br>'  # Count
    ),
    customdata=df_admin_counts['Admin Activity'].values.tolist()
)

df_admin = df_admin.groupby('Admin Activity').size().reset_index(name='Count')

# Create the pie chart for Administrative Activity distribution
admin_pie = px.pie(
    df_admin,
    names='Admin Activity',
    values='Count',
    color='Admin Activity',
    height=800,
    title= f'{current_quarter} Distribution of Administrative Activities'
).update_layout(
    title=dict(
        x=0.5,
        text= f'{current_quarter} Distribution of Administrative Activities',  # Title text
        font=dict(
            size=35,  # Increase this value to make the title bigger
            family='Calibri',  # Optional: specify font family
            color='black'  # Optional: specify font color
        ),
    ),  
    margin=dict(
        t=150,  # Adjust the top margin (increase to add more padding)
        l=20,   # Optional: left margin
        r=20,   # Optional: right margin
        b=20    # Optional: bottom margin
    )
).update_traces(
    rotation=140,  # Rotate pie chart 90 degrees counterclockwise
    textfont=dict(size=19),  # Increase text size
    textinfo='value+percent',
    # texttemplate='<br>%{percent:.1%}',  # Format percentage as whole numbers
    hovertemplate='<b>%{label}</b>: %{value}<extra></extra>'  # Hover details
)

# --------------------- Care Network Activity DF ------------------------ #

# Value counts for 'Care Activity'
care_value_counts = df['Care Activity'].value_counts()

# Convert the Series to a DataFrame
care_value_counts_df = care_value_counts.reset_index() 
care_value_counts_df.columns = ['Care Activity', 'Count']  # Rename columns

# Save the DataFrame to an Excel file
care_output_path = os.path.join(script_dir, 'care_activity_counts.xlsx')

care_value_counts_df.to_excel(care_output_path, index=False)
print(f"Care activity counts saved to {care_output_path}")

# print(df['Care Network Activity'].unique().tolist())
# print("Care Network Activity Value Counts:", care_value_counts)

custom_colors = {
    'January': 'Blues',
    'February': 'Greens',
    'March': 'Oranges',
    'April': 'Purples',
    'May': 'Reds',
    'June': 'Greys',
    'July': 'YlGn',
    'August': 'YlOrBr',
    'September': 'PuRd',
    'October': 'BuPu',
    'November': 'GnBu',
    'December': 'YlGnBu',
# The code snippet provided is a Python dictionary with a key-value pair. The key is 'Count'
# and the value is 'Number of Submissions'. This dictionary is used to store information
# related to the count of submissions.
}

care_unique = [
'Clinical Provider', '', 'Government', 'BMHC Team', 'SDoH Provider', 'Outreach & Navigation', 'Religious', 'Movement is Medicine', "Men's Mental Health 1st Saturdays at BMHC (Man In Man)", 'Give Back Program', 'Movement is Medicine ', 'Academic', 'Movement is medicine', 'Work Force Development', 'Community Partnership in media', 'BMHC - Austin', 'Policy Documentation Reviewed, Signed & Sent', 'BMHC - Pflugerville Navigation Meeting', 'Care Network Prospect', 'Pink Bus Program', 'Community partnership for health and wellness', 'Health Resource', 'BMHC + Sustainable Food Center Follow Up Meeting', 'ECHO Pilot Program', 'Administrative Support', 'Outreach Onboarding (Jordan Calbert)', 'Community Partner', 'Black Nurses Association Community Partner', 'ECHO PSH Pilot Program ', 'KAZI 88.7 FM (Marketing & Exposure)', 'Community First Village Onsite Outreach', 'Discussed coordination and referral services for D. Bell', 'Community ', 'University of Texas at Austin', 'PSH CASEWORKER UPDATES AND CALLS', 'PSH HMIS updates and caseworker notes', 'Community Fitness Gym', 'Caseworker calls for PSH', 'PSH caseworker and BMHC updates', 'Outreach Team Meeting', 'Agency Partnership/Collaboration ', 'Kensington Integral Care housing ', 'community partnerships/engagement', 'Referals'
]

care_categories = [
    'Clinical Providers & Government',
    'BMHC & Team Activities',
    'Outreach & Navigation',
    'Health & Wellness Programs',
    'Academic & Workforce Development',
    'Community Partnerships & Media',
    'PSH & Case Management',
    'Administrative & Support Services',
    'Special Programs & Initiatives',
    'Partner & Stakeholder Engagement',
    'Administrative & Communications',
    'Outreach & Engagement',
    'Stakeholder & Key Leader Meetings',
    'Research & Planning',
    'Reports & Documentation',
    'Special Event Support',
    'Financial & Budgeting',
    'Human Resources & Office Management',
    'BMHC & GUD LIFE Huddle Meetings',
    'Performance & Reviews',
    'Training & Onboarding',
    'PSH & Client Support',
    '1 to 1 Outreach Strategy Meetings',
]

# print("Care Network Activity Unique Before:", df['Care Activity'].unique().tolist())

df['Care Activity'] = (
    df['Admin Activity']
    .astype(str)
    .str.strip()
    .replace({
        
        "" : pd.NA,
        "<NA>": pd.NA,
        
        'Clinical Provider': 'Clinical Provider',
        'Government': 'Government',
        'BMHC Team': 'BMHC Team',
        'SDoH Provider': 'SDoH Provider',
        'Outreach & Navigation': 'Outreach & Navigation',
        'Religious': 'Religious',
        'Movement is Medicine': 'Movement is Medicine',
        "Men's Mental Health 1st Saturdays at BMHC (Man In Man)": "Men's Mental Health 1st Saturdays at BMHC (Man In Man)",
        'Give Back Program': 'Give Back Program',
        'Movement is Medicine ': 'Movement is Medicine',
        'Academic': 'Academic',
        'Movement is medicine': 'Movement is Medicine',
        'Work Force Development': 'Work Force Development',
        'Community Partnership in media': 'Community Partnership in Media',
        'BMHC - Austin': 'BMHC - Austin',
        'Policy Documentation Reviewed, Signed & Sent': 'Policy Documentation Reviewed, Signed & Sent',
        'BMHC - Pflugerville Navigation Meeting': 'BMHC - Pflugerville Navigation Meeting',
        'Care Network Prospect': 'Care Network Prospect',
        'Pink Bus Program': 'Pink Bus Program',
        'Community partnership for health and wellness': 'Community Partnership for Health and Wellness',
        'Health Resource': 'Health Resource',
        'BMHC + Sustainable Food Center Follow Up Meeting': 'BMHC + Sustainable Food Center Follow Up Meeting',
        'ECHO Pilot Program': 'ECHO Pilot Program',
        'Administrative Support': 'Administrative Support',
        'Outreach Onboarding (Jordan Calbert)': 'Outreach Onboarding (Jordan Calbert)',
        'Community Partner': 'Community Partner',
        'Black Nurses Association Community Partner': 'Black Nurses Association Community Partner',
        'ECHO PSH Pilot Program ': 'ECHO PSH Pilot Program',
        'KAZI 88.7 FM (Marketing & Exposure)': 'KAZI 88.7 FM (Marketing & Exposure)',
        'Community First Village Onsite Outreach': 'Community First Village Onsite Outreach',
        'Discussed coordination and referral services for D. Bell': 'Discussed Coordination and Referral Services for D. Bell',
        'Community ': 'Community',
        'University of Texas at Austin': 'University of Texas at Austin',
        'Human Resources (HR) & Office Management': 'Human Resources & Office Management',

        # Standardize case for PSH activities
        'PSH CASEWORKER UPDATES AND CALLS': 'PSH Caseworker Updates and Calls',
        'PSH HMIS updates and caseworker notes': 'PSH HMIS Updates and Caseworker Notes',
        'Community Fitness Gym': 'Community Fitness Gym',
        'Caseworker calls for PSH': 'Caseworker Calls for PSH',
        'PSH caseworker and BMHC updates': 'PSH Caseworker and BMHC Updates',

        # Meeting-related activities
        'Outreach Team Meeting': 'Outreach Team Meeting',
        'Agency Partnership/Collaboration ': 'Agency Partnership/Collaboration',
        'Kensington Integral Care housing ': 'Kensington Integral Care Housing',
        'community partnerships/engagement': 'Community Partnerships/Engagement',

        # Correct spelling of "Referrals"
        'Referals': 'Referrals',

        # Unmatched Care Network Activities
        'Administrative & Communications': 'Administrative & Communications',
        '1 to 1 Outreach Strategy Meetings': '1 to 1 Outreach Strategy Meetings',
        'Outreach & Engagement': 'Outreach & Engagement',
        'Stakeholder & Key Leader Meetings': 'Stakeholder & Key Leader Meetings',
        'Research & Planning': 'Research & Planning',
        'Reports & Documentation': 'Reports & Documentation',
        'Special Event Support': 'Special Event Support',
        'Financial & Budgeting': 'Financial & Budgeting',
        'Human Resources (HR) & Office Management': 'Human Resources (HR) & Office Management',
        'BMHC & GUD LIFE Huddle Meetings': 'BMHC & GUD LIFE Huddle Meetings',
        'Performance & Reviews': 'Performance & Reviews',
        'Training & Onboarding': 'Training & Onboarding',
        'PSH & Client Support': 'PSH & Client Support'
    })
)

df_care = df[df['Care Activity'].notna()]

# Find any remaining unmatched purposes
unmatched_care = df_care[~df_care['Care Activity'].isin(care_categories)]['Care Activity'].unique().tolist()

# Find any remaining unmatched purposes
unmatched_care = df_care[~df_care['Care Activity'].isin(care_categories)]['Care Activity'].unique().tolist()
# print("Unmatched Care Network Activities:", unmatched_care)

# Group the data by 'Month' and 'Admin Activity' and count occurrences
df_care_counts = (
    df_care.groupby(['Month', 'Care Activity'], sort=True)
    .size()
    .reset_index(name='Count')
)

# Assign categorical ordering to the 'Month' column
df_care_counts['Month'] = pd.Categorical(
    df_care_counts['Month'],
    categories=months_in_quarter,
    ordered=True
)

# Sort df:
df_care_counts = df_care_counts.sort_values(by=['Month', 'Care Activity'])

# Create the grouped bar chart
care_fig = px.bar(
    df_care_counts,
    x='Month',
    y='Count',
    color='Care Activity',
    barmode='group',
    text='Count',
    title= f'{current_quarter} Care Network Activities by Month',
    labels={
        'Count': 'Number of Activities',
        'Month': 'Month',
        'Care Activity': 'Care Network Activity'
    }
).update_layout(
    xaxis_title='Month',
    yaxis_title='Count',
    height=900,  # Adjust graph height
    title=dict(
        text= f'{current_quarter} Care Network Activities by Month',
        x=0.5, 
        font=dict(
            size=35,
            family='Calibri',
            color='black',
            )
    ),
    font=dict(
        family='Calibri',
        size=17,
        color='black'
    ),
    xaxis=dict(
        tickmode='array',
        tickvals=df_care_counts['Month'].unique(),
        tickangle=-35  # Rotate x-axis labels for better readability
    ),
    legend=dict(
        # title='Administrative Activity',
        title=None,
        orientation="v",  # Vertical legend
        x=1.05,  # Position legend to the right
        xanchor="left",  # Anchor legend to the left
        y=1,  # Position legend at the top
        yanchor="top"  # Anchor legend at the top
    ),
    hovermode='x unified'  # Display unified hover info
).update_traces(
    textposition='outside',  # Display text above bars
    textfont=dict(size=30),  # Increase text size in each bar
    hovertemplate=(
        '<br>'
        '<b>Count: </b>%{y}<br>'  # Count
    ),
    customdata=df_care_counts['Care Activity'].values.tolist()
)

df_care = df_care.groupby('Care Activity').size().reset_index(name='Count')

# Create the pie chart for Administrative Activity distribution
care_pie = px.pie(
    df_care,
    names='Care Activity',
    values='Count',
    color='Care Activity',
    height=800,
    title= f'{current_quarter} Distribution of Care Network Activities'
).update_layout(
    title=dict(
        x=0.5,
        text= f'{current_quarter} Distribution of Care Network Activities',  # Title text
        font=dict(
            size=35,  # Increase this value to make the title bigger
            family='Calibri',  # Optional: specify font family
            color='black'  # Optional: specify font color
        ),
    ),  # Center-align the title
    margin=dict(
        t=150,  # Adjust the top margin (increase to add more padding)
        l=20,   # Optional: left margin
        r=20,   # Optional: right margin
        b=20    # Optional: bottom margin
    )
).update_traces(
    rotation=140,  # Rotate pie chart 90 degrees counterclockwise
    textfont=dict(size=19),  # Increase text size
    textinfo='value+percent',
    # texttemplate='<br>%{percent:.1%}',  # Format percentage as whole numbers
    hovertemplate='<b>%{label}</b>: %{value}<extra></extra>'  # Hover details
)

# --------------------- Community Outreach Activity DF ------------------------ #

# Value counts for 'Outreach Activity'
outreach_value_counts = df['Outreach Activity'].value_counts()

# Convert the Series to a DataFrame
outreach_value_counts_df = outreach_value_counts.reset_index()
outreach_value_counts_df.columns = ['Outreach Activity', 'Count']  # Rename columns

# Save the DataFrame to an Excel file
outreach_output_path = os.path.join(script_dir, 'outreach_activity_counts.xlsx')

outreach_value_counts_df.to_excel(outreach_output_path, index=False)
print(f"Outreach activity counts saved to {outreach_output_path}")

# print("Community Outreach Activities Unique Before:", df['Outreach Activity'].unique().tolist())
# print("Community Outreach Activities Value Counts: \n", outreach_value_counts)

comm_unique = [
    '', 'Meeting', 'Advocacy', 'Healthy Cuts Event', 'Presentation', 'Onsite Outreach ', 'Movement is medicine', 'Weekly Meeting Updates', 'NA', 'NA - Team Meeting', 'Movement is Medicine', ' Movement is Medicine', 'Potential partnering for mammogram services on site.', 'Healthy Cuts/Know Your Numbers Event at Community First Village', 'CTAAF Conference Presentation (advocacy of BMHC + AMEN movement is medicine ) ', 'BMHC Weekly Team Huddle ', 'Outreach 1 to 1 Strategy Meetings', 'Community First Village Onsite Outreach', 'Movement Is Medicine', 'Downtown Austin Community Court Onsite Outreach', 'Tabling', 'BMHC + KAZI Basketball Tournament', 'Outreach & Navigation', 'Health Event', 'ECHO Pilot Program ', 'Advocacy, Tabling, Presentation', 'Coordination of services', 'Collaboration', 'PSH Caseworker calls and updates', 'PSH HMIS Updates', 'PSH File updates', 'Collaboration of development of co-programs (ministry and GUD LIFE)', 'Discovery Meeting: Learn about each organization’s mission, values, and potential alignment.', 'psh updates', 'build relationship ', 'Building Relationships ', 'meeting via phone'
]

comm_categories = [
    'Advocacy & Presentations',
    'Outreach Activities',
    'Meetings',
    'PSH & Case Management',
    'Health Events',
    'Miscellaneous',
]

# print("Community Outreach Activity Unique Before:", df['Outreach Activity'].unique().tolist())

df['Outreach Activity'] = (
    df['Outreach Activity']
    .astype(str)
    .str.strip()
    .replace({
        
        "" : pd.NA,
        "<NA>" : pd.NA,
        
        # Advocacy & Presentations
        'Advocacy': 'Advocacy & Presentations',
        'Presentation': 'Advocacy & Presentations',
        'CTAAF Conference Presentation (advocacy of BMHC + AMEN movement is medicine )': 'Advocacy & Presentations',
        'Advocacy, Tabling, Presentation': 'Advocacy & Presentations',
        
        # Outreach Activities
        'Onsite Outreach ': 'Outreach Activities',
        'Community First Village Onsite Outreach': 'Outreach Activities',
        'Downtown Austin Community Court Onsite Outreach': 'Outreach Activities',
        'Outreach & Navigation': 'Outreach Activities',
        'Healthy Cuts/Know Your Numbers Event at Community First Village': 'Outreach Activities',
        'Healthy Cuts Event': 'Outreach Activities',
        'Outreach 1 to 1 Strategy Meetings': 'Outreach Activities',
        'BMHC + KAZI Basketball Tournament': 'Outreach Activities',
        
        # Meetings
        'Meeting': 'Meetings',
        'Weekly Meeting Updates': 'Meetings',
        'BMHC Weekly Team Huddle ': 'Meetings',
        'NA - Team Meeting': 'Meetings',
        'NA': 'Meetings',
        'Movement is medicine': 'Meetings',  # Assuming this can be categorized as a type of recurring meeting/event
        'Movement Is Medicine': 'Meetings',
        ' Movement is Medicine': 'Meetings',  # Handling spacing issues
        
        # PSH & Case Management
        'PSH Caseworker calls and updates': 'PSH & Case Management',
        'PSH HMIS Updates': 'PSH & Case Management',
        'PSH File updates': 'PSH & Case Management',
        'psh updates': 'PSH & Case Management',
        'Building Relationships ': 'PSH & Case Management',
        'build relationship ': 'PSH & Case Management',
        'Coordination of services': 'PSH & Case Management',
        
        # Health Events
        'Health Event': 'Health Events',
        'Movement is medicine': 'Health Events',
        'Healthy Cuts/Know Your Numbers Event at Community First Village': 'Health Events',
        
        # Miscellaneous
        'Tabling': 'Miscellaneous',
        'Potential partnering for mammogram services on site.': 'Miscellaneous',
        'Discovery Meeting: Learn about each organization’s mission, values, and potential alignment.': 'Miscellaneous',
        'meeting via phone': 'Miscellaneous',
        'Collaboration': 'Miscellaneous',
        'Collaboration of development of co-programs (ministry and GUD LIFE)': 'Miscellaneous',
        
        # Unmatched Community Outreach Activities
        'Onsite Outreach': 'Outreach Activities',
        'Movement is Medicine': 'Meetings',
        'BMHC Weekly Team Huddle': 'Meetings',
        'ECHO Pilot Program': 'Health Events',
        'build relationship': 'PSH & Case Management',
        'Building Relationships': 'PSH & Case Management',
    })
)

df_comm = df[df['Outreach Activity'].notna()]

# Find any remaining unmatched purposes
unmatched_comm = df_comm[~df_comm['Outreach Activity'].isin(comm_categories)]['Outreach Activity'].unique().tolist()
# print("Unmatched Community Outreach Activities:", unmatched_comm)

# print("Community Outreach Activity Unique After:", df['Outreach Activity'].unique().tolist())

# Group the data by 'Month' and 'Community Outreach Activity:' and count occurrences
df_comm_counts = (
    df.groupby(['Month', 'Outreach Activity'], sort=False)
    .size()
    .reset_index(name='Count')
)

# Assign categorical ordering to the 'Month' column
df_comm_counts['Month'] = pd.Categorical(
    df_comm_counts['Month'],
    categories=months_in_quarter,
    ordered=True
)

# Sort df
df_comm_counts = df_comm_counts.sort_values(by=['Month', 'Outreach Activity'])

# Create the grouped bar chart
comm_fig = px.bar(
    df_comm_counts,
    x='Month',
    y='Count',
    color='Outreach Activity',
    barmode='group',
    text='Count',
    labels={
        'Count': 'Number of Activities',
        'Month': 'Month',
        'Outreach Activity': 'Community Outreach Activity'
    }
).update_layout(
    xaxis_title='Month',
    yaxis_title='Count',
    height=900,  # Adjust graph height
    title=dict(
        text= f'{current_quarter} Community Outreach Activities By Month',
        x=0.5, 
        font=dict(
            size=35,
            family='Calibri',
            color='black',
            )
    ),
    font=dict(
        family='Calibri',
        size=17,
        color='black'
    ),
    xaxis=dict(
        tickmode='array',
        tickvals=df_comm_counts['Month'].unique(),
        tickangle=-35  # Rotate x-axis labels for better readability
    ),
    legend=dict(
        title='',
        orientation="v",  # Vertical legend
        x=1.05,  # Position legend to the right
        xanchor="left",  # Anchor legend to the left
        y=1,  # Position legend at the top
        yanchor="top"  # Anchor legend at the top
    ),
    hovermode='x unified'  # Display unified hover info
).update_traces(
    textposition='outside',  # Display text above bars
    textfont=dict(size=30),  
    hovertemplate=(
        '<br>'
        '<b>Count: </b>%{y}<br>'  # Count
    ),
    customdata=df_comm_counts['Outreach Activity'].values.tolist()
)

df_comm = df_comm.groupby('Outreach Activity').size().reset_index(name='Count')

# Create the pie chart for Administrative Activity distribution
comm_pie = px.pie(
    df_comm,
    names='Outreach Activity',
    values='Count',
    color='Outreach Activity',
    height=800,
    title= f'{current_quarter} Distribution of Community Outreach Activities'
).update_layout(
    title=dict(
        x=0.5,
        text= f'{current_quarter} Distribution of Community Outreach Activities',  # Title text
        font=dict(
            size=35,  # Increase this value to make the title bigger
            family='Calibri',  # Optional: specify font family
            color='black'  # Optional: specify font color
        ),
    ),  # Center-align the title
    margin=dict(
        t=250,  # Adjust the top margin (increase to add more padding)
        l=20,   # Optional: left margin
        r=20,   # Optional: right margin
        b=20    # Optional: bottom margin
    )
).update_traces(
    rotation=140, 
    textfont=dict(size=19),  # Increase text size
    textinfo='value+percent',
    #  texttemplate='<br>%{value}\n %{percent:.1%}',  # Format to show both value and percentage
    hovertemplate='<b>%{label}</b>: %{percent}<extra></extra>'  # Hover details
)

# ------------------------ Person Submitting Form DF ---------------------------- #

person_unique = [
    'Larry Wallace Jr.', 
    'Cameron Morgan',
    'Sonya Hosey', 
    'Kiounis Williams', 
    'Antonio Montgomery', 
    'Toya Craney', 
    'KAZI 88.7 FM Radio Interview & Preparation', 
    'Kim Holiday', 
    'Jordan Calbert', 
    'Dominique Street', 
    'Eric Roberts'
]

# print("Person Unique Before:", df["Person submitting this form:"].unique().tolist())

# Create a new dataframe with 'Person' and 'Date of Activity'
df_person = df[['Person', 'Date of Activity']].copy()

# Remove trailing whitespaces and perform the replacements
df['Person'] = (
    df['Person']
    .str.strip()
    .replace({
        "Larry Wallace Jr": "Larry Wallace Jr.",
        "`Larry Wallace Jr": "Larry Wallace Jr.",
        "Antonio Montggery": "Antonio Montgomery",
        "KAZI 88.7 FM Radio Interview & Preparation": "Larry Wallace Jr.",
    })
)

# Group the data by 'Month' and 'Person' and count occurrences
df_person_counts = (
    df.groupby(['Month', 'Person'], sort=True)
    .size()
    .reset_index(name='Count')
)

# Assign categorical ordering to the 'Month' column
df_person_counts['Month'] = pd.Categorical(
    df_person_counts['Month'],
    categories=months_in_quarter,
    ordered=True
)

# Sort df
df_person_counts = df_person_counts.sort_values(by=['Month', 'Person'])

# Create the grouped bar chart
person_fig = px.bar(
    df_person_counts,
    x='Month',
    y='Count',
    color='Person',
    barmode='group',
    text='Count',
    title=f'{current_quarter} Form Submissions by Month',
    labels={
        'Count': 'Number of Submissions',
        'Month': 'Month',
        'Person': 'Person'
    }
).update_layout(
    title_x=0.5,
    xaxis_title='Month',
    yaxis_title='Count',
    height=900,  # Adjust graph height
    title=dict(
        text= f'{current_quarter} Form Submissions by Month',
        x=0.5, # Center title
        font=dict(
            size=35,
            family='Calibri',
            color='black',
            )
    ),
    font=dict(
        family='Calibri',
        size=17,
        color='black'
    ),
    xaxis=dict(
        tickmode='array',
        tickvals=df_person_counts['Month'].unique(),
        tickangle=-35  # Rotate x-axis labels for better readability
    ),
    legend=dict(
        title='',
        orientation="v",  # Vertical legend
        x=1.05,  # Position legend to the right
        xanchor="left",  # Anchor legend to the left
        y=1,  # Position legend at the top
        yanchor="top"  # Anchor legend at the top
    ),
    hovermode='x unified',  # Display unified hover info
    bargap=0.08,  # Reduce the space between bars
    bargroupgap=0,  # Reduce space between individual bars in groups
    margin = dict(t=80, b=100, l=0, r=0),
).update_traces(
    textposition='outside',  # Display text above bars
    textfont=dict(size=30),  # Increase text size in each bar
    hovertemplate=(
        '<br>'
        '<b>Count: </b>%{y}<br>'  # Count
    ),
    customdata=df_person_counts['Person'].values.tolist()
).add_vline(
    x=0.5,  # Adjust the position of the line
    line_dash="dash",
    line_color="gray",
    line_width=2
).add_vline(
    x=1.5,  # Position of the second line
    line_dash="dash",
    line_color="gray",
    line_width=2
)

# Group by person submitting form:
df_pf = df.groupby('Person').size().reset_index(name='Count')

# Pie chart:
person_pie = px.pie(
    df_pf,
    names='Person',
    values='Count',
    color='Person',
    height=850
).update_layout(
    title=dict(
        x=0.5,
        text=f'{current_quarter} Distribution of Form Submissions',  # Title text
        font=dict(
            size=35,  # Increase this value to make the title bigger
            family='Calibri',  # Optional: specify font family
            color='black'  # Optional: specify font color
        ),
    ),
    legend=dict(
        # title='',
        title=None,
        orientation="v",  # Vertical legend
        x=1.05,  # Position legend to the right
        xanchor="left",  # Anchor legend to the left
        y=1,  # Position legend at the top
        yanchor="top"  # Anchor legend at the top
    ),
    margin = dict(t=80, b=0, l=0, r=0),
).update_traces(
    rotation=70,  # Rotate pie chart 90 degrees counterclockwise
    textfont=dict(size=19),  # Increase text size in each bar
    texttemplate='%{value}<br>%{percent:.1%}',  # Format percentage as whole numbers
    hovertemplate='<b>%{label}</b>: %{value}<extra></extra>'
)

# # ========================== DataFrame Table ========================== #

# Engagement Table
engagement_table = go.Figure(data=[go.Table(
    # columnwidth=[50, 50, 50],  # Adjust the width of the columns
    header=dict(
        values=list(df.columns),
        fill_color='paleturquoise',
        align='center',
        height=30,  # Adjust the height of the header cells
        # line=dict(color='black', width=1),  # Add border to header cells
        font=dict(size=12)  # Adjust font size
    ),
    cells=dict(
        values=[df[col] for col in df.columns],
        fill_color='lavender',
        align='left',
        height=25,  # Adjust the height of the cells
        # line=dict(color='black', width=1),  # Add border to cells
        font=dict(size=12)  # Adjust font size
    )
)])

engagement_table.update_layout(
    margin=dict(l=50, r=50, t=30, b=60),  # Remove margins
    height=800,
    # width=1500,  # Set a smaller width to make columns thinner
    paper_bgcolor='rgba(0,0,0,0)',  # Transparent background
    plot_bgcolor='rgba(0,0,0,0)'  # Transparent plot area
)

# Entity Name Table
# entity_name_table = go.Figure(data=[go.Table(
#     header=dict(
#         values=list(entity_name_group.columns),
#         fill_color='paleturquoise',
#         align='center',
#         height=30,
#         font=dict(size=12)
#     ),
#     cells=dict(
#         values=[entity_name_group[col] for col in entity_name_group.columns],
#         fill_color='lavender',
#         align='left',
#         height=25,
#         font=dict(size=12)
#     )
# )])

# entity_name_table.update_layout(
#     margin=dict(l=50, r=50, t=30, b=40),
#     height=400,
#     paper_bgcolor='rgba(0,0,0,0)',
#     plot_bgcolor='rgba(0,0,0,0)'
# )

# ============================== Dash Application ========================== #

app = dash.Dash(__name__)
server= app.server 

app.layout = html.Div(
  children=[ 
    html.Div(
        className='divv', 
        children=[ 
          html.H1(
              f'BMHC Partner Engagement Report {current_quarter} 2025', 
              className='title'),
          html.H2( 
              '01/01/2025 - 3/31/2024', 
              className='title2'),
          html.Div(
              className='btn-box', 
              children=[
                  html.A(
                    'Repo',
                    href= f'https://github.com/CxLos/Eng_{current_quarter}_2025',
                    className='btn'),
    ]),
  ]),    

# Data Table
# html.Div(
#     className='row00',
#     children=[
#         html.Div(
#             className='graph00',
#             children=[
#                 html.Div(
#                     className='table',
#                     children=[
#                         html.H1(
#                             className='table-title',
#                             children='Engagements Table'
#                         )
#                     ]
#                 ),
#                 html.Div(
#                     className='table2', 
#                     children=[
#                         dcc.Graph(
#                             className='data',
#                             figure=engagement_table
#                         )
#                     ]
#                 )
#             ]
#         ),
#     ]
# ),

# ROW 1
html.Div(
    className='row0',
    children=[
        html.Div(
            className='graph11',
            children=[
            html.Div(
                className='high1',
                children=[f'{current_quarter} Engagements']
            ),
            html.Div(
                className='circle1',
                children=[
                    html.Div(
                        className='hilite',
                        children=[
                            html.H1(
                            className='high3',
                            children=[total_engagements]
                    ),
                        ]
                    )
 
                ],
            ),
            ]
        ),
        html.Div(
            className='graph22',
            children=[
            html.Div(
                className='high2',
                children=[f'{current_quarter} Engagement Hours']
            ),
            html.Div(
                className='circle2',
                children=[
                    html.Div(
                        className='hilite',
                        children=[
                            html.H1(
                            className='high4',
                            children=[eng_hours]
                    ),
                        ]
                    )
 
                ],
            ),
            ]
        ),
    ]
),

# ROW 1
html.Div(
    className='row1',
    children=[
        html.Div(
            className='graph11',
            children=[
            html.Div(
                className='high1',
                children=[f'{current_quarter} Travel Hours']
            ),
            html.Div(
                className='circle1',
                children=[
                    html.Div(
                        className='hilite',
                        children=[
                            html.H1(
                            className='high3',
                            children=[total_travel_time]
                    ),
                        ]
                    )
 
                ],
            ),
            ]
        ),
        html.Div(
            className='graph2',
            children=[
                dcc.Graph(
                    # figure=status_fig
                )
            ]
        ),
    ]
),

# ROW 1
# html.Div(
#     className='row1',
#     children=[
#         html.Div(
#             className='graph1',
#             children=[
#                 dcc.Graph(
#                     # figure=status_fig
#                 )
#             ]
#         ),
#         html.Div(
#             className='graph2',
#             children=[
#                 dcc.Graph(
#                     # figure=status_fig
#                 )
#             ]
#         ),
#     ]
# ),

# ROW 1
html.Div(
    className='row1',
    children=[
        html.Div(
            className='graph1',
            children=[
                dcc.Graph(
                    figure=hours_fig
                )
            ]
        ),
        html.Div(
            className='graph2',
            children=[
                dcc.Graph(
                    figure=hours_pie
                )
            ]
        ),
    ]
),

# ROW 1
html.Div(
    className='row1',
    children=[
        html.Div(
            className='graph1',
            children=[
                dcc.Graph(
                    figure=travel_fig
                )
            ]
        ),
        html.Div(
            className='graph2',
            children=[
                dcc.Graph(
                    figure=travel_pie
                )
            ]
        ),
    ]
),

# ROW 1
html.Div(
    className='row1',
    children=[
        html.Div(
            className='graph1',
            children=[
                dcc.Graph(
                    figure=status_fig
                )
            ]
        ),
        html.Div(
            className='graph2',
            children=[
                dcc.Graph(
                    figure=status_pie
                )
            ]
        ),
    ]
),

# ROW 
html.Div(
    className='row3',
    children=[
        html.Div(
            className='graph0',
            children=[
                dcc.Graph(
                    figure=admin_fig
                )
            ]
        )
    ]
),
# ROW 
html.Div(
    className='row3',
    children=[
        html.Div(
            className='graph0',
            children=[
                dcc.Graph(
                    figure=admin_pie
                )
            ]
        )
    ]
),

# ROW 
html.Div(
    className='row3',
    children=[
        html.Div(
            className='graph0',
            children=[
                dcc.Graph(
                    figure=care_fig
                )
            ]
        )
    ]
),

# ROW 
html.Div(
    className='row3',
    children=[
        html.Div(
            className='graph0',
            children=[
                dcc.Graph(
                    figure=care_pie
                )
            ]
        )
    ]
),

# ROW 
html.Div(
    className='row3',
    children=[
        html.Div(
            className='graph0',
            children=[
                dcc.Graph(
                    figure=comm_fig
                )
            ]
        )
    ]
),

# ROW 
html.Div(
    className='row3',
    children=[
        html.Div(
            className='graph0',
            children=[
                dcc.Graph(
                    figure=comm_pie
                )
            ]
        )
    ]
),

# ROW 
html.Div(
    className='row3',
    children=[
        html.Div(
            className='graph0',
            children=[
                dcc.Graph(
                    figure=person_fig
                )
            ]
        )
    ]
),

# ROW 
html.Div(
    className='row3',
    children=[
        html.Div(
            className='graph0',
            children=[
                dcc.Graph(
                    figure=person_pie
                )
            ]
        )
    ]
),
])

print(f"Serving Flask app '{current_file}'! 🚀")

if __name__ == '__main__':
    app.run_server(debug=
                   True)
                #    False)
# =================================== Updated Database ================================= #

# updated_path = f'data/Engagement_{current_quarter}_{report_year}.xlsx'
# data_path = os.path.join(script_dir, updated_path)
# df.to_excel(data_path, index=False)
# print(f"DataFrame saved to {data_path}")

# updated_path1 = 'data/service_tracker_q4_2024_cleaned.csv'
# data_path1 = os.path.join(script_dir, updated_path1)
# df.to_csv(data_path1, index=False)
# print(f"DataFrame saved to {data_path1}")

# -------------------------------------------- KILL PORT ---------------------------------------------------

# netstat -ano | findstr :8050
# taskkill /PID 24772 /F
# npx kill-port 8050

# ---------------------------------------------- Host Application -------------------------------------------

# 1. pip freeze > requirements.txt
# 2. add this to procfile: 'web: gunicorn impact_11_2024:server'
# 3. heroku login
# 4. heroku create
# 5. git push heroku main

# Create venv 
# virtualenv venv 
# source venv/bin/activate # uses the virtualenv

# Update PIP Setup Tools:
# pip install --upgrade pip setuptools

# Install all dependencies in the requirements file:
# pip install -r requirements.txt

# Check dependency tree:
# pipdeptree
# pip show package-name

# Remove
# pypiwin32
# pywin32
# jupytercore

# ----------------------------------------------------

# Name must start with a letter, end with a letter or digit and can only contain lowercase letters, digits, and dashes.

# Heroku Setup:
# heroku login
# heroku create mc-impact-11-2024
# heroku git:remote -a mc-impact-11-2024
# git push heroku main

# Clear Heroku Cache:
# heroku plugins:install heroku-repo
# heroku repo:purge_cache -a mc-impact-11-2024

# Set buildpack for heroku
# heroku buildpacks:set heroku/python

# Heatmap Colorscale colors -----------------------------------------------------------------------------

#   ['aggrnyl', 'agsunset', 'algae', 'amp', 'armyrose', 'balance',
            #  'blackbody', 'bluered', 'blues', 'blugrn', 'bluyl', 'brbg',
            #  'brwnyl', 'bugn', 'bupu', 'burg', 'burgyl', 'cividis', 'curl',
            #  'darkmint', 'deep', 'delta', 'dense', 'earth', 'edge', 'electric',
            #  'emrld', 'fall', 'geyser', 'gnbu', 'gray', 'greens', 'greys',
            #  'haline', 'hot', 'hsv', 'ice', 'icefire', 'inferno', 'jet',
            #  'magenta', 'magma', 'matter', 'mint', 'mrybm', 'mygbm', 'oranges',
            #  'orrd', 'oryel', 'oxy', 'peach', 'phase', 'picnic', 'pinkyl',
            #  'piyg', 'plasma', 'plotly3', 'portland', 'prgn', 'pubu', 'pubugn',
            #  'puor', 'purd', 'purp', 'purples', 'purpor', 'rainbow', 'rdbu',
            #  'rdgy', 'rdpu', 'rdylbu', 'rdylgn', 'redor', 'reds', 'solar',
            #  'spectral', 'speed', 'sunset', 'sunsetdark', 'teal', 'tealgrn',
            #  'tealrose', 'tempo', 'temps', 'thermal', 'tropic', 'turbid',
            #  'turbo', 'twilight', 'viridis', 'ylgn', 'ylgnbu', 'ylorbr',
            #  'ylorrd'].

# rm -rf ~$bmhc_data_2024_cleaned.xlsx
# rm -rf ~$bmhc_data_2024.xlsx
# rm -rf ~$bmhc_q4_2024_cleaned2.xlsx
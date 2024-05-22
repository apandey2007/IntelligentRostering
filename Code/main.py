## The files contain below columns:
## Agent-Availability: Date,Agent_ID, Agent_Name, Shift_Timing_Start, Shift_Timing_End, STATION, STATION_Duration
## Contract-Information: Contract_Start_Date, Contract_End_Date, Airline, STATION
## Airline-Schedule:Date, Arrival, Flight, Airline
## Step 1: Check Schedule information, select unique Airlines from schedule file; column name: "Airline"
## Step 2: Cross-check with Contract-Information file, on="Airline" and remove those airlines from Schedule Information which do not have valid contract based on current date and "Contract_Start_Date" and "Contract_End_Date"
## Step 3: Show these removed airlines and Reason for Rejection in a table in streamlit app
## Step 4: Add a column "STATION" from the merge in Step 2, for all valid STATIONS for each airline in Contract-Information
## Step 5: Based on the flight arrival schedule and STATION_Duration (the amount of time required to complete a STATION), map the available agent based on their Shift_Stat_Time and Shift_End_Time and showcase the agent which will be present based on fight schedule
## Visualize in a bar chart with the Hour of the day in x-axis, Duration across a STATION on y-axis, Agent-STATION mapping as color and a filter for date
## All this should be shown in the same streamlit application, with grid for uploading the three input excel files namely: Agent-Availability, Airline-Schedule and Contract-Information with Airlines
## After processing the data, the app should return button which allows an excel to be downloaded with below sheets:
##  1. Valid Airlines for each day, 2. Airlines Invalid Contracts, 3. Agent Schedule showing "Date", "Agent", "Flight", "Airline", "Arrival", "STATION" and "STATION_Duration"

import pandas as pd
import streamlit as st
from datetime import date
import plotly.graph_objects as go
import plotly.express as px
from io import BytesIO


def get_data(file_uploader, header=0):
  """Reads data from uploaded excel file"""
  if file_uploader is not None:
    try:
      byte_data = file_uploader.read()
      data = pd.read_excel(BytesIO(byte_data), header=header)
      return data
    except Exception as e:
      st.error(f"Error reading file: {e}")
      return None
  else:
    return None

def filter_valid_contracts(schedule, contracts, today):
  """Filters airlines with valid contracts based on current date"""
  if schedule is not None and contracts is not None:
    merged_data = schedule.merge(contracts[['Airline', 'Contract_Start_Date', 'Contract_End_Date']], how='left', on='Airline')

    # Convert 'Contract_Start_Date' and 'Contract_End_Date' to datetime format (assuming they're strings)
    merged_data['Contract_Start_Date'] = pd.to_datetime(merged_data['Contract_Start_Date'])
    merged_data['Contract_End_Date'] = pd.to_datetime(merged_data['Contract_End_Date'])

    # Extract the date part from both sides of the comparison for valid check
    merged_data['Valid_Contract'] = (merged_data['Contract_Start_Date'].dt.date <= today) & (merged_data['Contract_End_Date'].dt.date >= today)
    valid_airlines = merged_data[merged_data['Valid_Contract']]['Airline'].unique()
    return schedule[schedule['Airline'].isin(valid_airlines)], merged_data[~merged_data['Valid_Contract']]
  else:
    return None, None

def map_agents(flights, agents, contracts, today):
  """Maps available agents based on flight arrival, shift timings, station duration, and agent availability"""
  if flights is not None and agents is not None and contracts is not None:
    # Convert 'today' to pandas Timestamp object
    today_timestamp = pd.Timestamp(today)

    # Filter flights with valid contracts based on airline and contract dates
    valid_airlines = contracts[(contracts['Contract_Start_Date'] <= today_timestamp) &
                               (contracts['Contract_End_Date'] >= today_timestamp)]['Airline']
    valid_flights = flights[flights['Airline'].isin(valid_airlines)]

    # Keep only the required stations for each flight of each valid airline
    valid_stations = contracts[['Airline', 'STATION']].drop_duplicates()
    valid_flights = valid_flights.merge(valid_stations, on='Airline')

    # Merge agents with flights based on station
    merged_data = valid_flights.merge(agents, how='cross', suffixes=('_flight', '_agent'))

    # Filter agents based on shift timings
    merged_data['Arrival_Time'] = pd.to_datetime(merged_data['Arrival'], format='%H:%M:%S', errors='coerce')
    merged_data['Start_Time'] = pd.to_datetime(merged_data['Shift_Timing_Start'], format='%H:%M:%S', errors='coerce')
    merged_data['End_Time'] = pd.to_datetime(merged_data['Shift_Timing_End'], format='%H:%M:%S', errors='coerce')

    # Calculate the remaining available time for each agent based on station duration
    merged_data['Available_Time'] = merged_data.apply(
      lambda row: min((row['End_Time'] - row['Start_Time']).seconds / 3600,
                      row['STATION_Duration']), axis=1)

    # Filter agents whose shift timings overlap with flight arrival time and have enough available time
    valid_agents = merged_data[(merged_data['Arrival_Time'] >= merged_data['Start_Time']) &
                               (merged_data['Arrival_Time'] < merged_data['End_Time']) &
                               (merged_data['Available_Time'] > 0) &
                               (merged_data['STATION_flight'] == merged_data['STATION_agent'])]
    valid_agents = valid_agents.reset_index().drop_duplicates()
    return valid_agents[['Date_flight', 'Arrival', 'Flight', 'Airline', 'STATION_flight', 'STATION_agent', 'STATION_Duration', 'Agent_ID', 'Agent_Name']]
  else:
    return None


def download_excel(data, sheet_names):
  """Downloads data as an excel file with specified sheet names"""
  # Create a BytesIO object to store the Excel file
  excel_buffer = BytesIO()
  with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
    for i, sheet_data in enumerate(data):
      sheet_data.to_excel(writer, sheet_name=sheet_names[i], index=False)
  # Seek to the beginning of the buffer
  excel_buffer.seek(0)
  # Provide the download button with the Excel file content
  st.download_button(label="Download Excel", data=excel_buffer, file_name='Agents_Flight_Schedule.xlsx',
                     mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

def plot_agent_schedule(data):
  """Creates a bar chart with plotly"""
  fig = px.bar(  # Using px.bar from plotly express
      data,
      x='Arrival',
      y='STATION_Duration',
      color='Agent_Name',
      title="Agent Schedule - Duration by Hour",
      barmode='group',
      labels={
          'Arrival': 'Hour of the Day',
          'STATION_agent': 'STATION of Agent'
      }
  )
  st.plotly_chart(fig)

# App Title and File Upload
st.title("Airline Schedule and Agent Availability")

col1, col2, col3 = st.columns(3)

with col1:
  agent_availability_file = st.file_uploader("Upload Agent Availability", type="xlsx")

with col2:
  airline_schedule_file = st.file_uploader("Upload Airline Schedule", type="xlsx")

with col3:
  contract_information_file = st.file_uploader("Upload Contract Information", type="xlsx")

# Load data from uploaded files
agent_availability_data = get_data(agent_availability_file)
airline_schedule_data = get_data(airline_schedule_file)
contract_information_data = get_data(contract_information_file)

# Check for errors in uploaded files
if any(df is None for df in [agent_availability_data, airline_schedule_data, contract_information_data]):
  st.error("Upload all files before processing!")
else:
  today = date.today()

  # Filter airlines with valid contracts
  filtered_schedule_data, invalid_contracts_data = filter_valid_contracts(airline_schedule_data.copy(), contract_information_data.copy(), today)

  # Map available agents based on flight arrival, timings, and station duration
  processed_data = map_agents(filtered_schedule_data.copy(), agent_availability_data.copy(), contract_information_data, today)

  # Download processed data (optional)
  if st.button("Process Data"):
    download_excel([processed_data, invalid_contracts_data], ["Processed_Data", "Invalid_Contracts"])

  # # Display processed data (optional)
  if invalid_contracts_data is not None:
    st.dataframe(invalid_contracts_data[["Date", "Arrival", "Airline"]].drop_duplicates(), hide_index=True)

  # Generate and display agent schedule chart (optional)
  if processed_data is not None:
    plot_agent_schedule(processed_data.copy())

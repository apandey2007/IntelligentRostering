## The files contain below columns:
## Agent-Availability: Date	Agent_ID	Agent_Name	Shift_Timing_Start	Shift_Timing_End	Expertise
## Contract-Information: Contract_Start_Date	Contract_End_Date	Airline_Name	STATION
## Airline-Schedule:Date	Arrival	Flight	Airline
## Step 1: Check Schedule information, select unique Airlines from schedule file; column name: "Airline"
## Step 2: Merge with Contract, on="Airline" and remove those airlines from Schedule Information which do not have valid contract based on current date and "Contract_Start_Date" and "Contract_End_Date"
## Step 3: Show these removed airlines and Reason for Rejection in a table in streamlit app
## Step 4: Add a column "STATION" from the merge in Step 2, for all valid STATIONS for the airlines in contract
## Step 5: Based on the flight schedule and STATION_Duration (the amount of time required to complete a STATION), map the available agent and showcase the agent which will be present based on fight schedule
## Visualize the flight and associated agent in best possible chart
## All this should be shown in the same streamlit application, with grid for uploading the three input excel files namely: Agent Availability, Airline Schedule and Contract Information with Airlines

import streamlit as st
import pandas as pd
import plotly.express as px


# Function to read Excel files
def read_excel(file):
    df = pd.read_excel(file)
    return df


# Function to visualize flight schedule with assigned agents as Gantt chart
def visualize_flight_schedule(agent_availability_df, airline_schedule_df, contract_info_df):
    # Merge with contract and remove airlines without valid contracts
    merged_df = pd.merge(airline_schedule_df, contract_info_df, on="Airline", how="left")
    valid_airlines = merged_df[
        ~merged_df['Contract_Start_Date'].isnull() & (merged_df['Contract_Start_Date'] <= merged_df['Date']) & (
                    merged_df['Contract_End_Date'] >= merged_df['Date'])]
    removed_airlines = list(set(airline_schedule_df['Airline'].unique()) - set(valid_airlines['Airline'].unique()))

    # Show removed airlines and Reason for Rejection
    removed_airlines_table = pd.DataFrame({"Airline": removed_airlines, "Reason for Rejection": "No valid contract"})

    # Add STATION column for valid stations
    merged_df = valid_airlines

    # Merge available agents based on flight schedule
    merged_df = pd.merge(merged_df, agent_availability_df, on='Date', how='left')

    # Calculate arrival time as datetime object combining Date and Arrival hour
    merged_df['Arrival'] = merged_df.apply(
        lambda row: pd.to_datetime(row['Date'].date()) + pd.DateOffset(hours=row['Arrival'].hour), axis=1)
    # Calculate end time by adding station duration to arrival time
    merged_df['End_Time'] = merged_df['Arrival'] + pd.to_timedelta(merged_df['STATION_Duration'], unit='h')

    # Visualization using Gantt chart
    fig = px.timeline(merged_df, x_start='Arrival', x_end='End_Time', y='Flight', color='STATION_x', facet_row='Airline',
                      facet_col='Airline',
                      labels={'Flight': 'Flight Number', 'Agent_Name': 'Agent', 'Arrival': 'Arrival Time',
                              'End_Time': 'End Time'},
                      title='Flight Schedule with Assigned Agents (Gantt Chart)')
    fig.update_layout(xaxis_title='Time', yaxis_title='Flight Number', xaxis_tickangle=-45)

    # Display removed airlines table
    st.write("Airlines without valid contracts:")
    st.table(removed_airlines_table)

    # Display Gantt chart
    st.plotly_chart(fig)

    # Download the schedule as Excel
    filename = "flight_schedule.xlsx"
    with pd.ExcelWriter(filename) as writer:
        valid_airlines.to_excel(writer, sheet_name="Flight_Schedule", index=False)
        pd.DataFrame({"Airline": removed_airlines, "Reason for Rejection": "No valid contract"}).to_excel(writer,
                                                                                                          sheet_name="Rejected_Airlines",
                                                                                                          index=False)

        # Create a sheet for agents and associated STATION and flight
        agents_df = merged_df[['Date', 'Agent_Name', 'STATION_x', 'Flight']].dropna().drop_duplicates()
        agents_df.to_excel(writer, sheet_name="Agents_Schedule", index=False)

    # Download the Excel file
    st.markdown(f"### [Download {filename}]({filename})")


def main():
    # Title and description
    st.title('Flight Schedule Visualization')
    st.write(
        'This application visualizes flight schedules with assigned agents based on contract information and agent availability.')

    # Upload files
    agent_availability_file = st.file_uploader('Upload Agent Availability Excel File', type=['xls', 'xlsx'])
    airline_schedule_file = st.file_uploader('Upload Airline Schedule Excel File', type=['xls', 'xlsx'])
    contract_info_file = st.file_uploader('Upload Contract Information Excel File', type=['xls', 'xlsx'])

    if agent_availability_file and airline_schedule_file and contract_info_file:
        # Read Excel files
        agent_availability_df = read_excel(agent_availability_file)
        airline_schedule_df = read_excel(airline_schedule_file)
        contract_info_df = read_excel(contract_info_file)

        # Visualize flight schedule with assigned agents as Gantt chart
        visualize_flight_schedule(agent_availability_df, airline_schedule_df, contract_info_df)


if __name__ == "__main__":
    main()

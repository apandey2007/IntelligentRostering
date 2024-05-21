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
import plotly.graph_objects as go
from io import BytesIO


# Function to read Excel files
def read_excel(file):
    df = pd.read_excel(file)
    return df


# Function to visualize agent assignments as a bar chart
def visualize_agent_assignments(agent_availability_df, airline_schedule_df, contract_info_df, selected_date):
    # Combine date and time to create datetime objects for Shift_Timing_Start
    agent_availability_df['Shift_Timing_Start'] = agent_availability_df['Date'] + pd.to_timedelta(agent_availability_df['Shift_Timing_Start'].astype(str))
    # Combine date and time to create datetime objects for Shift_Timing_End
    agent_availability_df['Shift_Timing_End'] = agent_availability_df['Date'] + pd.to_timedelta(agent_availability_df['Shift_Timing_End'].astype(str))
    # Adjust Shift_Timing_End for cases where it represents the next day
    next_day_mask = agent_availability_df['Shift_Timing_End'].dt.hour == 0
    agent_availability_df.loc[next_day_mask, 'Shift_Timing_End'] += pd.Timedelta(days=1)

    agent_availability_df['Date'] = pd.to_datetime(agent_availability_df['Date'])

    filtered_agents = agent_availability_df[(agent_availability_df['Date'] == selected_date)]

    # Create a list of agents and their assignments for each station
    agent_assignments = []
    for _, row in filtered_agents.iterrows():
        for i in range(row['Shift_Timing_Start'].hour, row['Shift_Timing_End'].hour + 1):
            agent_assignments.append(
                {'Agent': row['Agent_Name'], 'Hour': i, 'Station': row['STATION'], 'Duration': row['STATION_Duration']})

    # Create a DataFrame from the list
    agent_assignments_df = pd.DataFrame(agent_assignments)

    # Group the DataFrame by agent and hour, and sum the durations for each station
    grouped_df = agent_assignments_df.groupby(['Date','Flight', 'Airline','Arrival', 'Agent', 'Hour', 'Station']).sum().reset_index()

    # Create a plot for each agent
    fig = go.Figure()

    for agent in grouped_df['Agent'].unique():
        agent_data = grouped_df[grouped_df['Agent'] == agent]
        for station in agent_data['Station'].unique():
            station_data = agent_data[agent_data['Station'] == station]
            fig.add_trace(go.Bar(x=station_data['Hour'], y=station_data['Duration'], name=agent + ' - ' + station))

    # Update the layout
    fig.update_layout(barmode='stack', xaxis_title='Hour of the Day', yaxis_title='Duration (Hours)',
                      title='Agent Assignments for ' + selected_date.strftime('%Y-%m-%d'))

    # Display the plot
    st.plotly_chart(fig)

    # Create Excel file
    with BytesIO() as output:
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            # Write sheets
            #valid_airlines.to_excel(writer, sheet_name="Flight_Schedule", index=False)
            agent_assignments_df.to_excel(writer, sheet_name="Assigned_Agents", index=False)

        # Retrieve the Excel file content as bytes
        excel_content = output.getvalue()

    # Download the Excel file
    st.download_button(label="Download Excel", data=excel_content, file_name="flight_schedule.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def main():
    # Title and description
    st.title('Agent Assignments Visualization')
    st.write('This application visualizes agent assignments based on contract information and agent availability.')

    # Upload files
    agent_availability_file = st.file_uploader('Upload Agent Availability Excel File', type=['xls', 'xlsx'])
    airline_schedule_file = st.file_uploader('Upload Airline Schedule Excel File', type=['xls', 'xlsx'])
    contract_info_file = st.file_uploader('Upload Contract Information Excel File', type=['xls', 'xlsx'])

    if agent_availability_file and airline_schedule_file and contract_info_file:
        # Read Excel files
        agent_availability_df = read_excel(agent_availability_file)
        airline_schedule_df = read_excel(airline_schedule_file)
        contract_info_df = read_excel(contract_info_file)

        # Filter unique dates from agent availability data
        available_dates = agent_availability_df['Date'].dt.date.unique()

        # Select date filter
        selected_date = st.selectbox('Select Date', available_dates)

        # Visualize agent assignments for selected date
        visualize_agent_assignments(agent_availability_df, airline_schedule_df, contract_info_df,
                                    pd.to_datetime(selected_date))


if __name__ == "__main__":
    main()

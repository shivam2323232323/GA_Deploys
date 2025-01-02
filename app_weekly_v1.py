import streamlit as st
import json
import xlsxwriter
from io import BytesIO
from google.oauth2.service_account import Credentials
from google.analytics.data_v1beta import BetaAnalyticsDataClient
from google.analytics.data_v1beta.types import DateRange, Metric, Dimension

# Streamlit app setup
st.title("GA Weekly Comparison Data")
st.markdown("### Compare weekly performance data for selected channels.")

# User inputs
st.sidebar.header("Input Details")
key_file = st.sidebar.file_uploader("Upload Service Account Key File", type="json")
property_id = st.sidebar.text_input("Enter GA4 Property ID", "")
this_week_start = st.sidebar.date_input("This Week Start Date")
this_week_end = st.sidebar.date_input("This Week End Date")
prev_week_start = st.sidebar.date_input("Previous Week Start Date")
prev_week_end = st.sidebar.date_input("Previous Week End Date")
selected_channels = st.sidebar.multiselect(
    "Select Channels",
    options=[
        "Cross-network", "Direct", "Display", "Email", "Mobile Push Notifications",
        "Organic Search", "Organic Shopping", "Organic Social", "Organic Video",
        "Paid Other", "Paid Search", "Paid Shopping", "Paid Social", "Referral", "Unassigned"
    ],
    default=["Organic Search", "Paid Search"]
)

if st.sidebar.button("Fetch and Download Data"):
    if not key_file or not property_id:
        st.error("Please upload a key file and enter the Property ID.")
    else:
        try:
            # Parse the uploaded key file
            key_file_data = json.loads(key_file.read().decode("utf-8"))
            credentials = Credentials.from_service_account_info(key_file_data)
            client = BetaAnalyticsDataClient(credentials=credentials)

            # Prepare weekly date ranges
            weeks = [
                {"name": "This Week", "start_date": this_week_start, "end_date": this_week_end},
                {"name": "Prev Week", "start_date": prev_week_start, "end_date": prev_week_end},
            ]

            # Fetch data
            weekly_data = {}
            for week in weeks:
                request = {
                    "property": f"properties/{property_id}",
                    "metrics": [
                        Metric(name="sessions"),
                        Metric(name="totalUsers"),
                        Metric(name="engagedSessions"),
                    ],
                    "dimensions": [
                        Dimension(name="sessionDefaultChannelGrouping"),
                    ],
                    "date_ranges": [DateRange(start_date=str(week["start_date"]), end_date=str(week["end_date"]))],
                }

                try:
                    response = client.run_report(request)
                    week_data = {}

                    for row in response.rows:
                        channel = row.dimension_values[0].value
                        if channel in selected_channels:
                            sessions = int(row.metric_values[0].value)
                            users = int(row.metric_values[1].value)
                            engaged_sessions = int(row.metric_values[2].value)
                            week_data[channel] = {
                                "sessions": sessions,
                                "users": users,
                                "engaged_sessions": engaged_sessions,
                            }

                    weekly_data[week["name"]] = week_data
                except Exception as e:
                    st.error(f"Error fetching data for {week['name']}: {e}")

            # Create Excel file in memory
            output = BytesIO()
            workbook = xlsxwriter.Workbook(output, {'in_memory': True})
            worksheet = workbook.add_worksheet("Weekly Comparison")

            # Set header formatting
            header_format = workbook.add_format({
                'bold': True,
                'font_color': 'white',
                'bg_color': '#2F75B5',  # Blue background
                'align': 'center',
                'valign': 'vcenter',
                'border': 1
            })

            # Set number formatting
            number_format = workbook.add_format({
                'align': 'center',
                'valign': 'vcenter',
                'border': 1
            })

            # Add table headers
            headers = ['Date', 'Users', 'Sessions', 'Engaged Sessions']
            worksheet.write(0, 0, "", header_format)  # Empty header for channels
            for col, header in enumerate(headers):
                worksheet.write(0, col + 1, header, header_format)

            # Populate the worksheet
            row = 1
            for channel in selected_channels:
                this_week_data = weekly_data.get("This Week", {}).get(channel, {"sessions": 0, "users": 0, "engaged_sessions": 0})
                prev_week_data = weekly_data.get("Prev Week", {}).get(channel, {"sessions": 0, "users": 0, "engaged_sessions": 0})

                change_data = {
                    "sessions": (
                        (this_week_data["sessions"] - prev_week_data["sessions"]) / prev_week_data["sessions"] * 100
                        if prev_week_data["sessions"] else 0
                    ),
                    "users": (
                        (this_week_data["users"] - prev_week_data["users"]) / prev_week_data["users"] * 100
                        if prev_week_data["users"] else 0
                    ),
                    "engaged_sessions": (
                        (this_week_data["engaged_sessions"] - prev_week_data["engaged_sessions"])
                        / prev_week_data["engaged_sessions"] * 100 if prev_week_data["engaged_sessions"] else 0
                    ),
                }

                # Write channel name
                worksheet.merge_range(row, 0, row + 2, 0, channel, header_format)

                # Write "This Week" data
                worksheet.write(row, 1, "This Week", number_format)
                worksheet.write(row, 2, this_week_data["users"], number_format)
                worksheet.write(row, 3, this_week_data["sessions"], number_format)
                worksheet.write(row, 4, this_week_data["engaged_sessions"], number_format)
                row += 1

                # Write "Prev Week" data
                worksheet.write(row, 1, "Prev Week", number_format)
                worksheet.write(row, 2, prev_week_data["users"], number_format)
                worksheet.write(row, 3, prev_week_data["sessions"], number_format)
                worksheet.write(row, 4, prev_week_data["engaged_sessions"], number_format)
                row += 1

                # Write % Change
                worksheet.write(row, 1, "% Change", number_format)
                worksheet.write(row, 2, f"{change_data['users']:.2f}%", number_format)
                worksheet.write(row, 3, f"{change_data['sessions']:.2f}%", number_format)
                worksheet.write(row, 4, f"{change_data['engaged_sessions']:.2f}%", number_format)
                row += 1

            # Save and close the workbook
            workbook.close()
            output.seek(0)

            # Provide download link
            st.download_button(
                label="Download Excel File",
                data=output,
                file_name="Weekly_Comparison_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Invalid key file or Property ID: {e}")

import streamlit as st
import xlsxwriter
from google.oauth2.service_account import Credentials
from google.analytics.data_v1beta import BetaAnalyticsDataClient
from google.analytics.data_v1beta.types import DateRange, Metric, Dimension, Filter, FilterExpression
from datetime import datetime, timedelta
import tempfile

# Streamlit interface for user input
st.title("GA4 Data Report Generator")
st.sidebar.header("Input Parameters")

# User inputs for start and end month/year
start_date = st.sidebar.date_input("Start Date", value=datetime.now().replace(day=1))
end_date = st.sidebar.date_input("End Date", value=datetime.now())

# Validate date range
if start_date > end_date:
    st.error("Start Date must be before End Date.")
    st.stop()

# User input for GA4 property ID
property_id = st.sidebar.text_input("Enter GA4 Property ID", placeholder="Enter your GA4 Property ID")

# Upload JSON key file
uploaded_file = st.sidebar.file_uploader("Upload JSON Key File", type=["json"])

# Define all available metrics
all_available_metrics = {
    "Sessions": "sessions",
    "Users": "totalUsers",
    "Engagement Rate": "engagementRate",
    "Bounce Rate": "bounceRate",
    "Transactions": "transactions",
    "Add to Carts": "addToCarts",
    "Revenue": "purchaseRevenue",
    "Pageviews": "screenPageViews",
    "Conversions": "conversions",
    "Engaged Sessions":"engagedSessions",
}

# User selects metrics dynamically
selected_metrics = st.sidebar.multiselect(
    "Select Metrics to Include",
    options=list(all_available_metrics.keys()),
    default=["Sessions", "Users"]
)

if not selected_metrics:
    st.error("Please select at least one metric.")
    st.stop()

# Run button
run_report = st.sidebar.button("Run Report")

if run_report:
    if not uploaded_file or not property_id:
        st.warning("Please provide all required inputs.")
        st.stop()

    # Save uploaded JSON key file temporarily
    with tempfile.NamedTemporaryFile(delete=False, suffix=".json") as temp_file:
        temp_file.write(uploaded_file.getvalue())
        temp_file_path = temp_file.name

    # Initialize the GA4 Data API client
    try:
        credentials = Credentials.from_service_account_file(temp_file_path)
        client = BetaAnalyticsDataClient(credentials=credentials)
    except Exception as e:
        st.error(f"Error initializing GA4 client: {e}")
        st.stop()

    # Generate a list of months between the selected dates
    def generate_months(start_date, end_date):
        current = start_date
        months = []
        while current <= end_date:
            month_start = current.replace(day=1)
            next_month = (month_start + timedelta(days=31)).replace(day=1)
            month_end = (next_month - timedelta(days=1))
            months.append({
                "month": month_start.strftime("%b"),
                "start_date": month_start.strftime("%Y-%m-%d"),
                "end_date": min(month_end, end_date).strftime("%Y-%m-%d"),
            })
            current = next_month
        return months

    months = generate_months(start_date, end_date)

    # Prepare request metrics dynamically based on user selection
    request_metrics = [Metric(name=all_available_metrics[metric]) for metric in selected_metrics]
    headers = ['Month'] + selected_metrics
    all_data = []

   
    # Create request with filter for 'Organic Search' traffic
    request = {
        "property": f"properties/{property_id}",
        "metrics": request_metrics,
        "dimensions": [
            Dimension(name="sessionDefaultChannelGrouping"),
        ],
        "dimension_filter": FilterExpression(
            filter=Filter(
                field_name="sessionDefaultChannelGrouping",
                string_filter=Filter.StringFilter(
                    match_type=Filter.StringFilter.MatchType.EXACT,
                    value="Organic Search"
                )
            )
        )
    }


    # Fetch data for each month
    for month in months:
        try:
            request['date_ranges'] = [DateRange(start_date=month['start_date'], end_date=month['end_date'])]
            response = client.run_report(request)

            monthly_data = [month['month']]
            for metric_index in range(len(request_metrics)):
                metric_value = sum(
                    float(row.metric_values[metric_index].value) for row in response.rows
                )
                monthly_data.append(metric_value)

            all_data.append(monthly_data)

            st.success(f"Data for {month['month']} successfully fetched.")

        except Exception as e:
            st.error(f"Error fetching data for {month['month']}: {e}")

    # Generate Excel report
    output_file = "GA4_Report_Insights.xlsx"
    workbook = xlsxwriter.Workbook(output_file)

    # First sheet: GA4 Data (Table + Conditional Formatting + Insights)
    worksheet = workbook.add_worksheet("GA4 Data")
    header_format = workbook.add_format({'bold': True, 'bg_color': '#1F4E78', 'font_color': 'white', 'align': 'center', 'border': 1})
    content_format = workbook.add_format({'align': 'center', 'border': 1})
    percentage_format = workbook.add_format({'align': 'center', 'border': 1, 'num_format': '0.00%'})
    improvement_format = workbook.add_format({'font_color': '#006400', 'bold': True})  # Dark green for improvement
    drop_format = workbook.add_format({'font_color': 'red', 'bold': True})

    # Write headers
    for col_num, header in enumerate(headers):
        worksheet.write(0, col_num, header, header_format)

    # Write data
    for row_num, row_data in enumerate(all_data, start=1):
        for col_num, cell_data in enumerate(row_data):
            worksheet.write(row_num, col_num, cell_data, content_format)

    # Add % Difference row
    if len(all_data) >= 2:
        last_month_data = all_data[-1][1:]  # Exclude 'Month'
        second_last_month_data = all_data[-2][1:]  # Exclude 'Month'
        percentage_differences = [
            round((last - second_last) / second_last * 100, 2)*0.01 if second_last != 0 else 0
            for last, second_last in zip(last_month_data, second_last_month_data)
        ]

        percent_diff_row = ["% Difference"] + percentage_differences
        for col_num, cell_data in enumerate(percent_diff_row):
            worksheet.write(len(all_data) + 1, col_num, cell_data, content_format)

    # Conditional formatting
    for col_num in range(1, len(headers)):  # Skip the "Month" column
        col_letter = chr(65 + col_num)
        data_range = f"{col_letter}2:{col_letter}{len(all_data) + 1}"
        worksheet.conditional_format(
            data_range,
            {
                'type': '3_color_scale',
                'min_color': "#F8696B",  # Red (low values)
                'mid_color': "#FFEB84",  # Yellow (mid values)
                'max_color': "#63BE7B",  # Green (high values)
            }
        )

    # Add insights
    if len(all_data) >= 2:
        last_month = all_data[-1][0]
        second_last_month = all_data[-2][0]
        insights_row_start = len(all_data) + 3
        insights_col_start = len(headers)

        worksheet.write(insights_row_start - 1, insights_col_start, f"Highlights - {last_month} vs {second_last_month}", header_format)

        for i in range(1, len(headers)):  # Skip the 'Month' column
            metric_name = headers[i]
            previous_value = all_data[-2][i]
            current_value = all_data[-1][i]
            percentage_change = round((current_value - previous_value) / previous_value * 100, 2) if previous_value != 0 else 0
            change_type = "improvement" if percentage_change > 0 else "drop"
            percentage_text = f"{change_type} of {abs(percentage_change)}%"
            worksheet.write_rich_string(
                insights_row_start + i - 1, insights_col_start,
                "We observed a ",
                improvement_format if change_type == "improvement" else drop_format, percentage_text,
                f" in {metric_name} in {last_month} compared to {second_last_month}.", content_format
            )

    # Second sheet: Graphs
    chart_sheet = workbook.add_worksheet("Graphs")
    chart_row = 0
    chart_col = 0

    for metric_index, metric in enumerate(selected_metrics):
        chart = workbook.add_chart({'type': 'column'})

        chart.add_series({
            'name': f"{metric} vs. Month",
            'categories': f"='GA4 Data'!A2:A{len(all_data) + 1}",
            'values': f"='GA4 Data'!{chr(66 + metric_index)}2:{chr(66 + metric_index)}{len(all_data) + 1}",
            'fill': {'color': 'blue'},
        })

        chart.set_title({'name': f"{metric} per Month"})
        chart.set_x_axis({'name': 'Month'})
        chart.set_y_axis({'name': metric})
        chart.set_legend({'position': 'none'})

        chart_sheet.insert_chart(chart_row, chart_col, chart)
        chart_row += 15

    # Close the workbook
    workbook.close()

    st.success(f"Data and graphs successfully written to {output_file}.")
    st.download_button("Download Report", data=open(output_file, "rb"), file_name=output_file)

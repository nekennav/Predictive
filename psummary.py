import streamlit as st
import pandas as pd
from datetime import datetime, date
import os
from werkzeug.utils import secure_filename
import io
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment

# Set page configuration
st.set_page_config(page_title="Neeiil", page_icon="💎", layout="wide")

# Custom CSS for aesthetic computer-themed styling
st.markdown(
    """
    <style>
    /* Main app background with aesthetic computer wallpaper */
    .stApp {
        background-image: url('https://wallpaperbat.com/img/881836-aesthetic-computer-4k-wallpaper.jpg');
        background-size: cover;
        background-position: center;
        background-repeat: no-repeat;
        background-color: rgba(26, 26, 46, 0.9); /* Dark purple-blue semi-transparent fallback */
        color: #e0e0e0; /* Light gray for text readability */
        animation: panningBackground 20s linear infinite;
    }
    @keyframes panningBackground {
        0% { background-position: 50% 0%; }
        50% { background-position: 50% 100%; }
        100% { background-position: 50% 0%; }
    }
    /* Title styling with abstract-inspired colors and white stroke */
    h1 {
        color: #a6b1e1; /* Light purple for PREDICTIVE SUMMARY */
        font-family: 'Arial', sans-serif;
        text-align: center;
        text-shadow:
            -1px -1px 0 #ffffff,
             1px -1px 0 #ffffff,
            -1px 1px 0 #ffffff,
             1px 1px 0 #ffffff; /* White stroke effect */
        font-size: 2.5rem; /* Base font size */
    }
<   /* File uploader styling */
    .stFileUploader {
        background-color: rgba(162, 138, 255, 0.1); /* Light purple tint */
        border: 2px solid #a28aff; /* Solid purple border */
        border-radius: 10px;
        padding: 10px;
        margin-top: 20px; /* Slightly lower positioning */
    }
    /* File uploader label styling */
    .stFileUploader label {
        color: #e0e0e0; /* Light gray for file uploader label */
        font-size: 1.1rem; /* Base font size */
    }
    /* File uploader button text ("Browse files") */
    .stFileUploader input[type="file"]::file-selector-button {
        color: #000000; /* Black color for "Browse files" text */
        font-size: 1rem; /* Match button font size */
    }
    /* Button styling */
    .stButton > button {
        background-color: #a28aff; /* Vibrant purple */
        color: #e0e0e0; /* Light gray text */
        border-radius: 8px;
        padding: 10px 20px;
        font-weight: bold;
        transition: background-color 0.3s;
        border: 1px solid #dcd6ff; /* Light purple border */
        font-size: 1rem; /* Base font size */
    }
    .stButton > button:hover {
        background-color: #8b74e6; /* Darker purple on hover */
        border-color: #a28aff; /* Purple on hover */
    }
    /* Download button text ("Download data as Excel") */
    .stDownloadButton > button {
        color: #000000; /* Black color for download button text */
    }
    /* Text area and dataframe styling */
    .stTextArea textarea, .stDataFrame {
        background-color: rgba(162, 138, 255, 0.15); /* Subtle purple tint */
        color: #e0e0e0; /* Light gray text */
        border: 1px solid #a28aff; /* Solid purple border */
        border-radius: 8px;
        font-size: 0.9rem; /* Base font size */
    }
    /* Alert styling */
    .stAlert {
        background-color: rgba(220, 214, 255, 0.2); /* Light purple tint */
        color: #e0e0e0; /* Light gray text */
        border-radius: 8px;
        font-size: 0.9rem; /* Base font size */
    }
    /* Table text */
    .stText, table {
        color: #e0e0e0; /* Light gray for table text */
        font-family: 'Arial', sans-serif;
        font-size: 0.9rem; /* Base font size */
    }
    /* Date input label styling */
    .stDateInput label {
        color: #ffffff; /* White color for date input labels */
        font-size: 1.1rem; /* Match file uploader label size */
    }
    /* Date input text color */
    .stDateInput > div > div > input {
        color: #000000 !important;
    }
    /* Selectbox styling for dropdown */
    div[data-baseweb="select"] {
        background-color: rgba(162, 138, 255, 0.15);
        border: 1px solid #a28aff;
        border-radius: 8px;
    }
    div[data-baseweb="select"] input {
        background-color: rgba(162, 138, 255, 0.15) !important;
        color: #000000 !important;
    }
    div[data-baseweb="select"] div[role="listbox"] {
        background-color: rgba(162, 138, 255, 0.15);
        color: #000000;
        border: 1px solid #a28aff;
        border-radius: 8px;
    }
    /* Magnifying glass icon for multiselect */
    div[data-baseweb="select"] > div:first-child::before {
        content: '🔍';
        position: absolute;
        left: 10px;
        top: 50%;
        transform: translateY(-50%);
        pointer-events: none;
        z-index: 1;
        font-size: 1.2rem;
    }
    div[data-baseweb="select"] > div:first-child > div {
        padding-left: 30px !important; /* Space for icon */
    }
    /* Responsive design for mobile */
    @media (max-width: 600px) {
        h1 {
            font-size: 1.8rem; /* Smaller title on mobile */
        }
        .stFileUploader {
            margin-top: 10px; /* Slightly less margin on mobile */
        }
        .stFileUploader label {
            font-size: 0.9rem; /* Smaller label on mobile */
        }
        .stButton > button {
            padding: 8px 16px;
            font-size: 0.85rem; /* Smaller button text */
        }
        .stTextArea textarea, .stDataFrame, .stAlert, .stText, table {
            font-size: 0.8rem; /* Smaller text on mobile */
        }
        .stDateInput label {
            font-size: 0.9rem; /* Smaller date input label on mobile */
        }
        .stFileUploader input[type="file"]::file-selector-button {
            font-size: 0.85rem; /* Smaller file uploader button text on mobile */
        }
        .stDownloadButton > button {
            font-size: 0.85rem; /* Smaller download button text on mobile */
        }
    }
    </style>
    """,
    unsafe_allow_html=True
)

# Initialize session state for show_data if not already set
if 'show_data' not in st.session_state:
    st.session_state.show_data = False

if 'last_uploaded_file' not in st.session_state:
    st.session_state.last_uploaded_file = None

# Check and update session schema
def init_session(columns):
    normalized_columns = [col.strip() for col in columns]
    if 'CAMPAIGN' in normalized_columns:
        normalized_columns[normalized_columns.index('CAMPAIGN')] = 'Campaign'
    elif 'campaign' in normalized_columns:
        normalized_columns[normalized_columns.index('campaign')] = 'Campaign'
    elif 'Campaign Name' in normalized_columns:
        normalized_columns[normalized_columns.index('Campaign Name')] = 'Campaign'
    elif 'CAMPAIGN NAME' in normalized_columns:
        normalized_columns[normalized_columns.index('CAMPAIGN NAME')] = 'Campaign'
  
    # Expected columns are those from the uploaded file plus 'filename', 'upload_date'
    expected_columns = ['filename'] + normalized_columns + ['upload_date']
  
    if 'df' not in st.session_state:
        st.session_state.df = pd.DataFrame(columns=expected_columns)
    else:
        #

        current_columns = set(st.session_state.df.columns)
      
        if current_columns != set(expected_columns):
            # If schema doesn't match, recreate (reset)
            st.warning(f"Schema mismatch. Current columns: {list(current_columns)}, Expected: {expected_columns}. Resetting stored data.")
            st.session_state.df = pd.DataFrame(columns=expected_columns)

# Save data from file to session without transformations
def save_file_to_session(uploaded_file):
    filename = secure_filename(uploaded_file.name)
    ext = os.path.splitext(filename)[1].lower()
  
    try:
        if ext == '.csv':
            df = pd.read_csv(uploaded_file, dtype=str) # Read as strings to preserve data
        elif ext in ['.xls', '.xlsx']:
            df = pd.read_excel(uploaded_file, dtype=str) # Read as strings to preserve data
        else:
            raise ValueError("Unsupported file type. Please upload CSV, XLS, or XLSX.")
      
        # Normalize column names (strip spaces, keep case)
        df.columns = df.columns.str.strip()
      
        # Normalize 'Campaign' column name variations (including all caps)
        campaign_variants = ['CAMPAIGN', 'campaign', 'Campaign Name', 'CAMPAIGN NAME']
        for variant in campaign_variants:
            if variant in df.columns:
                df = df.rename(columns={variant: 'Campaign'})
                st.info(f"Renamed column '{variant}' to 'Campaign' for consistency.")
                break
      
        # Validate required columns
        required_cols = ['Collector Name', 'Campaign']
        if not all(col in df.columns for col in required_cols):
            raise ValueError(f"File must contain required columns: {', '.join(required_cols)}. Note: 'CAMPAIGN' should be renamed to 'Campaign'.")
      
        # Check for valid campaign values
        valid_campaigns = {'SBC CURING B2', 'SBC CURING B4', 'SBC RECOVERY', 'SBF RECOVERY'}
        if not df['Campaign'].isin(valid_campaigns | {''}).all():
            invalid_campaigns = df['Campaign'][~df['Campaign'].isin(valid_campaigns | {''})].unique()
            st.warning(f"Invalid campaign names found: {', '.join(invalid_campaigns)}. Expected: {', '.join(valid_campaigns)}")
      
        # Normalize Date column to MM/DD/YYYY with leading zeros
        if 'Date' in df.columns:
            def normalize_date(x):
                if pd.notnull(x) and x.strip():
                    try:
                        dt = pd.to_datetime(x, errors='coerce', dayfirst=False, yearfirst=False)
                        if pd.notnull(dt):
                            return dt.strftime('%m/%d/%Y')
                        return x.strip()
                    except:
                        return x.strip()
                return x
            df['Date'] = df['Date'].apply(normalize_date)
      
        # Initialize session with normalized columns from the file
        init_session(df.columns)
      
        upload_date = datetime.now().strftime('%Y-%m-%d')
      
        # Add metadata columns
        df['filename'] = filename
        df['upload_date'] = upload_date
      
        # Append to session df (columns already aligned by init)
        if st.session_state.df.empty:
            st.session_state.df = df
        else:
            st.session_state.df = pd.concat([st.session_state.df, df], ignore_index=True)
      
        return filename
    except Exception as e:
        raise Exception(f"Failed to process file: {str(e)}")

# Get all records as DataFrame
def get_all_records():
    try:
        if 'df' in st.session_state and not st.session_state.df.empty:
            df = st.session_state.df.copy()
          
            # Keep all data as strings, no datetime conversion
            display_df = df.copy()
            # Reorder columns to match uploaded file order, with filename first
            data_cols = [col for col in df.columns if col not in ['filename', 'upload_date']]
            cols = ['filename'] + data_cols + ['upload_date']
            display_df = display_df[cols]
            return display_df, df
        else:
            return pd.DataFrame(), pd.DataFrame()
    except Exception as e:
        st.error(f"Error retrieving records: {str(e)}")
        return pd.DataFrame(), pd.DataFrame()

# Filter records by date range (using Date column in MM/DD/YYYY format)
def filter_records_by_date(start_date, end_date):
    try:
        # Convert input dates to MM/DD/YYYY for string comparison
        start_date_str = start_date.strftime('%m/%d/%Y')
        end_date_str = end_date.strftime('%m/%d/%Y')
      
        if 'df' not in st.session_state or st.session_state.df.empty:
            st.warning("No data stored in session.")
            return pd.DataFrame(), pd.DataFrame()
      
        df = st.session_state.df.copy()
      
        if 'Date' not in df.columns:
            st.warning("No 'Date' column found. Returning all data.")
            filtered_df = df
        else:
            # Use string comparison
            if start_date == end_date:
                filtered_df = df[df['Date'] == start_date_str]
            else:
                filtered_df = df[df['Date'].between(start_date_str, end_date_str)]
      
        if filtered_df.empty:
            st.warning(f"No records found for date range {start_date_str} to {end_date_str}. Check if dates in the 'Date' column are in MM/DD/YYYY format.")
      
        # Keep all data as strings, no datetime conversion
        display_df = filtered_df.copy()
        # Reorder columns to match uploaded file order, with filename first
        data_cols = [col for col in filtered_df.columns if col not in ['filename', 'upload_date']]
        cols = ['filename'] + data_cols + ['upload_date']
        display_df = display_df[cols]
        return display_df, filtered_df
    except Exception as e:
        st.error(f"Error filtering records: {str(e)}")
        return pd.DataFrame(), pd.DataFrame()

# Helper function to convert [H]:MM:SS time string to seconds
def time_to_seconds(t):
    if pd.isna(t) or str(t).strip() == '':
        return 0.0 # Return 0 as float to avoid integer conversion issues
    try:
        parts = str(t).strip().split(':')
        if len(parts) == 3: # [H]:MM:SS
            hours = float(parts[0]) if parts[0] and parts[0].replace('-', '').isdigit() else 0.0
            minutes = float(parts[1]) if parts[1].isdigit() else 0.0
            seconds = float(parts[2]) if parts[2].isdigit() else 0.0
            return hours * 3600 + minutes * 60 + seconds
        elif len(parts) == 2: # MM:SS (treat as 0 hours)
            minutes = float(parts[0]) if parts[0].isdigit() else 0.0
            seconds = float(parts[1]) if parts[1].isdigit() else 0.0
            return minutes * 60 + seconds
        else:
            return 0.0 # Invalid format
    except (ValueError, TypeError):
        return 0.0 # Return 0 for any parsing errors

# Helper function to convert seconds back to [H]:MM:SS string
def seconds_to_time(seconds):
    if pd.isna(seconds) or seconds == 0:
        return '0:00:00'
    hours = int(float(seconds) // 3600) # Convert to float first, then to int
    minutes = int((float(seconds) % 3600) // 60)
    secs = int(float(seconds) % 60)
    return f"{hours}:{minutes:02d}:{secs:02d}"

# Helper function to convert seconds to MM:SS for averages
def seconds_to_avg_time(seconds):
    if pd.isna(seconds) or seconds == 0 or not isinstance(seconds, (int, float)) or seconds == float('inf'):
        return '00:00'
    minutes = int(float(seconds) // 60) # Convert to float first, then to int
    secs = int(float(seconds) % 60)
    return f"{minutes:02d}:{secs:02d}"

# Calculate totals and averages by agent and campaign
def calculate_agent_totals(df):
    try:
        # Define the columns we want to aggregate
        metric_columns = [
            'Total Calls', 'Spent Time', 'Talk Time', 'Wait Time',
            'Write Time', 'Pause Time'
        ]
        # Ensure only available columns are used
        available_metrics = [col for col in metric_columns if col in df.columns]
      
        if 'Collector Name' not in df.columns or not available_metrics:
            missing_cols = [col for col in ['Collector Name'] + metric_columns if col not in df.columns]
            st.warning(f"Required columns missing: {', '.join(missing_cols)}. Cannot calculate totals.")
            return pd.DataFrame()
      
        # Check for Campaign column
        if 'Campaign' not in df.columns:
            st.warning("No 'Campaign' column found in data. Totals will be grouped by Collector Name only.")
            group_cols = ['Collector Name']
        else:
            group_cols = ['Collector Name', 'Campaign']
      
        # Create a copy to avoid modifying the input DataFrame
        df = df.copy()
      
        # Convert Total Calls to numeric
        if 'Total Calls' in available_metrics:
            df['Total Calls'] = pd.to_numeric(df['Total Calls'], errors='coerce').fillna(0.0)
      
        # Convert time columns to seconds
        time_metrics = [col for col in available_metrics if col != 'Total Calls']
        for col in time_metrics:
            df[f'{col}_seconds'] = df[col].apply(time_to_seconds).fillna(0.0)
      
        # Group by Collector Name and Campaign (if available)
        sum_cols = ['Total Calls'] + [f'{col}_seconds' for col in time_metrics]
        totals_df = df.groupby(group_cols)[sum_cols].sum().reset_index()
      
        # Convert summed seconds back to [H]:MM:SS for display
        for col in time_metrics:
            totals_df[col] = totals_df[f'{col}_seconds'].apply(seconds_to_time)
            totals_df = totals_df.drop(columns=[f'{col}_seconds']) # Clean up temp column
      
        # Calculate averages (in seconds per call, then to MM:SS)
        if 'Total Calls' in available_metrics and totals_df['Total Calls'].sum() > 0:
            for col in ['Talk Time', 'Wait Time', 'Write Time']:
                if col in time_metrics:
                    # Avoid division by zero
                    totals_df[f'AVG {col}'] = totals_df.apply(
                        lambda row: seconds_to_avg_time(
                            time_to_seconds(row[col]) / row['Total Calls']
                            if row['Total Calls'] > 0 else 0.0
                        ), axis=1
                    )
      
        # Reorder columns for display
        display_columns = ['Collector Name']
        if 'Campaign' in df.columns:
            display_columns.append('Campaign')
        if 'Total Calls' in available_metrics:
            display_columns.append('Total Calls')
        for col in ['Spent Time', 'Talk Time', 'Wait Time', 'Write Time', 'Pause Time']:
            if col in available_metrics:
                display_columns.append(col)
            if col in ['Talk Time', 'Wait Time', 'Write Time'] and col in available_metrics:
                display_columns.append(f'AVG {col}')
      
        return totals_df[display_columns]
    except Exception as e:
        st.error(f"Error calculating totals: {str(e)}")
        return pd.DataFrame()

# Calculate daily averages per agent across all days
def calculate_daily_agent_averages(df):
    try:
        # Define the columns we want to aggregate
        metric_columns = [
            'Total Calls', 'Spent Time', 'Talk Time', 'Wait Time',
            'Write Time', 'Pause Time'
        ]
        # Ensure only available columns are used
        available_metrics = [col for col in metric_columns if col in df.columns]
      
        if 'Collector Name' not in df.columns or 'Date' not in df.columns or not available_metrics:
            missing_cols = [col for col in ['Collector Name', 'Date'] + metric_columns if col not in df.columns]
            st.warning(f"Required columns missing: {', '.join(missing_cols)}. Cannot calculate daily averages.")
            return pd.DataFrame()
      
        # Check for Campaign column
        if 'Campaign' not in df.columns:
            st.warning("No 'Campaign' column found in data. Averages will be grouped by Collector Name only.")
            group_cols = ['Collector Name']
            daily_group_cols = ['Collector Name', 'Date']
        else:
            group_cols = ['Collector Name', 'Campaign']
            daily_group_cols = ['Collector Name', 'Date', 'Campaign']
      
        # Create a copy to avoid modifying the input DataFrame
        df = df.copy()
      
        # Convert Total Calls to numeric
        if 'Total Calls' in available_metrics:
            df['Total Calls'] = pd.to_numeric(df['Total Calls'], errors='coerce').fillna(0.0)
      
        # Convert time columns to seconds
        time_metrics = [col for col in available_metrics if col != 'Total Calls']
        for col in time_metrics:
            df[f'{col}_seconds'] = df[col].apply(time_to_seconds).fillna(0.0)
      
        # Step 1: Group by Collector Name, Date, and Campaign (if available) to get daily metrics
        daily_mean_cols = ['Total Calls'] + [f'{col}_seconds' for col in time_metrics]
        daily_averages = df.groupby(daily_group_cols)[daily_mean_cols].mean().reset_index()
      
        # Step 2: Group by Collector Name and Campaign (if available) to average daily metrics
        mean_cols = ['Total Calls'] + [f'{col}_seconds' for col in time_metrics]
        averages_df = daily_averages.groupby(group_cols)[mean_cols].mean().reset_index()
      
        # Convert averaged seconds back to [H]:MM:SS for display
        for col in time_metrics:
            averages_df[col] = averages_df[f'{col}_seconds'].apply(seconds_to_time)
            averages_df = averages_df.drop(columns=[f'{col}_seconds']) # Clean up temp column
      
        # Calculate per-call averages (in seconds per call, then to MM:SS)
        if 'Total Calls' in available_metrics and averages_df['Total Calls'].sum() > 0:
            for col in ['Talk Time', 'Wait Time', 'Write Time']:
                if col in time_metrics:
                    # Avoid division by zero
                    averages_df[f'AVG {col}'] = averages_df.apply(
                        lambda row: seconds_to_avg_time(
                            time_to_seconds(row[col]) / row['Total Calls']
                            if row['Total Calls'] > 0 else 0.0
                        ), axis=1
                    )
      
        # Reorder columns for display
        display_columns = ['Collector Name']
        if 'Campaign' in df.columns:
            display_columns.append('Campaign')
        if 'Total Calls' in available_metrics:
            display_columns.append('Total Calls')
        for col in ['Spent Time', 'Talk Time', 'Wait Time', 'Write Time', 'Pause Time']:
            if col in available_metrics:
                display_columns.append(col)
            if col in ['Talk Time', 'Wait Time', 'Write Time'] and col in available_metrics:
                display_columns.append(f'AVG {col}')
      
        return averages_df[display_columns]
    except Exception as e:
        st.error(f"Error calculating daily agent averages: {str(e)}")
        return pd.DataFrame()

# Calculate averages by campaign
def calculate_campaign_averages(df):
    try:
        # Define the columns we want to aggregate
        metric_columns = [
            'Total Calls', 'Spent Time', 'Talk Time', 'Wait Time',
            'Write Time', 'Pause Time'
        ]
        # Ensure only available columns are used
        available_metrics = [col for col in metric_columns if col in df.columns]
      
        if 'Campaign' not in df.columns or not available_metrics:
            missing_cols = [col for col in ['Campaign'] + metric_columns if col not in df.columns]
            st.warning(f"Required columns missing for campaign averages: {', '.join(missing_cols)}. Please ensure uploaded files include a 'Campaign' column with values like 'SBC CURING B2', 'SBC RECOVERY', etc.")
            return pd.DataFrame()
      
        # Create a copy to avoid modifying the input DataFrame
        df = df.copy()
      
        # Warn if all Campaign values are empty or null
        if df['Campaign'].isna().all() or (df['Campaign'] == '').all():
            st.warning("All 'Campaign' values are empty or null. Please ensure the 'Campaign' column contains valid values like 'SBC CURING B2', 'SBC RECOVERY', etc.")
            return pd.DataFrame()
      
        # Convert Total Calls to numeric
        if 'Total Calls' in available_metrics:
            df['Total Calls'] = pd.to_numeric(df['Total Calls'], errors='coerce').fillna(0.0)
      
        # Convert time columns to seconds
        time_metrics = [col for col in available_metrics if col != 'Total Calls']
        for col in time_metrics:
            df[f'{col}_seconds'] = df[col].apply(time_to_seconds).fillna(0.0)
      
        # Group by Campaign and calculate mean for the metrics
        mean_cols = ['Total Calls'] + [f'{col}_seconds' for col in time_metrics]
        averages_df = df.groupby('Campaign')[mean_cols].mean().reset_index()
      
        # Convert averaged seconds back to [H]:MM:SS for display
        for col in time_metrics:
            averages_df[col] = averages_df[f'{col}_seconds'].apply(seconds_to_time)
            averages_df = averages_df.drop(columns=[f'{col}_seconds']) # Clean up temp column
      
        # Calculate per-call averages (in seconds per call, then to MM:SS)
        if 'Total Calls' in available_metrics and averages_df['Total Calls'].sum() > 0:
            for col in ['Talk Time', 'Wait Time', 'Write Time']:
                if col in time_metrics:
                    # Avoid division by zero
                    averages_df[f'AVG {col}'] = averages_df.apply(
                        lambda row: seconds_to_avg_time(
                            time_to_seconds(row[col]) / row['Total Calls']
                            if row['Total Calls'] > 0 else 0.0
                        ), axis=1
                    )
      
        # Reorder columns for display
        display_columns = ['Campaign']
        if 'Total Calls' in available_metrics:
            display_columns.append('Total Calls')
        for col in ['Spent Time', 'Talk Time', 'Wait Time', 'Write Time', 'Pause Time']:
            if col in available_metrics:
                display_columns.append(col)
            if col in ['Talk Time', 'Wait Time', 'Write Time'] and col in available_metrics:
                display_columns.append(f'AVG {col}')
      
        return averages_df[display_columns]
    except Exception as e:
        st.error(f"Error calculating campaign averages: {str(e)}")
        return pd.DataFrame()

# Streamlit app
st.title("PREDICTIVE SUMMARY")

# File upload section
st.header("Upload a File")
uploaded_file = st.file_uploader("Choose a file", type=['csv', 'xls', 'xlsx'])

if uploaded_file is not None:
    # Show preview of the uploaded file
    st.subheader("File Preview")
    try:
        ext = os.path.splitext(uploaded_file.name)[1].lower()
        if ext == '.csv':
            preview_df = pd.read_csv(uploaded_file, dtype=str)
        elif ext in ['.xls', '.xlsx']:
            preview_df = pd.read_excel(uploaded_file, dtype=str)
        preview_df.columns = preview_df.columns.str.strip()
        # Normalize Date column for preview
        if 'Date' in preview_df.columns:
            def normalize_date(x):
                if pd.notnull(x) and x.strip():
                    try:
                        dt = pd.to_datetime(x, errors='coerce', dayfirst=False, yearfirst=False)
                        if pd.notnull(dt):
                            return dt.strftime('%m/%d/%Y')
                        return x.strip()
                    except:
                        return x.strip()
                return x
            preview_df['Date'] = preview_df['Date'].apply(normalize_date)
        st.dataframe(preview_df, use_container_width=True, height=300)
    except Exception as e:
        st.error(f"Error previewing file: {str(e)}")

    # Automatic upload
    current_file_name = uploaded_file.name
    if st.session_state.last_uploaded_file != current_file_name:
        try:
            filename = save_file_to_session(uploaded_file)
            st.session_state.last_uploaded_file = current_file_name
            st.markdown(f'<p style="color:white;">Data from \'{filename}\' uploaded and stored successfully!</p>', unsafe_allow_html=True)
        except Exception as e:
            st.error(f"Error: {str(e)}")

# Date range filter section
st.header("Filter Data by Date Range")
# Determine default start and end dates from the session DataFrame
default_start_date = datetime.today().date()
default_end_date = datetime.today().date()
if 'df' in st.session_state and not st.session_state.df.empty and 'Date' in st.session_state.df.columns:
    try:
        # Filter out invalid or empty dates and convert to datetime
        valid_dates = pd.to_datetime(st.session_state.df['Date'], format='%m/%d/%Y', errors='coerce')
        valid_dates = valid_dates.dropna()
        if not valid_dates.empty:
            default_start_date = valid_dates.min().date()
            default_end_date = valid_dates.max().date()
    except Exception as e:
        st.warning(f"Error determining date range: {str(e)}. Using today's date as default.")

col1, col2 = st.columns(2)
with col1:
    start_date = st.date_input("Start Date", value=default_start_date)
with col2:
    end_date = st.date_input("End Date", value=default_end_date)

# Show All button
if st.button("Show All"):
    st.session_state.show_data = True

# Display data sections only if show_data is True
if st.session_state.get('show_data', False):
    # TOTAL section
    st.header("TOTAL PER AGENT")
    if start_date <= end_date:
        _, raw_df = filter_records_by_date(start_date, end_date)
        if not raw_df.empty:
            totals_df = calculate_agent_totals(raw_df)
            if not totals_df.empty:
                all_agents = sorted(totals_df['Collector Name'].unique())
                st.markdown('<p style="color:white;">Search for Agent (Total)</p>', unsafe_allow_html=True)
                search_agent_total = st.multiselect("", options=all_agents, placeholder="Select one or more agents", key="search_total")
                if search_agent_total:
                    filtered_totals = totals_df[totals_df['Collector Name'].isin(search_agent_total)]
                    st.dataframe(filtered_totals, use_container_width=True)
                else:
                    st.write("Select one or more agents to view details.")
            else:
                st.warning("No data available for total calculations or required columns are missing.")
        else:
            st.warning("No data found for the selected date range.")
    else:
        st.error("Start date must be before or equal to end date.")
        raw_df = pd.DataFrame()
    
    # AVERAGE DAILY PER AGENT section
    st.header("AVERAGE DAILY PER AGENT")
    if not raw_df.empty:
        daily_averages_df = calculate_daily_agent_averages(raw_df)
        if not daily_averages_df.empty:
            all_agents_daily = sorted(daily_averages_df['Collector Name'].unique())
            st.markdown('<p style="color:white;">Search for Agent (Daily Average)</p>', unsafe_allow_html=True)
            search_agent_daily = st.multiselect("", options=all_agents_daily, placeholder="Select one or more agents", key="search_daily")
            if search_agent_daily:
                filtered_daily = daily_averages_df[daily_averages_df['Collector Name'].isin(search_agent_daily)]
                st.dataframe(filtered_daily, use_container_width=True)
            else:
                st.write("Select one or more agents to view details.")
        else:
            st.warning("No data available for daily agent average calculations or required columns (e.g., 'Collector Name', 'Date') are missing.")
    else:
        st.warning("No data to display daily agent averages.")
    
    # AVERAGE PER CAMPAIGN section
    st.header("AVERAGE PER CAMPAIGN")
    if not raw_df.empty:
        campaign_averages_df = calculate_campaign_averages(raw_df)
        if not campaign_averages_df.empty:
            st.dataframe(campaign_averages_df, use_container_width=True)
        else:
            st.warning("No data available for campaign average calculations or required columns are missing. Please ensure uploaded files include a 'Campaign' column with valid values like 'SBC CURING B2', 'SBC RECOVERY', etc.")
    else:
        st.warning("No data to display campaign averages.")
    
    # Single download button for Excel with three sheets
    if not raw_df.empty:
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # Write Total per Agent
            if 'filtered_totals' in locals() and not filtered_totals.empty:
                filtered_totals.to_excel(writer, index=False, sheet_name='Total per Agent')
                worksheet = writer.sheets['Total per Agent']
                for col_idx in range(1, len(filtered_totals.columns) + 1):
                    col_letter = get_column_letter(col_idx)
                    for row in range(1, len(filtered_totals) + 2):
                        cell = worksheet[f"{col_letter}{row}"]
                        cell.alignment = Alignment(horizontal='left')
            elif not totals_df.empty:
                totals_df.to_excel(writer, index=False, sheet_name='Total per Agent')
                worksheet = writer.sheets['Total per Agent']
                for col_idx in range(1, len(totals_df.columns) + 1):
                    col_letter = get_column_letter(col_idx)
                    for row in range(1, len(totals_df) + 2):
                        cell = worksheet[f"{col_letter}{row}"]
                        cell.alignment = Alignment(horizontal='left')
            
            # Write Average Daily per Agent
            if 'filtered_daily' in locals() and not filtered_daily.empty:
                filtered_daily.to_excel(writer, index=False, sheet_name='Average Daily per Agent')
                worksheet = writer.sheets['Average Daily per Agent']
                for col_idx in range(1, len(filtered_daily.columns) + 1):
                    col_letter = get_column_letter(col_idx)
                    for row in range(1, len(filtered_daily) + 2):
                        cell = worksheet[f"{col_letter}{row}"]
                        cell.alignment = Alignment(horizontal='left')
            elif not daily_averages_df.empty:
                daily_averages_df.to_excel(writer, index=False, sheet_name='Average Daily per Agent')
                worksheet = writer.sheets['Average Daily per Agent']
                for col_idx in range(1, len(daily_averages_df.columns) + 1):
                    col_letter = get_column_letter(col_idx)
                    for row in range(1, len(daily_averages_df) + 2):
                        cell = worksheet[f"{col_letter}{row}"]
                        cell.alignment = Alignment(horizontal='left')
            
            # Write Average per Campaign
            if not campaign_averages_df.empty:
                campaign_averages_df.to_excel(writer, index=False, sheet_name='Average per Campaign')
                worksheet = writer.sheets['Average per Campaign']
                for col_idx in range(1, len(campaign_averages_df.columns) + 1):
                    col_letter = get_column_letter(col_idx)
                    for row in range(1, len(campaign_averages_df) + 2):
                        cell = worksheet[f"{col_letter}{row}"]
                        cell.alignment = Alignment(horizontal='left')
        
        output.seek(0)
        st.download_button(
            label="Download data as Excel",
            data=output,
            file_name=f"summary_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

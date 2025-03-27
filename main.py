import win32com.client
from datetime import datetime, timedelta
import pandas as pd
import os
import sys
import argparse

def parse_arguments():
    parser = argparse.ArgumentParser(description='Analyze working hours from Outlook calendar.')
    parser.add_argument('-e', '--export', action='store_true', help='Export results to Excel')
    parser.add_argument('-m', '--mode', type=int, choices=[1, 2, 3, 4],
                       help='Date range mode: 1=This week, 2=This month, 3=Custom date range, 4=Specific week')
    parser.add_argument('-y', '--year', type=int, help='Year for specific week mode')
    parser.add_argument('-w', '--week', type=int, help='Week number for specific week mode')
    parser.add_argument('-s', '--start', help='Start date for custom range (YYYY-MM-DD)')
    parser.add_argument('-t', '--end', help='End date for custom range (YYYY-MM-DD)')

    return parser.parse_args()

def connect_to_outlook():
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        calendar_folder = namespace.GetDefaultFolder(9)  # 9 = Calendar
        print(f"Connected to Outlook calendar: {calendar_folder.Name}")
        print(f"Total items in calendar: {calendar_folder.Items.Count}")
        return calendar_folder
    except Exception as e:
        print(f"Error connecting to Outlook: {e}")
        sys.exit(1)

def get_appointments_in_range(calendar_folder, start_date, end_date):
    # Format the date range for display
    print(f"\nFetching appointments from {start_date.strftime('%Y-%m-%d')} to {end_date.strftime('%Y-%m-%d')}...")

    try:
        # Get calendar items
        items = calendar_folder.Items
        print(f"Total calendar items before filtering: {items.Count}")

        # Important: Sort items by start date to optimize filtering
        items.Sort("[Start]")
        items.IncludeRecurrences = True

        # Format dates for the Outlook filter using proper formatting
        # Use the format that Outlook expects: "mm/dd/yyyy hh:mm AM/PM"
        start_str = start_date.strftime('%m/%d/%Y %I:%M %p')
        end_str = end_date.strftime('%m/%d/%Y %I:%M %p')

        # Use a restriction to filter by date range - using >= for start and <= for end
        restriction = f"[Start] >= '{start_str}' AND [End] <= '{end_str}'"
        print(f"Using filter: {restriction}")
        filtered_items = items.Restrict(restriction)
        print(f"Items after date restriction: {filtered_items.Count}")

        # Safety check to prevent processing too many items
        if filtered_items.Count > 10000:
            print(f"Warning: Very large number of items ({filtered_items.Count}). Limiting to 10000.")

        # DEBUG: Try to get at least one item to verify access
        if filtered_items.Count > 0:
            try:
                first_item = filtered_items.GetFirst()
                if first_item:
                    print(f"DEBUG: First filtered item subject: {first_item.Subject}")
                    print(f"DEBUG: First filtered item start: {first_item.Start}")
                else:
                    print("DEBUG: First item is None despite Count > 0")
            except Exception as e:
                print(f"DEBUG: Error accessing first filtered item: {e}")

        # Process the filtered items
        appointments = []
        processed_count = 0
        max_items = min(filtered_items.Count, 10000)  # Safety limit

        current_item = filtered_items.GetFirst()
        while current_item and processed_count < max_items:
            processed_count += 1
            if processed_count % 100 == 0:
                print(f"Processed {processed_count} items...")

            try:
                # Get the start time
                start_time = None
                try:
                    start_time = current_item.Start
                    if isinstance(start_time, str):
                        start_time = pd.to_datetime(start_time)
                except Exception as e:
                    print(f"DEBUG: Error getting start time: {e}")
                    current_item = filtered_items.GetNext()
                    continue

                # Convert to Python datetime if it's a COM datetime
                if hasattr(start_time, 'year'):  # Outlook datetime object
                    start_time = datetime(
                        year=start_time.year,
                        month=start_time.month,
                        day=start_time.day,
                        hour=start_time.hour,
                        minute=start_time.minute,
                        second=start_time.second
                    )

                # Double-check if in date range
                if not (start_date <= start_time <= end_date):  # Removed '+ timedelta(days=1)'
                    print(f"DEBUG: Item outside date range. Item date: {start_time}, Range: {start_date} to {end_date}")
                    current_item = filtered_items.GetNext()
                    continue

                # Extract end time
                end_time = None
                try:
                    end_time = current_item.End
                    if isinstance(end_time, str):
                        end_time = pd.to_datetime(end_time)

                    # Convert to Python datetime if it's a COM datetime
                    if hasattr(end_time, 'year'):
                        end_time = datetime(
                            year=end_time.year,
                            month=end_time.month,
                            day=end_time.day,
                            hour=end_time.hour,
                            minute=end_time.minute,
                            second=end_time.second
                        )
                except:
                    # If we can't get end time, estimate it
                    try:
                        end_time = start_time + timedelta(minutes=current_item.Duration)
                    except:
                        end_time = start_time + timedelta(hours=1)

                # Calculate duration in hours
                try:
                    duration = (end_time - start_time).total_seconds() / 3600
                except:
                    try:
                        duration = current_item.Duration / 60  # Minutes to hours
                    except:
                        duration = 1.0  # Default 1 hour

                # Get category (or set default)
                try:
                    category = current_item.Categories if current_item.Categories else "Uncategorized"
                except:
                    category = "Uncategorized"

                # Add to our list
                appointments.append({
                    'Subject': current_item.Subject,
                    'Start': start_time,
                    'End': end_time,
                    'Duration': duration,
                    'Category': category
                })
            except Exception as e:
                if processed_count < 20:
                    print(f"DEBUG: Error processing item {processed_count}: {e}")

            # Get the next item
            current_item = filtered_items.GetNext()

        print(f"\nTotal items processed: {processed_count}")
        print(f"Total appointments found in date range: {len(appointments)}")
        return appointments

    except Exception as e:
        print(f"Error getting appointments: {e}")
        return []

def analyze_work_hours(appointments, exclude_personal=True):
    if not appointments:
        return None

    # Convert to DataFrame for easy analysis
    df = pd.DataFrame(appointments)

    # Convert datetime objects if needed
    for col in ['Start', 'End']:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col])

    # Process categories (handle semicolon-separated lists)
    if 'Category' in df.columns:
        df['Category'] = df['Category'].astype(str).str.split(';')
        df = df.explode('Category').reset_index(drop=True)
        df['Category'] = df['Category'].str.strip()

    # Extract date information
    df['Date'] = df['Start'].dt.date
    df['Day'] = df['Start'].dt.day_name()
    df['Week'] = df['Start'].dt.strftime('%Y-W%U')  # Using %U instead of %W for week number (Sunday as the first day of the week)
    df['Month'] = df['Start'].dt.strftime('%Y-%m')

    # Store complete data before filtering out personal
    complete_df = df.copy()

    # Get personal hours info before filtering
    personal_hours = 0
    if 'Pessoal' in df['Category'].values:
        personal_hours = df[df['Category'] == 'Pessoal']['Duration'].sum()

    # Filter out personal events if requested
    if exclude_personal:
        work_df = df[df['Category'] != 'Pessoal'].copy()  # Create a copy to fix SettingWithCopyWarning
    else:
        work_df = df.copy()  # Create a copy to fix SettingWithCopyWarning

    # Total work hours
    total_work_hours = work_df['Duration'].sum()

    # Summary by category
    category_summary = work_df.groupby('Category', observed=True)['Duration'].agg(['sum', 'count']).reset_index()
    category_summary.columns = ['Category', 'Total Hours', 'Number of Events']

    # Add percentage column
    category_summary['Percentage'] = (category_summary['Total Hours'] / total_work_hours * 100).round(1)

    # Sort by total hours
    category_summary = category_summary.sort_values('Total Hours', ascending=False)

    # Define the correct order of days
    day_order = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']

    # Convert 'Day' to a categorical type with the specified order
    work_df['Day'] = pd.Categorical(work_df['Day'], categories=day_order, ordered=True)

    # Summary by day - add observed=True to fix FutureWarning
    daily_summary = work_df.groupby(['Date', 'Day'], observed=True)['Duration'].sum().reset_index()
    daily_summary = daily_summary.sort_values('Date')

    # Weekly summary
    weekly_summary = work_df.groupby('Week', observed=True)['Duration'].sum().reset_index()

    # Monthly summary
    monthly_summary = work_df.groupby('Month', observed=True)['Duration'].sum().reset_index()

    # Day of week summary - add observed=True to fix FutureWarning
    dow_summary = work_df.groupby('Day', observed=True)['Duration'].agg(['sum', 'mean']).reset_index()
    dow_summary.columns = ['Day', 'Total Hours', 'Average Hours']
    dow_summary['Day'] = pd.Categorical(dow_summary['Day'], categories=day_order, ordered=True)
    dow_summary = dow_summary.sort_values('Day')

    # Calculate daily average for working days (excluding weekends)
    working_days = work_df[~work_df['Day'].isin(['Saturday', 'Sunday'])]
    working_day_avg = working_days.groupby('Date', observed=True)['Duration'].sum().mean() if not working_days.empty else 0

    return {
        'total_hours': total_work_hours,
        'personal_hours': personal_hours,
        'daily_average': work_df.groupby('Date', observed=True)['Duration'].sum().mean(),
        'working_day_average': working_day_avg,
        'category_summary': category_summary,
        'daily_summary': daily_summary,
        'weekly_summary': weekly_summary,
        'monthly_summary': monthly_summary,
        'day_of_week_summary': dow_summary,
        'dataframe': work_df,
        'complete_dataframe': complete_df
    }

def get_week_dates(year, week_number):
    """
    Get the start and end dates for a specific ISO week number.
    ISO weeks start on Monday and end on Sunday.
    """
    try:
        # Use the ISO calendar
        import datetime as dt
        from datetime import datetime
        import calendar

        # Calculate the Monday of the specified week
        # Use isocalendar() to get the ISO year, week number, and weekday
        date = datetime.fromisocalendar(year, week_number, 1)  # 1 represents Monday

        start_date = date
        end_date = start_date + timedelta(days=6)

        # Convert to datetime objects with time
        start_datetime = datetime.combine(start_date, datetime.min.time())
        end_datetime = datetime.combine(end_date, datetime.max.time())

        return start_datetime, end_datetime

    except ValueError as e:
        print(f"Error calculating week dates: {e}. Please ensure the week number is valid for the given year.")
        return None, None  # Handle invalid week numbers
    except Exception as e:
        print(f"Error calculating week dates: {e}")
        return None, None

def main():
    try:
        # Parse command line arguments
        args = parse_arguments()

        calendar = connect_to_outlook()

        now = datetime.now()
        option = "1"  # Default to current week

        # Determine the date range mode from command line or prompt user
        if args.mode:
            option = str(args.mode)
        else:
            print("\nDate range options:")
            print("1. This week")
            print("2. This month")
            print("3. Custom date range")
            print("4. Specific week number")
            option = input("Select option (1-4): ")

        if option == '1':
            # This week (Monday to Sunday)
            today = now.date()
            start_date = datetime.combine(today - timedelta(days=today.weekday()), datetime.min.time())
            end_date = datetime.combine(start_date + timedelta(days=6), datetime.max.time())
            print(f"Selected current week: {start_date.strftime('%Y-%m-%d')} to {end_date.strftime('%Y-%m-%d')}")
        elif option == '2':
            # This month
            start_date = datetime(now.year, now.month, 1, 0, 0, 0)

            # Calculate the last day of the current month
            if now.month == 12:
                next_month = datetime(now.year + 1, 1, 1, 0, 0, 0)
            else:
                next_month = datetime(now.year, now.month + 1, 1, 0, 0, 0)

            end_date = next_month - timedelta(seconds=1)
            print(f"Selected current month: {start_date.strftime('%Y-%m-%d')} to {end_date.strftime('%Y-%m-%d')}")
        elif option == '3':
            # Custom range
            if args.start and args.end:
                from_date = args.start
                to_date = args.end
            else:
                from_date = input("Start date (YYYY-MM-DD): ")
                to_date = input("End date (YYYY-MM-DD): ")

            start_date = datetime.strptime(from_date, '%Y-%m-%d')
            # Set end date to the end of the day
            end_date = datetime.strptime(to_date, '%Y-%m-%d')
            end_date = datetime.combine(end_date.date(), datetime.max.time())
            print(f"Selected custom range: {start_date.strftime('%Y-%m-%d')} to {end_date.strftime('%Y-%m-%d')}")
        elif option == '4':
            # Specific week number
            if args.year and args.week:
                year = args.year
                week_num = args.week
            else:
                year = int(input("Year (YYYY): ") or now.year)
                week_num = int(input("Week number (1-52): "))

            start_date, end_date = get_week_dates(year, week_num)
            if start_date is None or end_date is None:
                print("Failed to calculate dates for the specified week.")
                return  # Exit if week calculation failed

            print(f"Selected week {week_num}: {start_date.strftime('%Y-%m-%d')} to {end_date.strftime('%Y-%m-%d')}")
        else:
            print("Invalid option. Using current week.")
            today = now.date()
            start_date = datetime.combine(today - timedelta(days=today.weekday()), datetime.min.time())
            end_date = datetime.combine(start_date + timedelta(days=6), datetime.max.time())
            print(f"Default to current week: {start_date.strftime('%Y-%m-%d')} to {end_date.strftime('%Y-%m-%d')}")

        appointments = get_appointments_in_range(calendar, start_date, end_date)

        if appointments:
            # Analyze excluding personal hours by default
            results = analyze_work_hours(appointments, exclude_personal=True)

            if not results:
                print("Error analyzing appointments.")
                return

            # Calculate hours needed to reach 40 (if within current week)
            hours_left = 0
            if option == '1':  # Current week
                if results['total_hours'] < 40:
                    hours_left = 40 - results['total_hours']

            print("\n===== WORK HOURS SUMMARY =====\n")
            print(f"Date Range: {start_date.strftime('%Y-%m-%d')} to {end_date.strftime('%Y-%m-%d')}")
            print(f"Total working hours: {results['total_hours']:.2f}")
            print(f"Personal hours (excluded from analysis): {results['personal_hours']:.2f}")

            if hours_left > 0:
                print(f"Hours needed to reach 40 this week: {hours_left:.2f}")

            print(f"Daily average (working hours): {results['daily_average']:.2f} hours")
            print(f"Working day average (Mon-Fri): {results['working_day_average']:.2f} hours")

            print("\n----- BY CATEGORY -----")
            print(results['category_summary'][['Category', 'Total Hours', 'Percentage', 'Number of Events']].to_string(index=False))

            print("\n----- WEEKLY SUMMARY -----")
            print(results['weekly_summary'].to_string(index=False))

            print("\n----- DAY OF WEEK SUMMARY -----")
            print(results['day_of_week_summary'].to_string(index=False))

            # Check if export is requested via command line arg or ask if not specified
            export_to_excel = args.export

            if export_to_excel:
                filename = f"work_hours_{start_date.strftime('%Y%m%d')}_{end_date.strftime('%Y%m%d')}.xlsx"

                with pd.ExcelWriter(filename) as writer:
                    results['dataframe'].to_excel(writer, sheet_name='Work Events', index=False)
                    results['complete_dataframe'].to_excel(writer, sheet_name='All Events', index=False)
                    results['category_summary'].to_excel(writer, sheet_name='By Category', index=False)
                    results['daily_summary'].to_excel(writer, sheet_name='Daily', index=False)
                    results['weekly_summary'].to_excel(writer, sheet_name='Weekly', index=False)
                    results['monthly_summary'].to_excel(writer, sheet_name='Monthly', index=False)
                    results['day_of_week_summary'].to_excel(writer, sheet_name='By Day of Week', index=False)

                print(f"Results exported to {filename}")
                print(f"Full path: {os.path.abspath(filename)}")
        else:
            print("No appointments found in the specified date range.")

    except Exception as e:
        print(f"An error occurred: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
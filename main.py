import win32com.client
from datetime import datetime, timedelta
import pandas as pd
import os
import sys

def connect_to_outlook():
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        calendar_folder = namespace.GetDefaultFolder(9)  # 9 = Calendar
        print(f"Connected to Outlook calendar: {calendar_folder.Name}")
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
        
        # Important: Sort items by start date to optimize filtering
        items.Sort("[Start]")
        items.IncludeRecurrences = True
        
        # Find appointments one month at a time to limit processing
        current_month_start = start_date.replace(day=1)
        appointments = []
        
        while current_month_start <= end_date:
            # Calculate end of month
            if current_month_start.month == 12:
                next_month = current_month_start.replace(year=current_month_start.year + 1, month=1)
            else:
                next_month = current_month_start.replace(month=current_month_start.month + 1)
            
            month_end = next_month - timedelta(seconds=1)
            
            # If month_end is beyond our end_date, cap it
            if month_end > end_date:
                month_end = end_date
            
            print(f"Processing calendar for {current_month_start.strftime('%B %Y')}...")
            
            # Create a filter for just this month
            month_start_str = current_month_start.strftime("%m/%d/%Y")
            month_end_str = month_end.strftime("%m/%d/%Y")
            
            # Try different date formats for filtering
            filter_attempts = [
                f"[Start] >= '{month_start_str}' AND [Start] <= '{month_end_str}'",
                f"[Start] >= '{current_month_start.strftime('%d/%m/%Y')}' AND [Start] <= '{month_end.strftime('%d/%m/%Y')}'",
                f"[Start] >= '{current_month_start.strftime('%Y-%m-%d')}' AND [Start] <= '{month_end.strftime('%Y-%m-%d')}'"
            ]
            
            month_items = None
            for filter_str in filter_attempts:
                try:
                    print(f"Trying filter: {filter_str}")
                    month_items = items.Restrict(filter_str)
                    # If we get here without error, exit the loop
                    break
                except:
                    continue
            
            # If all filter attempts failed, get items directly and filter manually
            if month_items is None:
                print("Filter failed, using manual method for this month...")
                month_items = items
                
            # Process this month's appointments
            month_appts = []
            item_count = 0
            
            for item in month_items:
                item_count += 1
                if item_count % 100 == 0:
                    print(f"Processed {item_count} items...")
                
                try:
                    # Extract start time
                    start_time = None
                    try:
                        start_time = item.Start
                        if isinstance(start_time, str):
                            start_time = pd.to_datetime(start_time)
                    except:
                        continue  # Skip if we can't get a start time
                    
                    # Manual filter if needed
                    if month_items == items:  # If we're using the full calendar
                        if not (current_month_start <= start_time <= month_end):
                            continue
                    
                    # Only include appointments within our overall date range
                    if not (start_date <= start_time <= end_date):
                        continue
                    
                    # Extract end time
                    end_time = None
                    try:
                        end_time = item.End
                        if isinstance(end_time, str):
                            end_time = pd.to_datetime(end_time)
                    except:
                        # If we can't get end time, estimate it
                        try:
                            end_time = start_time + timedelta(minutes=item.Duration)
                        except:
                            end_time = start_time + timedelta(hours=1)
                    
                    # Calculate duration in hours
                    try:
                        duration = (end_time - start_time).total_seconds() / 3600
                    except:
                        try:
                            duration = item.Duration / 60  # Minutes to hours
                        except:
                            duration = 1.0  # Default 1 hour
                    
                    # Get category (or set default)
                    try:
                        category = item.Categories if item.Categories else "Uncategorized"
                    except:
                        category = "Uncategorized"
                    
                    # Add to our list
                    month_appts.append({
                        'Subject': item.Subject,
                        'Start': start_time,
                        'End': end_time,
                        'Duration': duration,
                        'Category': category
                    })
                except Exception as e:
                    # Skip problematic items
                    continue
            
            print(f"Found {len(month_appts)} appointments in {current_month_start.strftime('%B %Y')}")
            appointments.extend(month_appts)
            
            # Move to next month
            current_month_start = next_month
        
        print(f"Total appointments found: {len(appointments)}")
        return appointments
    
    except Exception as e:
        print(f"Error getting appointments: {e}")
        return []

def analyze_work_hours(appointments):
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
    df['Week'] = df['Start'].dt.strftime('%Y-W%U')
    df['Month'] = df['Start'].dt.strftime('%Y-%m')
    
    # Summary by category
    category_summary = df.groupby('Category')['Duration'].agg(['sum', 'count']).reset_index()
    category_summary.columns = ['Category', 'Total Hours', 'Number of Events']
    category_summary = category_summary.sort_values('Total Hours', ascending=False)
    
    # Summary by day
    daily_summary = df.groupby(['Date', 'Day'])['Duration'].sum().reset_index()
    daily_summary = daily_summary.sort_values('Date')
    
    # Weekly summary
    weekly_summary = df.groupby('Week')['Duration'].sum().reset_index()
    
    # Monthly summary
    monthly_summary = df.groupby('Month')['Duration'].sum().reset_index()
    
    # Day of week summary
    dow_summary = df.groupby('Day')['Duration'].agg(['sum', 'mean']).reset_index()
    dow_summary.columns = ['Day', 'Total Hours', 'Average Hours']
    
    # Calculate daily average for working days (excluding weekends)
    working_days = df[~df['Day'].isin(['Saturday', 'Sunday'])]
    working_day_avg = working_days.groupby('Date')['Duration'].sum().mean() if not working_days.empty else 0
    
    return {
        'total_hours': df['Duration'].sum(),
        'daily_average': df.groupby('Date')['Duration'].sum().mean(),
        'working_day_average': working_day_avg,
        'category_summary': category_summary,
        'daily_summary': daily_summary,
        'weekly_summary': weekly_summary,
        'monthly_summary': monthly_summary,
        'day_of_week_summary': dow_summary,
        'dataframe': df
    }

def main():
    try:
        calendar = connect_to_outlook()
        
        # Ask for date range
        print("\nDate range options:")
        print("1. Last X days")
        print("2. This week")
        print("3. This month")
        print("4. Custom date range")
        option = input("Select option (1-4): ")
        
        now = datetime.now()
        
        if option == '1':
            days = int(input("Number of days to analyze: "))
            end_date = now
            start_date = end_date - timedelta(days=days)
        elif option == '2':
            # This week (Monday to Sunday)
            today = now.date()
            start_date = datetime.combine(today - timedelta(days=today.weekday()), datetime.min.time())
            end_date = now
        elif option == '3':
            # This month
            start_date = datetime(now.year, now.month, 1)
            end_date = now
        else:
            # Custom range
            from_date = input("Start date (YYYY-MM-DD): ")
            to_date = input("End date (YYYY-MM-DD): ")
            
            start_date = datetime.strptime(from_date, '%Y-%m-%d')
            end_date = datetime.strptime(to_date, '%Y-%m-%d') + timedelta(days=1) - timedelta(seconds=1)
        
        appointments = get_appointments_in_range(calendar, start_date, end_date)
        
        if appointments:
            results = analyze_work_hours(appointments)
            
            if not results:
                print("Error analyzing appointments.")
                return
                
            print("\n===== WORK HOURS SUMMARY =====")
            print(f"Date Range: {start_date.strftime('%Y-%m-%d')} to {end_date.strftime('%Y-%m-%d')}")
            print(f"Total hours tracked: {results['total_hours']:.2f}")
            print(f"Daily average: {results['daily_average']:.2f} hours")
            print(f"Working day average (Mon-Fri): {results['working_day_average']:.2f} hours")
            
            print("\n----- BY CATEGORY -----")
            print(results['category_summary'].to_string(index=False))
            
            print("\n----- WEEKLY SUMMARY -----")
            print(results['weekly_summary'].to_string(index=False))
            
            print("\n----- DAY OF WEEK SUMMARY -----")
            print(results['day_of_week_summary'].to_string(index=False))
            
            # Ask if user wants to export to Excel
            if input("\nExport results to Excel? (y/n): ").lower() == 'y':
                filename = f"work_hours_{start_date.strftime('%Y%m%d')}_{end_date.strftime('%Y%m%d')}.xlsx"
                
                with pd.ExcelWriter(filename) as writer:
                    results['dataframe'].to_excel(writer, sheet_name='All Events', index=False)
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
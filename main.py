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
        
        # Use a restriction to filter by date range
        # This helps avoid the infinite loop problem with recurring events
        restriction = f"[Start] >= '{start_date.strftime('%m/%d/%Y')}' AND [Start] <= '{end_date.strftime('%m/%d/%Y')}'"
        filtered_items = items.Restrict(restriction)
        print(f"Items after date restriction: {filtered_items.Count}")
        
        # DEBUG: Try to get at least one item to verify access
        if filtered_items.Count > 0:
            try:
                first_item = filtered_items.GetFirst()
                print(f"DEBUG: First filtered item subject: {first_item.Subject}")
                print(f"DEBUG: First filtered item start: {first_item.Start}")
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
                
                # DEBUG: Print date to ensure correct format handling
                if processed_count < 10:
                    print(f"DEBUG: Item {processed_count} start_time: {start_time}")
                    print(f"DEBUG: Type of start_time: {type(start_time)}")
                
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
                
                # Double-check if in date range (belt and suspenders approach)
                if not (start_date <= start_time <= end_date):
                    if processed_count < 10:
                        print(f"DEBUG: Item outside date range. Item date: {start_time}")
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
                
                # DEBUG: Print found item
                print(f"FOUND: {current_item.Subject}, Start: {start_time}, Duration: {duration:.2f} hours")
                
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
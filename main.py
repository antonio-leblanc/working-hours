import win32com.client
from datetime import datetime, timedelta
import pandas as pd

def connect_to_outlook():
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    calendar_folder = namespace.GetDefaultFolder(9)  # 9 = Calendar
    print(f"Connected to Outlook calendar: {calendar_folder.Name}")
    return calendar_folder

def get_appointments_in_range(calendar_folder, start_date, end_date):
    items = calendar_folder.Items
    items.IncludeRecurrences = True
    items.Sort("[Start]")
    
    # Create a restriction for the date range
    restriction = f"[Start] >= '{start_date.strftime('%m/%d/%Y 00:00')}' AND [Start] <= '{end_date.strftime('%m/%d/%Y 23:59')}'"
    print(f"Using restriction: {restriction}")
    filtered_items = items.Restrict(restriction)
    
    appointments = []
    for item in filtered_items:
        try:
            # Format datetimes properly
            start_time = item.Start
            if isinstance(start_time, str):
                start_time = datetime.strptime(start_time, '%Y-%m-%d %H:%M:%S%z').replace(tzinfo=None)
                
            end_time = item.End
            if isinstance(end_time, str):
                end_time = datetime.strptime(end_time, '%Y-%m-%d %H:%M:%S%z').replace(tzinfo=None)
            
            # Calculate duration in hours
            duration = (end_time - start_time).total_seconds() / 3600
            
            # Get category or set to "Uncategorized"
            category = item.Categories if hasattr(item, 'Categories') and item.Categories else "Uncategorized"
            
            appointments.append({
                'Subject': item.Subject,
                'Start': start_time,
                'End': end_time,
                'Duration': duration,
                'Category': category
            })
        except Exception as e:
            print(f"Error processing appointment '{item.Subject if hasattr(item, 'Subject') else 'Unknown'}': {e}")
    
    # If we got no appointments with Restrict, try the alternative method
    if len(appointments) == 0:
        print("No appointments found with Restrict method, trying alternative method...")
        appointments = get_appointments_in_range_alt(calendar_folder, start_date, end_date)
    
    return appointments

def get_appointments_in_range_alt(calendar_folder, start_date, end_date):
    items = calendar_folder.Items
    items.IncludeRecurrences = True
    items.Sort("[Start]")
    
    appointments = []
    count = 0
    
    for item in items:
        count += 1
        if count % 100 == 0:
            print(f"Processed {count} items...")
        
        try:
            # Convert to Python datetime objects
            try:
                if isinstance(item.Start, str):
                    appt_start = datetime.strptime(item.Start, '%Y-%m-%d %H:%M:%S%z').replace(tzinfo=None)
                else:
                    appt_start = datetime(
                        item.Start.year, item.Start.month, item.Start.day,
                        item.Start.hour, item.Start.minute, item.Start.second
                    )
                    
                if isinstance(item.End, str):
                    appt_end = datetime.strptime(item.End, '%Y-%m-%d %H:%M:%S%z').replace(tzinfo=None)
                else:
                    appt_end = datetime(
                        item.End.year, item.End.month, item.End.day,
                        item.End.hour, item.End.minute, item.End.second
                    )
            except AttributeError:
                # Skip items that don't have proper Start/End
                continue
            
            # Check if appointment is in range
            if start_date <= appt_start <= end_date:
                duration = (appt_end - appt_start).total_seconds() / 3600
                category = item.Categories if hasattr(item, 'Categories') and item.Categories else "Uncategorized"
                
                appointments.append({
                    'Subject': item.Subject,
                    'Start': appt_start,
                    'End': appt_end,
                    'Duration': duration,
                    'Category': category
                })
        except Exception as e:
            print(f"Error processing an appointment: {e}")
    
    return appointments

def analyze_work_hours(appointments):
    # Convert to DataFrame for easy analysis
    df = pd.DataFrame(appointments)
    
    # Split categories (Outlook often stores multiple categories as semicolon-separated values)
    df['Category'] = df['Category'].str.split(';')
    df = df.explode('Category').reset_index(drop=True)
    df['Category'] = df['Category'].str.strip()
    
    # Summary by category
    category_summary = df.groupby('Category')['Duration'].agg(['sum', 'count']).reset_index()
    category_summary.columns = ['Category', 'Total Hours', 'Number of Appointments']
    category_summary = category_summary.sort_values('Total Hours', ascending=False)
    
    # Summary by day
    df['Date'] = pd.to_datetime(df['Start']).dt.date
    daily_summary = df.groupby('Date')['Duration'].sum().reset_index()
    
    # Weekly summary
    df['Week'] = pd.to_datetime(df['Start']).dt.strftime('%Y-W%U')
    weekly_summary = df.groupby('Week')['Duration'].sum().reset_index()
    
    return {
        'total_hours': df['Duration'].sum(),
        'category_summary': category_summary,
        'daily_summary': daily_summary,
        'weekly_summary': weekly_summary,
        'dataframe': df
    }

def main():
    calendar = connect_to_outlook()
    
    # Ask for date range approach
    print("\nDate range options:")
    print("1. Last X days")
    print("2. Custom date range")
    option = input("Select option (1/2): ")
    
    if option == '1':
        days_to_analyze = int(input("Enter number of days to analyze: "))
        end_date = datetime.now()
        start_date = end_date - timedelta(days=days_to_analyze)
    else:
        # Custom date range
        from_date = input("Enter start date (YYYY-MM-DD): ")
        to_date = input("Enter end date (YYYY-MM-DD): ")
        
        start_date = datetime.strptime(from_date, '%Y-%m-%d')
        end_date = datetime.strptime(to_date, '%Y-%m-%d') + timedelta(days=1) - timedelta(seconds=1)
    
    print(f"\nAnalyzing calendar from {start_date.strftime('%Y-%m-%d')} to {end_date.strftime('%Y-%m-%d')}")
    appointments = get_appointments_in_range(calendar, start_date, end_date)
    print(f"Found {len(appointments)} appointments")
    
    if appointments:
        results = analyze_work_hours(appointments)
        
        print("\n===== WORK HOURS SUMMARY =====")
        print(f"Total hours: {results['total_hours']:.2f}")
        
        print("\n----- BY CATEGORY -----")
        print(results['category_summary'].to_string(index=False))
        
        print("\n----- WEEKLY SUMMARY -----")
        print(results['weekly_summary'].to_string(index=False))
        
        # Ask if user wants to export to Excel
        if input("\nExport results to Excel? (y/n): ").lower() == 'y':
            filename = f"work_hours_{start_date.strftime('%Y%m%d')}_{end_date.strftime('%Y%m%d')}.xlsx"
            with pd.ExcelWriter(filename) as writer:
                results['dataframe'].to_excel(writer, sheet_name='All Appointments', index=False)
                results['category_summary'].to_excel(writer, sheet_name='Category Summary', index=False)
                results['daily_summary'].to_excel(writer, sheet_name='Daily Summary', index=False)
                results['weekly_summary'].to_excel(writer, sheet_name='Weekly Summary', index=False)
            print(f"Results exported to {filename}")
    else:
        print("No appointments found in the specified date range.")

if __name__ == "__main__":
    main()
import win32com.client
import datetime
import pytz
import argparse
import sys
from collections import defaultdict
import calendar
import math
import time # For analysis timing

# --- Configuration ---
# Set to True to enable detailed debugging output
DEBUG = False
# Set your local timezone (Find yours here: https://en.wikipedia.org/wiki/List_of_tz_database_time_zones)
# Example: 'America/Sao_Paulo', 'Europe/London', 'America/New_York'
LOCAL_TIMEZONE = 'America/Sao_Paulo'
# Specify the category name for personal events that should be excluded from work hours
PERSONAL_CATEGORY = "Pessoal"
# Week definition (0=Monday, 6=Sunday)
WEEK_START_DAY = 0 # Monday
WEEK_END_DAY = 6   # Sunday
# --- End Configuration ---

def print_debug(message):
    """Prints a message only if DEBUG is True."""
    if DEBUG:
        print(f"DEBUG: {message}")

def get_local_timezone():
    """Gets the pytz timezone object based on LOCAL_TIMEZONE configuration."""
    print_debug(f"Using timezone: {LOCAL_TIMEZONE}")
    try:
        return pytz.timezone(LOCAL_TIMEZONE)
    except pytz.exceptions.UnknownTimeZoneError:
        print(f"Error: Unknown timezone '{LOCAL_TIMEZONE}'. Please check configuration.")
        sys.exit(1)

def parse_arguments():
    """Parses command-line arguments."""
    parser = argparse.ArgumentParser(description="Analyze Outlook Calendar work hours.")
    # Group for mutually exclusive period definition arguments
    period_group = parser.add_mutually_exclusive_group()
    period_group.add_argument("--period", choices=['week', 'month', 'specific_week'],
                        help="Time period to analyze (current week, current month, specific week by YYYY/WW).")
    period_group.add_argument("--month", type=int, choices=range(1, 13), metavar='M',
                        help="Specific month number (1-12) in the current year.")

    # Arguments specific to --period=specific_week
    parser.add_argument("--year", type=int, help="Year for specific week (e.g., 2023) - use with --period=specific_week.")
    parser.add_argument("--week", type=int, help="ISO week number for specific week (1-53) - use with --period=specific_week.")

    parser.add_argument("--debug", action='store_true', help="Enable debug printing.")

    args = parser.parse_args()

    # --- Argument Validation ---
    # Override DEBUG constant if --debug flag is set
    global DEBUG
    if args.debug:
        DEBUG = True
        print("DEBUG mode enabled via command line.")

    # Validate combinations
    if args.period == 'specific_week' and (args.year is None or args.week is None):
        parser.error("--year and --week are required when --period is 'specific_week'")
    if args.period != 'specific_week' and (args.year is not None or args.week is not None):
         parser.error("--year and --week can only be used when --period is 'specific_week'")
    if args.month is not None and (args.year is not None or args.week is not None):
        parser.error("--month cannot be used with --year or --week")
    if args.month is not None and args.period is not None:
         # This case is already handled by add_mutually_exclusive_group, but belt-and-suspenders
         parser.error("--month cannot be used with --period")

    return args

# <<< MODIFIED Function >>>
def get_time_period_interactive(local_tz):
    """Interactively asks the user for the time period."""
    while True:
        print("\nSelect the time period to analyze:")
        print("1. Current Week (Mon-Sun)")
        print("2. Current Month")
        print("3. Specific Week (by year and week number)")
        print("4. Specific Month (current year)") # Added option
        choice = input("Enter choice (1-4): ")

        now = datetime.datetime.now(local_tz)
        current_year = now.year
        print_debug(f"Current time ({local_tz}): {now}")

        if choice == '1':
            start_of_this_week = now - datetime.timedelta(days=now.weekday())
            start_date = start_of_this_week + datetime.timedelta(days=WEEK_START_DAY)
            end_date = start_date + datetime.timedelta(days=(WEEK_END_DAY - WEEK_START_DAY))
            period_name = f"Current Week ({start_date.strftime('%Y-%m-%d')} to {end_date.strftime('%Y-%m-%d')})"
            break
        elif choice == '2':
            start_date = now.replace(day=1)
            _, last_day = calendar.monthrange(now.year, now.month)
            end_date = now.replace(day=last_day)
            period_name = f"Current Month ({start_date.strftime('%B %Y')})"
            break
        elif choice == '3':
            while True:
                try:
                    year_str = input("Enter year (e.g., 2023): ")
                    year = int(year_str)
                    week_str = input("Enter ISO week number (1-53): ")
                    week_num = int(week_str)
                    if 1 <= week_num <= 53:
                        start_date_str = f"{year}-W{week_num:02}-1"
                        end_date_str = f"{year}-W{week_num:02}-7"
                        print_debug(f"Attempting to parse start date: {start_date_str}")
                        start_date_naive = datetime.datetime.strptime(start_date_str, "%G-W%V-%u")
                        print_debug(f"Attempting to parse end date: {end_date_str}")
                        end_date_naive = datetime.datetime.strptime(end_date_str, "%G-W%V-%u")
                        start_date = local_tz.localize(start_date_naive)
                        end_date = local_tz.localize(end_date_naive)
                        period_name = f"Week {week_num}, {year} ({start_date.strftime('%Y-%m-%d')} to {end_date.strftime('%Y-%m-%d')})"
                        break
                    else:
                        print("Invalid week number. Please enter a value between 1 and 53.")
                except ValueError as e:
                    print(f"Invalid input: {e}. Please enter numbers for year and week.")
                    print_debug(f"strptime failed for year={year_str}, week={week_str}")
            break # Exit outer loop once specific week is valid
        # <<< Added Block for option 4 >>>
        elif choice == '4':
            while True:
                try:
                    month_str = input(f"Enter month number for {current_year} (1-12): ")
                    month_num = int(month_str)
                    if 1 <= month_num <= 12:
                        print_debug(f"Calculating period for {current_year}-{month_num:02}")
                        start_date_naive = datetime.datetime(current_year, month_num, 1)
                        _, last_day = calendar.monthrange(current_year, month_num)
                        end_date_naive = datetime.datetime(current_year, month_num, last_day)
                        # Localize the naive dates
                        start_date = local_tz.localize(start_date_naive)
                        end_date = local_tz.localize(end_date_naive)
                        period_name = start_date.strftime('%B %Y') # Format as "Month Year"
                        break # Exit inner month validation loop
                    else:
                        print("Invalid month number. Please enter a value between 1 and 12.")
                except ValueError:
                    print("Invalid input. Please enter a number for the month.")
            break # Exit outer choice loop
        # <<< End of Added Block >>>
        else:
            print("Invalid choice. Please enter 1, 2, 3, or 4.")

    # Set time to start of day for start_date and end of day for end_date
    start_datetime = start_date.replace(hour=0, minute=0, second=0, microsecond=0)
    end_datetime = end_date.replace(hour=23, minute=59, second=59, microsecond=999999)

    print_debug(f"Calculated period start (inclusive): {start_datetime}")
    print_debug(f"Calculated period end (inclusive):   {end_datetime}")
    return start_datetime, end_datetime, period_name
# <<< End of MODIFIED Function >>>


# <<< MODIFIED Function >>>
def get_time_period_from_args(args, local_tz):
    """Determines the time period based on command-line arguments."""
    now = datetime.datetime.now(local_tz)
    current_year = now.year
    print_debug(f"Current time ({local_tz}): {now}")
    period_name = "Selected Period" # Default name

    start_date_naive = None # Use naive first, then localize
    end_date_naive = None

    if args.period == 'week':
        print_debug("Calculating period for current week via --period=week")
        start_of_this_week = now - datetime.timedelta(days=now.weekday())
        start_date = start_of_this_week + datetime.timedelta(days=WEEK_START_DAY) # Already localized 'now'
        end_date = start_date + datetime.timedelta(days=(WEEK_END_DAY - WEEK_START_DAY)) # Based on localized start_date
        period_name = f"Current Week ({start_date.strftime('%Y-%m-%d')} to {end_date.strftime('%Y-%m-%d')})"
    elif args.period == 'month':
        print_debug("Calculating period for current month via --period=month")
        start_date = now.replace(day=1) # Already localized 'now'
        _, last_day = calendar.monthrange(now.year, now.month)
        end_date = now.replace(day=last_day) # Based on localized 'now'
        period_name = f"Current Month ({start_date.strftime('%B %Y')})"
    elif args.period == 'specific_week':
        print_debug(f"Calculating period for specific week {args.week}/{args.year} via --period=specific_week")
        try:
            start_date_str = f"{args.year}-W{args.week:02}-1" # Monday
            end_date_str = f"{args.year}-W{args.week:02}-7"   # Sunday
            print_debug(f"Attempting to parse start date: {start_date_str}")
            start_date_naive = datetime.datetime.strptime(start_date_str, "%G-W%V-%u")
            print_debug(f"Attempting to parse end date: {end_date_str}")
            end_date_naive = datetime.datetime.strptime(end_date_str, "%G-W%V-%u")
            period_name = f"Week {args.week}, {args.year} ({start_date_naive.strftime('%Y-%m-%d')} to {end_date_naive.strftime('%Y-%m-%d')})" # Use naive for name before localization
        except ValueError as e:
             print(f"Error: Invalid year ({args.year}) or week ({args.week}) combination: {e}")
             print_debug(f"strptime failed for year={args.year}, week={args.week}")
             sys.exit(1)
    # <<< Added Block for --month argument >>>
    elif args.month:
        month_num = args.month
        print_debug(f"Calculating period for month {month_num}/{current_year} via --month argument")
        try:
            start_date_naive = datetime.datetime(current_year, month_num, 1)
            _, last_day = calendar.monthrange(current_year, month_num)
            end_date_naive = datetime.datetime(current_year, month_num, last_day)
            period_name = start_date_naive.strftime('%B %Y') # Format as "Month Year"
        except ValueError as e: # Should not happen with choices=[1-12] but good practice
            print(f"Error creating date for month {month_num}, year {current_year}: {e}")
            sys.exit(1)
    # <<< End of Added Block >>>
    else:
        # This case should not be reached if logic in main is correct
        print("Error: No valid period specified via command line arguments.")
        sys.exit(1)

    # --- Finalize Datetimes (Localize if needed, set times) ---
    if start_date_naive and end_date_naive:
         # Localize dates calculated as naive datetimes
         try:
             start_date = local_tz.localize(start_date_naive)
             end_date = local_tz.localize(end_date_naive)
         except Exception as e:
              print(f"Error localizing calculated dates: {e}")
              sys.exit(1)
    elif not start_date or not end_date: # Check if calculation failed for 'week' or 'month' periods
        print("Error: Failed to determine start or end date.")
        sys.exit(1)
    # Else: start_date and end_date were already localized (for current week/month)

    # Set time to start of day for start_date and end of day for end_date
    start_datetime = start_date.replace(hour=0, minute=0, second=0, microsecond=0)
    end_datetime = end_date.replace(hour=23, minute=59, second=59, microsecond=999999)

    print_debug(f"Calculated period start (inclusive): {start_datetime}")
    print_debug(f"Calculated period end (inclusive):   {end_datetime}")
    return start_datetime, end_datetime, period_name
# <<< End of MODIFIED Function >>>


def format_timedelta(delta):
    """Formats a timedelta object into H:MM."""
    if not isinstance(delta, datetime.timedelta):
        return "0:00"
    total_seconds = int(delta.total_seconds())
    sign = "-" if total_seconds < 0 else ""
    total_seconds = abs(total_seconds)
    hours = total_seconds // 3600
    minutes = (total_seconds % 3600) // 60
    return f"{sign}{hours}:{minutes:02}"

def analyze_calendar(start_dt_local, end_dt_local, local_tz):
    """Connects to Outlook, fetches events, and analyzes hours."""
    print("Connecting to Outlook...")
    try:
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            print_debug("Successfully connected using Dispatch.")
        except Exception as e_dispatch:
            print_debug(f"Dispatch failed: {e_dispatch}. Trying DispatchEx...")
            try:
                 outlook = win32com.client.DispatchEx("Outlook.Application")
                 print_debug("Successfully connected using DispatchEx.")
            except Exception as e_dispatchex:
                 print(f"Error connecting to Outlook via Dispatch and DispatchEx: {e_dispatchex}")
                 print("Please ensure Outlook is installed and accessible.")
                 sys.exit(1)

        namespace = outlook.GetNamespace("MAPI")
        calendar_folder = namespace.GetDefaultFolder(9)
        print_debug(f"Accessing calendar folder: {calendar_folder.Name}")
        items = calendar_folder.Items
        items.IncludeRecurrences = True
        print_debug("IncludeRecurrences set to True.")
        items.Sort("[Start]")
        print_debug("Items sorted by [Start].")

    except Exception as e:
        print(f"Error during Outlook setup: {e}")
        print("Please ensure Outlook is running and accessible.")
        sys.exit(1)

    print(f"Analyzing events from {start_dt_local.strftime('%Y-%m-%d %H:%M')} to {end_dt_local.strftime('%Y-%m-%d %H:%M')} ({LOCAL_TIMEZONE})...")
    print(f"(Excluding category: '{PERSONAL_CATEGORY}')")

    total_work_duration = datetime.timedelta()
    total_personal_duration = datetime.timedelta()
    category_durations = defaultdict(datetime.timedelta)
    daily_durations = defaultdict(datetime.timedelta)
    processed_event_count = 0
    considered_item_count = 0
    work_days_in_period = set()

    start_analysis_time = time.time()

    print_debug("Starting iteration through calendar items (sorted by start date)...")
    # --- Main Event Loop ---
    for item_index, item in enumerate(items):
        considered_item_count += 1
        if DEBUG and considered_item_count % 100 == 0:
             print_debug(f"Processing item #{considered_item_count}...")

        try:
            # Basic check
            if not hasattr(item, 'Start') or not hasattr(item, 'End') or not hasattr(item, 'Subject'):
                 print_debug(f"Skipping item #{considered_item_count} (index {item_index}) - missing essential properties.")
                 continue

            # Timezone Handling
            item_start_raw = item.Start
            item_end_raw = item.End
            item_subject = item.Subject

            if item_start_raw.tzinfo is None:
                try:
                    item_start_local = local_tz.localize(item_start_raw, is_dst=None)
                except (pytz.exceptions.NonExistentTimeError, pytz.exceptions.AmbiguousTimeError) as tz_err:
                    print(f"WARNING: Skipping item '{item_subject}' due to start time localization error ({item_start_raw}): {tz_err}.")
                    continue
            else:
                item_start_local = item_start_raw.astimezone(local_tz)

            # Break condition
            if item_start_local > end_dt_local:
                print_debug(f"Item '{item_subject}' starts after analysis end date. Stopping iteration.")
                break

            if item_end_raw.tzinfo is None:
                try:
                    item_end_local = local_tz.localize(item_end_raw, is_dst=None)
                except (pytz.exceptions.NonExistentTimeError, pytz.exceptions.AmbiguousTimeError) as tz_err:
                    print(f"WARNING: Skipping item '{item_subject}' due to end time localization error ({item_end_raw}): {tz_err}.")
                    continue
            else:
                item_end_local = item_end_raw.astimezone(local_tz)

            print_debug(f"\nConsidering item #{considered_item_count}: '{item_subject}'")
            print_debug(f"  Localized Start: {item_start_local.strftime('%Y-%m-%d %H:%M %Z%z')}, Localized End: {item_end_local.strftime('%Y-%m-%d %H:%M %Z%z')}")

            # Filtering & Duration Calculation
            event_overlaps = (item_start_local < end_dt_local and item_end_local > start_dt_local)
            print_debug(f"  Does event overlap? {event_overlaps}")

            if event_overlaps:
                effective_start = max(item_start_local, start_dt_local)
                effective_end = min(item_end_local, end_dt_local)

                if effective_start < effective_end:
                    duration = effective_end - effective_start
                    print_debug(f"  Calculated Duration within period: {duration}")
                    processed_event_count += 1

                    # Category Handling
                    categories_raw = item.Categories if hasattr(item, 'Categories') else ""
                    categories = [c.strip() for c in categories_raw.split(';') if c.strip()]
                    is_personal = PERSONAL_CATEGORY in categories
                    print_debug(f"  Categories: {categories} (Is personal? {is_personal})")

                    # Accumulate Durations
                    day_of_week = effective_start.weekday()
                    if is_personal:
                        total_personal_duration += duration
                        print_debug(f"  Adding {format_timedelta(duration)} to personal total.")
                    else:
                        total_work_duration += duration
                        daily_durations[day_of_week] += duration
                        print_debug(f"  Adding {format_timedelta(duration)} to WORK total and day {day_of_week}.")

                        first_work_category = "Uncategorized"
                        if categories:
                            for cat in categories:
                                if cat != PERSONAL_CATEGORY:
                                    first_work_category = cat
                                    break
                        category_durations[first_work_category] += duration
                        print_debug(f"  Adding {format_timedelta(duration)} to work category '{first_work_category}'.")

                        if 0 <= day_of_week <= 4:
                             work_days_in_period.add(effective_start.date())
                             print_debug(f"  Added {effective_start.date()} to set of work days.")
                else:
                    print_debug("  Duration is zero or negative after clipping, skipping.")

        except AttributeError as ae:
            item_subject_err = getattr(item, 'Subject', 'N/A')
            print_debug(f"Skipping item #{considered_item_count} ('{item_subject_err}') due to AttributeError: {ae}.")
            pass
        except (pytz.exceptions.NonExistentTimeError, pytz.exceptions.AmbiguousTimeError) as tz_err:
             item_subject_err = getattr(item, 'Subject', 'N/A')
             print_debug(f"Item '{item_subject_err}' was skipped earlier due to timezone error: {tz_err}")
        except Exception as e:
            item_subject_err = getattr(item, 'Subject', 'N/A')
            print(f"\n--- ERROR processing item #{considered_item_count} ---")
            print(f"Subject: '{item_subject_err}'")
            try:
                print(f"Details: Start={getattr(item, 'Start', 'N/A')}, End={getattr(item, 'End', 'N/A')}, Cat={getattr(item, 'Categories', 'N/A')}")
            except: print("Could not retrieve item details.")
            print(f"Error Type: {type(e).__name__}: {e}")
            print("Continuing analysis...")
            print_debug("----------------------------------------")
    # --- End of Loop ---

    end_analysis_time = time.time()
    print_debug(f"\nFinished iterating. Total items considered: {considered_item_count}.")
    print_debug(f"Analysis loop took {end_analysis_time - start_analysis_time:.2f} seconds.")

    if processed_event_count == 0:
         print("\nNo relevant calendar events found within the specified period and criteria.")
    else:
         print(f"\nProcessed {processed_event_count} relevant calendar events found overlapping the period.")

    # Debug Final Totals
    print_debug("\n--- Raw Totals Before Display ---")
    print_debug(f"Total Work Duration: {total_work_duration} ({format_timedelta(total_work_duration)})")
    print_debug(f"Total Personal Duration: {total_personal_duration} ({format_timedelta(total_personal_duration)})")
    print_debug("Category Durations (Work Time):")
    for cat, dur in sorted(category_durations.items()): print_debug(f"  - {cat}: {format_timedelta(dur)}")
    print_debug("Daily Durations (Work Time, 0=Mon):")
    for day in sorted(daily_durations.keys()): print_debug(f"  - Day {day}: {format_timedelta(daily_durations[day])}")
    print_debug(f"Unique work days (Mon-Fri) with entries: {len(work_days_in_period)} days")
    print_debug(f"Work days set: {sorted(list(work_days_in_period))}")
    print_debug("--- End Raw Totals ---")

    return total_work_duration, total_personal_duration, category_durations, daily_durations, len(work_days_in_period)


def display_results(period_name, start_dt, end_dt, total_work, total_personal, categories, dailies, num_work_days_with_entries):
    """Formats and prints the analysis results, including percentages."""
    print("\n--- Outlook Calendar Analysis ---")
    print(f"Period: {period_name}")
    print(f"From:   {start_dt.strftime('%Y-%m-%d %H:%M %Z')}")
    print(f"To:     {end_dt.strftime('%Y-%m-%d %H:%M %Z')}")
    print(f"Timezone: {LOCAL_TIMEZONE}")
    print("---------------------------------")

    total_work_formatted = format_timedelta(total_work)
    total_personal_formatted = format_timedelta(total_personal)
    total_work_seconds = total_work.total_seconds()

    print(f"Total Work Hours:       {total_work_formatted}")
    if total_personal > datetime.timedelta(0) or PERSONAL_CATEGORY in categories:
        print(f"Total Personal Hours:   {total_personal_formatted} (Category: '{PERSONAL_CATEGORY}')")
    print("---------------------------------")

    print("Work Hours by Specific Category (% of Total Work Hours):")
    work_categories_filtered = {k: v for k, v in categories.items()
                                if k != PERSONAL_CATEGORY and k != "Uncategorized" and v > datetime.timedelta(0)}

    if work_categories_filtered:
        for category, duration in sorted(work_categories_filtered.items()):
            percentage = (duration.total_seconds() / total_work_seconds * 100) if total_work_seconds > 0 else 0.0
            print(f"  - {category}: {format_timedelta(duration)} ({percentage:.1f}%)")
    elif total_work_seconds > 0: # Only show if there *was* work time
         print("  (No time recorded under specific work categories)")

    uncategorized_duration = categories.get("Uncategorized", datetime.timedelta(0))
    if uncategorized_duration > datetime.timedelta(0):
        uncat_percentage = (uncategorized_duration.total_seconds() / total_work_seconds * 100) if total_work_seconds > 0 else 0.0
        print(f"  - Uncategorized Work: {format_timedelta(uncategorized_duration)} ({uncat_percentage:.1f}%)")

    if work_categories_filtered or uncategorized_duration > datetime.timedelta(0):
        print(f"\nNote: 'Total Work Hours' ({total_work_formatted}) is the base (100%) for category percentages.")
        print(f"      It includes all non-'{PERSONAL_CATEGORY}' time.")

    print("---------------------------------")

    print("Work Hours by Day of Week:")
    days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
    total_daily_check = datetime.timedelta()
    for i in range(7):
        duration = dailies.get(i, datetime.timedelta())
        print(f"  - {days[i]}: {format_timedelta(duration)}")
        total_daily_check += duration
    if not math.isclose(total_daily_check.total_seconds(), total_work_seconds, rel_tol=1e-5):
         print_debug(f"WARNING: Sum of daily hours ({format_timedelta(total_daily_check)}) doesn't match total work hours ({total_work_formatted})!")

    print("---------------------------------")

    # Calculate Averages
    num_days_in_period = (end_dt.date() - start_dt.date()).days + 1
    potential_working_days = 0
    current_day = start_dt.date()
    end_period_date = end_dt.date()
    while current_day <= end_period_date:
        if 0 <= current_day.weekday() <= 4:
            potential_working_days += 1
        current_day += datetime.timedelta(days=1)

    print_debug(f"Calculated {num_days_in_period} total calendar days in period.")
    print_debug(f"Calculated {potential_working_days} potential working days (Mon-Fri) in period.")
    print_debug(f"Detected {num_work_days_with_entries} actual working days (Mon-Fri) with logged work time.")

    avg_hours_per_day = total_work / num_days_in_period if num_days_in_period > 0 else datetime.timedelta()
    avg_hours_per_potential_working_day = total_work / potential_working_days if potential_working_days > 0 else datetime.timedelta()
    avg_hours_per_actual_working_day = total_work / num_work_days_with_entries if num_work_days_with_entries > 0 else datetime.timedelta()

    print("Averages (Based on Total Work Hours):")
    print(f"  - Avg / Day (all {num_days_in_period} days):           {format_timedelta(avg_hours_per_day)}")
    print(f"  - Avg / Potential Work Day (Mon-Fri): {format_timedelta(avg_hours_per_potential_working_day)} ({potential_working_days} potential days)")
    print(f"  - Avg / Actual Worked Day (Mon-Fri): {format_timedelta(avg_hours_per_actual_working_day)} ({num_work_days_with_entries} days with entries)")
    print("---------------------------------\n")

# --- Main Execution ---
if __name__ == "__main__":
    # Optional Exception Handler
    def handle_exception(exc_type, exc_value, exc_traceback):
        if issubclass(exc_type, KeyboardInterrupt):
            sys.__excepthook__(exc_type, exc_value, exc_traceback)
            return
        print("\n--- UNHANDLED EXCEPTION ---", file=sys.stderr)
        print(f"Type: {exc_type.__name__}", file=sys.stderr)
        print(f"Value: {exc_value}", file=sys.stderr)
        import traceback
        traceback.print_exception(exc_type, exc_value, exc_traceback, file=sys.stderr)
        print("--- END ---", file=sys.stderr)
        if DEBUG: input("Press Enter to exit after viewing error...")
        sys.exit(1)
    # sys.excepthook = handle_exception # Uncomment to enable

    args = parse_arguments()
    local_tz = get_local_timezone()

    start_dt, end_dt, period_name = (None, None, None)

    # <<< MODIFIED Logic: Decide input method >>>
    # Check if any command-line argument defining a period was used
    if args.period or args.month or (args.year and args.week):
         print_debug("Using command-line arguments for period selection.")
         # Ensure --year/--week are only used if --period=specific_week was set (already checked in parse_arguments)
         start_dt, end_dt, period_name = get_time_period_from_args(args, local_tz)
    else:
         print_debug("No period arguments provided, using interactive selection.")
         start_dt, end_dt, period_name = get_time_period_interactive(local_tz)
    # <<< End of MODIFIED Logic >>>


    if start_dt and end_dt:
        try:
            total_w, total_p, cats, days, work_days_count = analyze_calendar(start_dt, end_dt, local_tz)
            display_results(period_name, start_dt, end_dt, total_w, total_p, cats, days, work_days_count)
        except Exception as main_e:
            print(f"\n--- An unexpected error occurred during the main process ---")
            print(f"Error Type: {type(main_e).__name__}: {main_e}")
            print("----------------------------------------------------------")
            if DEBUG:
                 import traceback
                 print("Traceback:")
                 traceback.print_exc()
            sys.exit(1)
    else:
         print("Error: Could not determine valid start and end dates for analysis. Exiting.")
         sys.exit(1)

    print("Analysis complete.")
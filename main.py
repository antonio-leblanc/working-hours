import win32com.client
import datetime
import pytz
import argparse
import sys
from collections import defaultdict
import calendar
import math
import time # Import time for potential sleep/debug pauses

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
    """Gets the pytz timezone object."""
    print_debug(f"Using timezone: {LOCAL_TIMEZONE}")
    try:
        return pytz.timezone(LOCAL_TIMEZONE)
    except pytz.exceptions.UnknownTimeZoneError:
        print(f"Error: Unknown timezone '{LOCAL_TIMEZONE}'. Please check configuration.")
        sys.exit(1)

def parse_arguments():
    """Parses command-line arguments."""
    parser = argparse.ArgumentParser(description="Analyze Outlook Calendar work hours.")
    parser.add_argument("--period", choices=['week', 'month', 'specific_week'],
                        help="Time period to analyze (current week, current month, specific week).")
    parser.add_argument("--year", type=int, help="Year for specific week (e.g., 2023).")
    parser.add_argument("--week", type=int, help="ISO week number for specific week (1-53).")
    parser.add_argument("--debug", action='store_true', help="Enable debug printing.") # Add debug flag

    args = parser.parse_args()

    # Override DEBUG constant if --debug flag is set
    global DEBUG
    if args.debug:
        DEBUG = True
        print("DEBUG mode enabled via command line.")


    if args.period == 'specific_week' and (args.year is None or args.week is None):
        parser.error("--year and --week are required when --period is 'specific_week'")
    if args.period is None and (args.year is not None or args.week is not None):
        parser.error("--year and --week can only be used when --period is 'specific_week'")

    return args

def get_time_period_interactive(local_tz):
    """Interactively asks the user for the time period."""
    while True:
        print("\nSelect the time period to analyze:")
        print("1. Current Week (Mon-Sun)")
        print("2. Current Month")
        print("3. Specific Week (by year and week number)")
        choice = input("Enter choice (1-3): ")

        now = datetime.datetime.now(local_tz)
        print_debug(f"Current time ({local_tz}): {now}")

        if choice == '1':
            start_date = now - datetime.timedelta(days=now.weekday() - WEEK_START_DAY)
            end_date = start_date + datetime.timedelta(days=WEEK_END_DAY - WEEK_START_DAY)
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
                         # Use ISO standard %G (ISO year), %V (ISO week), %u (ISO weekday 1-7)
                        start_date_str = f"{year}-W{week_num:02}-1" # Monday
                        end_date_str = f"{year}-W{week_num:02}-7"   # Sunday
                        print_debug(f"Attempting to parse start date: {start_date_str}")
                        start_date = datetime.datetime.strptime(start_date_str, "%G-W%V-%u")
                        print_debug(f"Attempting to parse end date: {end_date_str}")
                        end_date = datetime.datetime.strptime(end_date_str, "%G-W%V-%u")
                        # Make dates timezone aware using the *local* timezone
                        start_date = local_tz.localize(start_date)
                        end_date = local_tz.localize(end_date)
                        period_name = f"Week {week_num}, {year} ({start_date.strftime('%Y-%m-%d')} to {end_date.strftime('%Y-%m-%d')})"
                        break
                    else:
                        print("Invalid week number. Please enter a value between 1 and 53.")
                except ValueError as e:
                    print(f"Invalid input: {e}. Please enter numbers for year and week.")
                    print_debug(f"strptime failed for year={year_str}, week={week_str}")
            break # Exit outer loop once specific week is valid
        else:
            print("Invalid choice. Please enter 1, 2, or 3.")

    # Set time to start of day for start_date and end of day for end_date
    start_datetime = start_date.replace(hour=0, minute=0, second=0, microsecond=0)
    end_datetime = end_date.replace(hour=23, minute=59, second=59, microsecond=999999)

    print_debug(f"Calculated period start (inclusive): {start_datetime}")
    print_debug(f"Calculated period end (inclusive):   {end_datetime}")
    return start_datetime, end_datetime, period_name


def get_time_period_from_args(args, local_tz):
    """Determines the time period based on command-line arguments."""
    now = datetime.datetime.now(local_tz)
    print_debug(f"Current time ({local_tz}): {now}")
    period_name = "Selected Period"

    if args.period == 'week':
        start_date = now - datetime.timedelta(days=now.weekday() - WEEK_START_DAY)
        end_date = start_date + datetime.timedelta(days=WEEK_END_DAY - WEEK_START_DAY)
        period_name = f"Current Week ({start_date.strftime('%Y-%m-%d')} to {end_date.strftime('%Y-%m-%d')})"
    elif args.period == 'month':
        start_date = now.replace(day=1)
        _, last_day = calendar.monthrange(now.year, now.month)
        end_date = now.replace(day=last_day)
        period_name = f"Current Month ({start_date.strftime('%B %Y')})"
    elif args.period == 'specific_week':
        try:
            # Use ISO standard %G (ISO year), %V (ISO week), %u (ISO weekday 1-7)
            start_date_str = f"{args.year}-W{args.week:02}-1" # Monday
            end_date_str = f"{args.year}-W{args.week:02}-7"   # Sunday
            print_debug(f"Attempting to parse start date: {start_date_str}")
            start_date = datetime.datetime.strptime(start_date_str, "%G-W%V-%u")
            print_debug(f"Attempting to parse end date: {end_date_str}")
            end_date = datetime.datetime.strptime(end_date_str, "%G-W%V-%u")
            # Make dates timezone aware using the *local* timezone
            start_date = local_tz.localize(start_date)
            end_date = local_tz.localize(end_date)
            period_name = f"Week {args.week}, {args.year} ({start_date.strftime('%Y-%m-%d')} to {end_date.strftime('%Y-%m-%d')})"
        except ValueError as e:
             print(f"Error: Invalid year ({args.year}) or week ({args.week}) combination: {e}")
             print_debug(f"strptime failed for year={args.year}, week={args.week}")
             sys.exit(1)
    else:
        # Should not happen if args are parsed correctly, but good for safety
        print("Error: Invalid period specified.")
        sys.exit(1)

    # Set time to start of day for start_date and end of day for end_date
    start_datetime = start_date.replace(hour=0, minute=0, second=0, microsecond=0)
    end_datetime = end_date.replace(hour=23, minute=59, second=59, microsecond=999999)

    print_debug(f"Calculated period start (inclusive): {start_datetime}")
    print_debug(f"Calculated period end (inclusive):   {end_datetime}")
    return start_datetime, end_datetime, period_name


def format_timedelta(delta):
    """Formats a timedelta object into H:MM."""
    total_seconds = int(delta.total_seconds())
    hours = total_seconds // 3600
    minutes = (total_seconds % 3600) // 60
    return f"{hours}:{minutes:02}"

def analyze_calendar(start_dt_local, end_dt_local, local_tz):
    """Connects to Outlook, fetches events, and analyzes hours."""
    print("Connecting to Outlook...")
    try:
        # Try Dispatch first, then DispatchEx if needed
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            print_debug("Successfully connected using Dispatch.")
        except Exception as e_dispatch:
            print_debug(f"Dispatch failed: {e_dispatch}. Trying DispatchEx...")
            try:
                 # DispatchEx might be needed if Outlook isn't already running well
                 outlook = win32com.client.DispatchEx("Outlook.Application")
                 print_debug("Successfully connected using DispatchEx.")
            except Exception as e_dispatchex:
                 print(f"Error connecting to Outlook via Dispatch and DispatchEx: {e_dispatchex}")
                 print("Please ensure Outlook is installed and accessible.")
                 sys.exit(1)

        namespace = outlook.GetNamespace("MAPI")
        calendar_folder = namespace.GetDefaultFolder(9) # 9 corresponds to the Calendar folder
        print_debug(f"Accessing calendar folder: {calendar_folder.Name}")
        items = calendar_folder.Items
        # Important: Include recurrences and sort by start time
        items.IncludeRecurrences = True
        print_debug("IncludeRecurrences set to True.")
        items.Sort("[Start]")
        print_debug("Items sorted by [Start].")

        # Note: Using Restrict can be faster but often struggles with dates/times/recurrences.
        # We will filter manually below, which is more robust.
        # filter_str = f"[Start] >= '{start_dt_local.strftime('%m/%d/%Y %H:%M %p')}' AND [End] <= '{end_dt_local.strftime('%m/%d/%Y %H:%M %p')}'"
        # print_debug(f"Filtering with string (for reference, not active): {filter_str}")
        # restricted_items = items.Restrict(filter_str) # Not using this actively

    except Exception as e:
        print(f"Error during Outlook setup: {e}")
        print("Please ensure Outlook is running and accessible.")
        sys.exit(1)

    print(f"Analyzing events from {start_dt_local.strftime('%Y-%m-%d %H:%M')} to {end_dt_local.strftime('%Y-%m-%d %H:%M')} ({LOCAL_TIMEZONE})...")
    print_debug(f"Personal category defined as: '{PERSONAL_CATEGORY}'")

    total_work_duration = datetime.timedelta()
    total_personal_duration = datetime.timedelta()
    category_durations = defaultdict(datetime.timedelta)
    daily_durations = defaultdict(datetime.timedelta)
    processed_event_count = 0
    considered_item_count = 0
    work_days_in_period = set() # Store unique work days (Mon-Fri) found with events

    start_analysis_time = time.time()

    # --- Iterate through ALL items and filter manually ---
    print_debug("Starting iteration through calendar items...")
    for item_index, item in enumerate(items):
        considered_item_count += 1
        if DEBUG and considered_item_count % 100 == 0:
             print_debug(f"Processing item #{considered_item_count}...")
             # time.sleep(0.01) # Optional small pause for very large calendars

        try:
            # Basic check if item has Start/End properties (skip non-appointment items)
            if not hasattr(item, 'Start') or not hasattr(item, 'End') or not hasattr(item, 'Subject'):
                 print_debug(f"Skipping item #{considered_item_count} (index {item_index}) - missing essential properties (likely not an appointment).")
                 continue

            # --- Timezone Handling ---
            item_start_raw = item.Start
            item_end_raw = item.End
            item_subject = item.Subject

            # Check if COM object time is naive or aware
            if item_start_raw.tzinfo is None:
                # Assume naive time from COM is in the system's *local* time, localize to target TZ
                try:
                    item_start_local = local_tz.localize(item_start_raw)
                except (pytz.exceptions.NonExistentTimeError, pytz.exceptions.AmbiguousTimeError) as tz_err:
                    print_debug(f"Skipping item '{item_subject}' due to start time localization error ({item_start_raw}): {tz_err}")
                    continue # Skip this item if time is invalid in local TZ
            else:
                # If already timezone-aware, convert to target local timezone
                item_start_local = item_start_raw.astimezone(local_tz)

            # --- *** ADDED BREAK CONDITION *** ---
            # If the item's start time is already past the end of our analysis period,
            # and since items are sorted, we can stop iterating.
            if item_start_local > end_dt_local:
                print_debug(f"Item '{item_subject}' starts at {item_start_local}, which is after the analysis end date {end_dt_local}. Stopping iteration.")
                break # Exit the loop

            # Process End Time (only if we haven't broken out)
            if item_end_raw.tzinfo is None:
                try:
                    item_end_local = local_tz.localize(item_end_raw)
                except (pytz.exceptions.NonExistentTimeError, pytz.exceptions.AmbiguousTimeError) as tz_err:
                    print_debug(f"Skipping item '{item_subject}' due to end time localization error ({item_end_raw}): {tz_err}")
                    continue # Skip this item
            else:
                item_end_local = item_end_raw.astimezone(local_tz)


            print_debug(f"\nConsidering item #{considered_item_count}: '{item_subject}'")
            print_debug(f"  Raw Start: {item_start_raw}, Raw End: {item_end_raw}")
            print_debug(f"  Localized Start: {item_start_local}, Localized End: {item_end_local}")


            # --- Filtering Logic ---
            # Check if the event *overlaps* with the desired period
            # Overlap: (StartA < EndB) and (EndA > StartB)
            event_overlaps = (item_start_local < end_dt_local and item_end_local > start_dt_local)
            print_debug(f"  Checking overlap with period: {start_dt_local} to {end_dt_local}")
            print_debug(f"  Does event overlap? {event_overlaps}")

            if event_overlaps:
                processed_event_count += 1
                # Calculate the duration *within* the analysis period if it spans boundaries
                # Effective start is the later of event start or period start
                effective_start = max(item_start_local, start_dt_local)
                # Effective end is the earlier of event end or period end
                effective_end = min(item_end_local, end_dt_local)

                duration = effective_end - effective_start
                print_debug(f"  Event overlaps. Effective Start: {effective_start}, Effective End: {effective_end}")
                print_debug(f"  Calculated Duration within period: {duration}")


                # Ignore events with zero or negative duration after clipping
                if duration <= datetime.timedelta(0):
                    print_debug("  Duration is zero or negative after clipping, skipping.")
                    continue

                # --- Category Handling ---
                categories_raw = item.Categories if hasattr(item, 'Categories') else ""
                categories = categories_raw.split(';') if categories_raw else []
                categories = [c.strip() for c in categories if c.strip()] # Clean up whitespace and empty strings
                print_debug(f"  Categories: {categories} (Raw: '{categories_raw}')")

                is_personal = PERSONAL_CATEGORY in categories
                print_debug(f"  Is personal ('{PERSONAL_CATEGORY}')? {is_personal}")

                # --- Accumulate Durations ---
                # Use effective_start for determining the day, as that's the portion within the period
                day_of_week = effective_start.weekday() # Monday = 0, Sunday = 6

                if is_personal:
                    total_personal_duration += duration
                    print_debug(f"  Adding {duration} to personal total.")
                else:
                    # This is a work event
                    total_work_duration += duration
                    daily_durations[day_of_week] += duration
                    print_debug(f"  Adding {duration} to work total and day {day_of_week}.")

                    # Track work hours per category
                    first_work_category = "Uncategorized"
                    if categories:
                        # Find the first category that isn't the personal one
                        found_work_cat = False
                        for cat in categories:
                             if cat != PERSONAL_CATEGORY:
                                 first_work_category = cat
                                 found_work_cat = True
                                 break
                        if not found_work_cat:
                             print_debug(f"  Item only had personal category '{PERSONAL_CATEGORY}', but was not flagged as personal? Check logic. Treating as Uncategorized work.")
                             # This case shouldn't ideally happen if is_personal is correct, but good safeguard
                    category_durations[first_work_category] += duration
                    print_debug(f"  Adding {duration} to work category '{first_work_category}'.")


                    # Track which days (Mon-Fri) had work events using the date part of effective_start
                    if 0 <= day_of_week <= 4: # Monday to Friday
                         work_days_in_period.add(effective_start.date())
                         print_debug(f"  Added {effective_start.date()} to set of work days with entries.")


        except AttributeError as ae:
            # Handle items that might not have expected properties (e.g., non-appointments)
            item_subject_err = getattr(item, 'Subject', 'N/A')
            print_debug(f"Skipping item #{considered_item_count} ('{item_subject_err}') due to AttributeError: {ae}")
            pass # Continue to the next item
        except pytz.exceptions.NonExistentTimeError as neste:
             item_subject_err = getattr(item, 'Subject', 'N/A')
             print(f"WARNING: Skipping item '{item_subject_err}' due to NonExistentTimeError (DST issue?): {neste}")
        except pytz.exceptions.AmbiguousTimeError as ate:
             item_subject_err = getattr(item, 'Subject', 'N/A')
             print(f"WARNING: Skipping item '{item_subject_err}' due to AmbiguousTimeError (DST issue?): {ate}")
        except Exception as e:
            # Catch other potential errors with specific items
            item_subject_err = getattr(item, 'Subject', 'N/A')
            print(f"ERROR processing item #{considered_item_count} ('{item_subject_err}'): {e}")
            print_debug(f"  Error details: {type(e).__name__}")
            # Optionally add more error details here if needed
            # Consider whether to continue or stop on such errors

    end_analysis_time = time.time()
    print_debug(f"Finished iterating through items. Total items considered: {considered_item_count}.")
    print_debug(f"Analysis loop took {end_analysis_time - start_analysis_time:.2f} seconds.")

    print(f"Processed {processed_event_count} relevant calendar events found within the period.")

    # --- Debug final totals before display ---
    print_debug("\n--- Raw Totals Before Display ---")
    print_debug(f"Total Work Duration: {total_work_duration}")
    print_debug(f"Total Personal Duration: {total_personal_duration}")
    print_debug("Category Durations:")
    for cat, dur in category_durations.items():
        print_debug(f"  - {cat}: {dur}")
    print_debug("Daily Durations (0=Mon):")
    for day, dur in daily_durations.items():
         print_debug(f"  - {day}: {dur}")
    print_debug(f"Unique work days (Mon-Fri) with entries: {len(work_days_in_period)} days")
    print_debug(f"Work days set: {work_days_in_period}")
    print_debug("--- End Raw Totals ---")

    return total_work_duration, total_personal_duration, category_durations, daily_durations, len(work_days_in_period)

def display_results(period_name, start_dt, end_dt, total_work, total_personal, categories, dailies, num_work_days_with_entries):
    """Formats and prints the analysis results."""
    print("\n--- Outlook Calendar Analysis ---")
    print(f"Period: {period_name}")
    print(f"Timezone: {LOCAL_TIMEZONE}")
    print("---------------------------------")

    print(f"Total Work Hours:       {format_timedelta(total_work)}")
    if total_personal > datetime.timedelta(0) or PERSONAL_CATEGORY in categories:
         # Show personal hours line if there were personal events OR if the category exists even with 0 hours tracked (for clarity)
        print(f"Total Personal Hours:   {format_timedelta(total_personal)} (Category: '{PERSONAL_CATEGORY}')")
    print("---------------------------------")

    print("Work Hours by Category:")
    # Exclude the personal category explicitly from this list if it somehow got work hours (shouldn't happen with current logic)
    work_categories = {k: v for k, v in categories.items() if k != PERSONAL_CATEGORY}
    if work_categories:
        # Sort categories alphabetically for consistent output
        for category, duration in sorted(work_categories.items()):
            print(f"  - {category}: {format_timedelta(duration)}")
    else:
        # Check if there was work time but it was all uncategorized
        if total_work > datetime.timedelta(0) and "Uncategorized" not in categories:
             print("  (All work time was uncategorized or category tracking failed)")
        elif not total_work > datetime.timedelta(0):
             print("  (No work time found)")
        else:
             print("  (No categorized work events found)")

    print("---------------------------------")

    print("Work Hours by Day of Week:")
    days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
    total_daily_check = datetime.timedelta() # For debug check
    for i in range(7):
        duration = dailies.get(i, datetime.timedelta())
        print(f"  - {days[i]}: {format_timedelta(duration)}")
        total_daily_check += duration
    print_debug(f"Sum of daily work hours: {total_daily_check} (Should match total work hours)")
    if not math.isclose(total_daily_check.total_seconds(), total_work.total_seconds()):
         print_debug("WARNING: Sum of daily hours doesn't match total work hours!")

    print("---------------------------------")

    # Calculate Averages
    num_days_in_period = (end_dt.date() - start_dt.date()).days + 1
    avg_hours_per_day = total_work / num_days_in_period if num_days_in_period > 0 else datetime.timedelta()

    # Calculate number of potential working days (Mon-Fri) in the period
    potential_working_days = 0
    current_day = start_dt.date()
    end_period_date = end_dt.date()
    while current_day <= end_period_date:
        # Check if the day's weekday() is between 0 (Monday) and 4 (Friday)
        if 0 <= current_day.weekday() <= 4:
            potential_working_days += 1
        current_day += datetime.timedelta(days=1)
    print_debug(f"Calculated {num_days_in_period} total days in period.")
    print_debug(f"Calculated {potential_working_days} potential working days (Mon-Fri) in period.")

    avg_hours_per_potential_working_day = total_work / potential_working_days if potential_working_days > 0 else datetime.timedelta()
    # Optional: Average based only on days you *actually* worked
    avg_hours_per_actual_working_day = total_work / num_work_days_with_entries if num_work_days_with_entries > 0 else datetime.timedelta()


    print("Averages:")
    print(f"  - Avg Work Hours / Day (all days):           {format_timedelta(avg_hours_per_day)}")
    print(f"  - Avg Work Hours / Working Day (Mon-Fri in period): {format_timedelta(avg_hours_per_potential_working_day)} ({potential_working_days} potential days)")
    print(f"  - Avg Work Hours / Actual Worked Day (Mon-Fri): {format_timedelta(avg_hours_per_actual_working_day)} ({num_work_days_with_entries} days with entries)")
    print("---------------------------------\n")


# --- Main Execution ---
if __name__ == "__main__":
    # --- Global Exception Handler for Debugging ---
    def handle_exception(exc_type, exc_value, exc_traceback):
        if issubclass(exc_type, KeyboardInterrupt):
            # Default behavior for Ctrl+C
            sys.__excepthook__(exc_type, exc_value, exc_traceback)
            return
        print("\n--- UNHANDLED EXCEPTION ---", file=sys.stderr)
        print(f"Type: {exc_type.__name__}", file=sys.stderr)
        print(f"Value: {exc_value}", file=sys.stderr)
        print("Traceback:", file=sys.stderr)
        import traceback
        traceback.print_exception(exc_type, exc_value, exc_traceback, file=sys.stderr)
        print("--- END UNHANDLED EXCEPTION ---\n", file=sys.stderr)
        # If DEBUG is on, maybe wait for input?
        if DEBUG:
             input("Press Enter to exit after viewing the error...")

    # sys.excepthook = handle_exception # Uncomment this line to catch *all* unhandled exceptions

    args = parse_arguments() # Parse args first to potentially enable DEBUG mode
    local_tz = get_local_timezone()

    if args.period:
        start_dt, end_dt, period_name = get_time_period_from_args(args, local_tz)
    else:
        start_dt, end_dt, period_name = get_time_period_interactive(local_tz)

    if start_dt and end_dt:
        try:
            total_w, total_p, cats, days, work_days_count = analyze_calendar(start_dt, end_dt, local_tz)
            display_results(period_name, start_dt, end_dt, total_w, total_p, cats, days, work_days_count)
        except Exception as main_e:
            print(f"\nAn unexpected error occurred during the main analysis execution: {main_e}")
            print_debug(f"Error type: {type(main_e).__name__}")
            # Optionally re-raise or print traceback if debugging
            if DEBUG:
                 import traceback
                 print("Traceback:")
                 traceback.print_exc()
            sys.exit(1)
    else:
         print("Error: Could not determine valid start and end dates for analysis.")
         sys.exit(1)
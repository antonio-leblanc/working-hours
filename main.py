import win32com.client

def connect_to_outlook():
    # Create an instance of the Outlook application
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    
    # Get the default Calendar folder (folder number 9 is typically the Calendar)
    calendar_folder = namespace.GetDefaultFolder(9)
    
    # Print some basic info to confirm connection
    print("Connected to Outlook.")
    print("Calendar folder name:", calendar_folder.Name)
    print("Total items in calendar:", len(calendar_folder.Items))
    
    return calendar_folder

if __name__ == "__main__":
    calendar = connect_to_outlook()

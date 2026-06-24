import pandas as pd
from openpyxl import load_workbook
from datetime import datetime

# Excel file path
EXCEL_PATH = r"C:\Users\csp\TO DO\To do list.xlsx"

# ── READ FUNCTIONS ──────────────────────────────────────────────

def read_all_tasks():
    """Read and display all tasks from the Excel file."""
    try:
        df = pd.read_excel(EXCEL_PATH)
        if df.empty:
            print("No tasks found.")
        else:
            print("\n📋 ALL TASKS:")
            print(df.to_string(index=False))
        return df
    except FileNotFoundError:
        print(f"❌ File not found at: {EXCEL_PATH}")
        return None

def read_tasks_by_status(status):
    """Read tasks filtered by status (e.g., 'To Do', 'In Progress', 'Done')."""
    try:
        df = pd.read_excel(EXCEL_PATH)
        filtered = df[df['Status'].str.lower() == status.lower()]
        if filtered.empty:
            print(f"No tasks with status: {status}")
        else:
            print(f"\n📋 TASKS WITH STATUS '{status.upper()}':")
            print(filtered.to_string(index=False))
        return filtered
    except FileNotFoundError:
        print(f"❌ File not found at: {EXCEL_PATH}")
        return None

# ── WRITE FUNCTIONS ─────────────────────────────────────────────

def add_task(task_name, priority="Medium", due_date="", status="To Do"):
    """Add a new task to the Excel file."""
    try:
        df = pd.read_excel(EXCEL_PATH)
    except FileNotFoundError:
        # Create new dataframe if file doesn't exist
        df = pd.DataFrame(columns=["Task Name", "Status", "Priority", "Due Date", "Created Date"])

    new_task = {
        "Task Name": task_name,
        "Status": status,
        "Priority": priority,
        "Due Date": due_date,
        "Created Date": datetime.today().strftime('%Y-%m-%d')
    }

    df = pd.concat([df, pd.DataFrame([new_task])], ignore_index=True)
    df.to_excel(EXCEL_PATH, index=False)
    print(f"✅ Task added: '{task_name}' with status '{status}'")

def update_task_status(task_name, new_status):
    """Update the status of an existing task."""
    try:
        df = pd.read_excel(EXCEL_PATH)
        mask = df['Task Name'].str.lower() == task_name.lower()

        if mask.any():
            df.loc[mask, 'Status'] = new_status
            df.to_excel(EXCEL_PATH, index=False)
            print(f"✅ Task '{task_name}' updated to '{new_status}'")
        else:
            print(f"❌ Task '{task_name}' not found.")
    except FileNotFoundError:
        print(f"❌ File not found at: {EXCEL_PATH}")

def delete_task(task_name):
    """Delete a task from the Excel file."""
    try:
        df = pd.read_excel(EXCEL_PATH)
        mask = df['Task Name'].str.lower() == task_name.lower()

        if mask.any():
            df = df[~mask]
            df.to_excel(EXCEL_PATH, index=False)
            print(f"🗑️ Task '{task_name}' deleted.")
        else:
            print(f"❌ Task '{task_name}' not found.")
    except FileNotFoundError:
        print(f"❌ File not found at: {EXCEL_PATH}")

# ── MAIN MENU ───────────────────────────────────────────────────

def main():
    print("\n🤖 TO-DO AGENT - Excel Task Manager")
    print("=====================================")
    while True:
        print("\nWhat would you like to do?")
        print("1. View all tasks")
        print("2. View tasks by status")
        print("3. Add a new task")
        print("4. Update task status")
        print("5. Delete a task")
        print("6. Exit")

        choice = input("\nEnter your choice (1-6): ").strip()

        if choice == "1":
            read_all_tasks()

        elif choice == "2":
            status = input("Enter status (To Do / In Progress / Done / Completed): ").strip()
            read_tasks_by_status(status)

        elif choice == "3":
            task_name = input("Enter task name: ").strip()
            priority = input("Enter priority (High / Medium / Low): ").strip() or "Medium"
            due_date = input("Enter due date (YYYY-MM-DD) or leave blank: ").strip()
            add_task(task_name, priority, due_date)

        elif choice == "4":
            task_name = input("Enter task name to update: ").strip()
            new_status = input("Enter new status (To Do / In Progress / Done / Completed): ").strip()
            update_task_status(task_name, new_status)

        elif choice == "5":
            task_name = input("Enter task name to delete: ").strip()
            delete_task(task_name)

        elif choice == "6":
            print("👋 Exiting To-Do Agent. Goodbye!")
            break

        else:
            print("❌ Invalid choice. Please enter 1-6.")

if __name__ == "__main__":
    main()

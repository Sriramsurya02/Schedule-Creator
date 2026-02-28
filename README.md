**Schedule-Creator**

**Features**
- Creates a new Excel workbook
- Generates a weekly layout (Monday–Sunday)
- Time slots from 09:00 to 17:00 in 30-minute intervals
- Alternating column layout for better readability
- Styled header (bold + centered)
- Thin borders applied to filled task cells
- Interactive CLI input for adding tasks
- Automatically calculates correct row and column placement

**Requirements**
- Python 3.x
- openpyxl

**How It Works**
*1. Workbook Creation*
- Creates a new Excel file
- Adds a worksheet titled "Schedule"
*2. Time Generation*
- Time slots are generated automatically:
09:00
09:30
10:00
...
16:30
*3. Layout Structure*
- Each day occupies two columns:
- One spacer column
- One task column
Days included:
Monday
Tuesday
Wednesday
Thursday
Friday
Saturday
Sunday
*4. Task Input Format*
- When running the script, enter tasks in the format:
day time task

Example:

Monday 09:00 Gym
Tuesday 13:30 Meeting with team
Friday 15:00 Study session

Type:

done

to finish input.

**Example Usage**

- Run the script:

python schedule.py

- Input example:

Enter tasks (day time task), or 'done' to finish:
Monday 09:00 Gym
Tuesday 13:30 Project Meeting
done

- Output:

Schedule saved to schedule.xlsx

The Excel file will be saved to:
D:\Code\Schedule\schedule.xlsx
You can modify the path variable in the script to change the save location.

**Customization**
- You can easily modify:
- Start and end times
- Time interval duration
- Column widths
- Styling (fonts, borders, alignment)
- File save path

**Limitations**
- Time must be entered in HH:MM (24-hour format)
- Day names must match exactly:
Monday–Sunday
- Existing files at the same path will be overwritten

**Future Improvements**
- GUI version (Tkinter or PyQt)
- Input validation for time ranges
- Color-coded tasks
- Automatic merging of adjacent cells
- Editable existing schedules instead of overwriting

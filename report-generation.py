import win32print
import win32ui
import mysql.connector

margin = 5  # Margin in pixels
tab_size = 60  # Adjust for desired text alignment


# Connect to MySQL database
conn = mysql.connector.connect(
    host='yourhost',  # Replace with your MySQL host
    user='youruser',  # Replace with your MySQL username
    password='yourpassword',  # Replace with your MySQL password
    database='yourdatabase'  # Replace with your database name
)

def print_text(text):
    printer_name = win32print.GetDefaultPrinter()
    hprinter = win32print.OpenPrinter(printer_name)

    hdc = win32ui.CreateDC()
    hdc.CreatePrinterDC(printer_name)
    hdc.StartDoc("Test doc")

    y = 0  # Start printing from the top of the page

    lines = text.splitlines()  # Split text into lines
    max_width = max(hdc.GetTextExtent(line)[0] for line in lines)  # Find the widest line
    line_spacing = int(1.5 * hdc.GetTextExtent(" ")[1])
    for line in lines:
        # Check if the line exceeds the available width and wrap it if needed
        if hdc.GetTextExtent(line)[0] > max_width:
            for word in line.split():
                # Wrap long words if they exceed the available width
                if hdc.GetTextExtent(word)[0] > max_width:
                    for sub_word in word.split():
                        hdc.TextOut(180, y+500, sub_word)
                        y += hdc.GetTextExtent(sub_word)[1]
                else:
                    hdc.TextOut(350, y+500, word)
                    y += hdc.GetTextExtent(word)[1]
                hdc.TextOut(180, y+500, " ")  # Add a space between wrapped words
        else:
            hdc.TextOut(350, y+300, line)
            y += hdc.GetTextExtent(line)[1]

    font = win32ui.CreateFont({
        "name": "Times New Roman",
        "height": 12,
        "weight": 400,
        "italic": 0,
    })

    old_font = hdc.SelectObject(font)

    hdc.EndDoc()
    hdc.DeleteDC()
    win32print.ClosePrinter(hprinter)



def fetch_data_from_database():
    # Connect to the database
   if conn.is_connected():
    print("Connected to MySQL database")
    cursor = conn.cursor()

    # Fetch data from the database
    cursor.execute("SELECT * FROM lg_info")
    data = cursor.fetchall()

    # Close the connection
    conn.close()

    return data


def print_python(data):
    # Generate report
        for row in data[:2]:
            print(f"{row[1]} {row[2]} {row[3]}")

def create_report(data):
    # Generate report
        for row in data[:2]: #adjust according to need
          #add your desired content here  
          text_to_print = f"""Employee Performance Evaluation 
  
Date: [Date]
  
Employee Name: {row[1]} {row[2]}
  
Introduction
  
This report serves as an evaluation of {row[1]}'s performance as a {row[3]} at XYZ 
Company. It aims to highlight her achievements, identify areas for improvement, and 
assess her overall impact on the organization. {row[1]} has been an integral part of
our team, consistently demonstrating dedication and proficiency in her role.
  
Performance Overview
  
{row[1]} has consistently displayed exemplary performance throughout her tenure at 
XYZ Company. Her commitment to delivering high-quality work and her collaborative 
approach have positioned her as a valuable asset to the team. She consistently exceeds
expectations and contributes significantly to the achievement of both departmental and
corporate objectives.
  
  
Strengths
  
Commitment and Reliability: 
{row[1]}'s dedication to her work and her reliability in meeting deadlines 
have been commendable. She consistently delivers superior results within the allocated 
time frames, contributing to the success of our projects.
  
Collaborative Spirit: 
{row[1]} actively engages in team discussions, provides valuable feedback, and 
works cohesively with her colleagues towards common goals. Her collaborative approach 
enhances team dynamics and fosters a positive working environment.
  
  
Areas for Improvement

Communication Skills: 
While {row[1]} excels in her technical abilities, there is room for improvement in her
communication skills. Enhancing her ability to articulate ideas and provide timely 
updates will promote better collaboration within the team and ensure clarity in project 
communication.
  
  
Recommendations
  
To support {row[1]}'s continued growth and development, the following recommendations 
are proposed:
  
Communication Skills Enhancement: {row[1]} is encouraged to participate in a 
communication skills enhancement program to refine her ability to convey ideas and 
provide status updates effectively. This will enable her to communicate more clearly 
with her team members and stakeholders.
  
  
Conclusion
  
In conclusion, {row[1]} has demonstrated exceptional performance as a {row[3]} at XYZ
Company. Her dedication, reliability, and collaborative spirit have contributed 
significantly to the success of our projects. While there are areas for improvement,
{row[1]}'s potential for advancement and her willingness to learn make her a valuable 
asset to the organization. With access to focused developmental programs, she is 
well-positioned to further excel in her role and make substantial contributions to the
company's prosperity."""
            print_text(text_to_print)

def main():
    # Fetch data from the database
    data = fetch_data_from_database()

    # Create report and print the Report
    print_python(data)
    create_report(data)

if __name__ == "__main__":
    main()

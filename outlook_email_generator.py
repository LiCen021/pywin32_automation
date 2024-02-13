import win32com.client as win32
import pandas as pd

# Create a new instance of the Outlook application
outlook = win32.Dispatch('outlook.application')

# Create a new email item
mail = outlook.CreateItem(0)

# Configure the email
# mail.To = 'recip@domain.com'  # Add recipient
mail.Subject = 'Email Subject Here'  # Add subject

# Define an HTML table
html_content_1 = """
<html>
<head>
<style>
  table, th, td {
    border: 1px solid black;
    border-collapse: collapse;
  }
  th, td {
    padding: 5px;
    text-align: left;
  }
</style>
</head>
<body>
<h2>Title Format 2</h2>
<h1>Title Format 1</h1>
<h1>Title Format 1</h1>
"""

# Sample DataFrame
data = {
    'Name': ['Alice', 'Bob', 'Charlie'],
    'Age': [25, 30, 35],
    'City': ['New York', 'Paris', 'London']
}
df = pd.DataFrame(data)

# Convert the DataFrame to an HTML table
html_content_2 = df.to_html() + '<br>'

# Sample DataFrame
data = {
    'Name': ['Alice23', 'Bob23', 'Charlie123'],
    'Age': [25, 30, 35],
    'City': ['New York', 'Paris', 'London']
}
df = pd.DataFrame(data)

# Convert the DataFrame to an HTML table
html_content_3 = df.to_html() + '<br>'

# Set the HTML body including the table
mail.HTMLBody = html_content_1 + html_content_2 + html_content_3

# To add an attachment uncomment the following line and specify the path to the file
# mail.Attachments.Add('path_to_attachment')
mail.Display(True)
# Send the email
# mail.Send()
# Alternatively, you could use mail.Display(True) to display the email window, allowing for manual review before sending.

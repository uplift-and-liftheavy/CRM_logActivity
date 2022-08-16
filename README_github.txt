# CRM_logActivity

# I work in outside sales. Part of my responsibility is to log my activity to a CRM, which is time consuming. This code automates that.
# The code looks at an Excel file, which has a column for the ContactID (unique identifier), Activity Description, Notes, and Start Date/Time.


# The code first asks the user to input how many activities he/she would like to log. The user inputs value, which is equal to the number of rows of data in Excel
# (minus 1 for header).
# Next, code navigates to my CRM via Chrome, navigates to the appropriate contactID, clicks on "Add Activity" under contact, swaps to the new window popup to select the 
# appropriate activity type, switches back to the main window, inputs the data read from Excel to the fields for Activity Description, Notes, and Start Date/Time & 
# End Date/Time (my activities never go longer than 1 day, so I re-use Start Date/Time for that field).
# When this is completed, the program clicks save, and returns to the home screen of CRM to begin the process again row-by-row, as many times as indicated by user.

# For security reasons, I removed the website and my username in the actual code was ommitted. 


# This code works. I hope it helps someone!


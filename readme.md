# Outlook Meeting Invite Setup and Run Guide

This guide will walk you through the setup and execution of this script to create and manage Outlook events. The script is designed to send meeting invites and feedback request emails based on data provided in a CSV file.

## Prerequisites

Before running the script, make sure you have the following prerequisites installed:

1. [Python](https://www.python.org/) (version 3.10 or higher)
2. [Microsoft 365 Subscription](https://www.microsoft.com/en-us/microsoft-365)

## Setup Instructions

### Step 1: Download the Code

1. If you are familiar with Git, you can clone this repository.

    ```bash
    git clone https://github.com/nishanthrj/Outlook-Meeting-Invite.git
    ```

    You can now skip to Step 2.

2. If you are not familiar with Git, follow these steps:
    - Look for the green "Code" button on the top.
    - Click on "Code" and select "Download Zip" from the dropdown menu.
    - Save the downloaded ZIP file to a location on your computer.
    - Locate the downloaded ZIP file.
    - Right-click on the ZIP file.
    - From the context menu, select "Extract All."
    - Choose a folder for extraction and click "Extract."
    - Open the extracted folder to access its contents.

### Step 2: Setup the Environment

1. Open the project folder.
2. Look for a file named `setup.bat`.
3. Double-click on `setup.bat` to run it. This script will automatically set up the necessary environment and install required packages.

### Step 3: Microsoft Graph API Setup

1. Visit [Microsoft Entra](https://entra.microsoft.com/).
2. Sign in using the Microsoft account associated with your Microsoft 365 Subscription.
3. In the top search bar, type "App registrations" and select it.
4. Click on "New Registration."
5. Enter a name for the application (e.g., Outlook Meeting Invite) and click "Register."
6. Once registered, copy the "Application (client) ID" and "Directory (tenant) ID" from the overview page. Keep this information somewhere.
7. Navigate to "Certificates & Secrets."
8. Click on "New client secret," provide a description, set an expiry, and create the secret.
9. Copy the generated secret and store it too. (NOTE: You must copy the value right away since you cannot see it later. If you lose this value, delete this secret and create a new one.)
10. In "API permissions," click on "Add a permission" > "Microsoft Graph" > "Application permissions."
11. Search for "Calendars.ReadWrite" and "Mail.Send." Check the respective boxes for these permissions and click "Add permissions."
12. Remove the default "User.Read" permission by clicking the three dots (...) in the end and selecting "Remove permission."
13. Click on "Grant admin consent for MSFT" to allow permissions.
14. Go back to the search bar, type "Users" and find your account.
15. Click on your Display name to access your user page.
16. Copy the "Object ID" and store it along with the other values.

### Step 4: Setup Environment Variables

1. Go back to the project folder.
2. Right-click in the folder, select "New" and create a new text document.
3. Name the file ".env" (without the .txt extension).
4. Open the file and input the following, fill the empty quotes with the values you copied earlier.

    ```bash
    CLIENT_ID=""
    CLIENT_SECRET=""
    TENANT_ID=""
    OBJECT_ID=""
    ```

5. Once done, save the file and close it.

### Step 5: Run the Program

1. Return to the project folder.
2. Look for a file named `run.bat`.
3. Double-click on `run.bat` to automatically execute the code.

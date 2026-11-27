
# Gmail Labeler

[English](README.md) | [Italiano](README.it.md)

Gmail Labeler is a Windows application that allows you to apply Gmail labels to large numbers of emails based on their dates.

The application helps you organize your messages without having to select them manually one by one.

## Features

- **Date range labeling:** applies a label to emails between two dates, including both the start and end dates.
- **Labeling emails before a date:** selects messages earlier than the selected date, excluding that day.
- **Labeling emails after a date:** selects messages from the selected date onward, including that day.
- **Automatic label creation:** if the specified label does not exist, it is created automatically; if it already exists, it is reused.
- **Bulk email labeling:** applies a single label to multiple messages in one operation.
- **Confirmation before applying labels:** displays the number of matching emails and asks for confirmation before proceeding.
- **Processing cancellation:** allows you to stop an operation while it is running.
- **Read and unread emails:** includes both read and unread messages.
- **Archived emails:** includes archived messages that match the search criteria.
- **Multilingual interface:** Italian and English.
- **Visual themes:** Light, Dark, Ocean, Forest, Sunset and Purple.

Gmail Labeler adds a label to messages. It does not delete, move or modify the content of emails.

The application uses Gmail's standard search functionality, which normally excludes messages in Spam and Trash.

## Project structure

GmailLabeler/
- src/
  - Setup-GmailLabeler.ps1
  - Build-GmailLabeler.ps1
- assets/
  - icon.ico
- Compilator.bat
- LICENSE
- README.it.md
- README.md

### File descriptions

- `Setup-GmailLabeler.ps1`: contains the application code and installation wizard. The Python application code is embedded directly in this script.
- `Build-GmailLabeler.ps1`: handles the compilation of the PowerShell installer script into a Windows executable using PS2EXE.
- `icon.ico`: icon used by the application and installer.
- `Compilator.bat`: starts the compilation process.

## Requirements

### For users

- 64-bit Windows.
- Internet connection during installation.
- Administrator permissions if required by the selected installation folder.

The installer checks whether a compatible version of Python is already installed. If not, it can install Python automatically.

### For developers

- Windows.
- Windows PowerShell.
- PS2EXE, which can be installed automatically during the compilation process.
- Internet connection to download PS2EXE if necessary.

## Google Cloud setup

To allow Gmail Labeler to access your mailbox, you need to configure a Google Cloud project and enable the Gmail API.

1. Open [Google Cloud Console](https://console.cloud.google.com/).
2. Create a new project or select an existing one.
3. Enable **Gmail API**.
4. Configure the OAuth consent screen.
5. Create an OAuth client ID and select **Desktop app** as the application type.
6. Download the credentials JSON file.

The downloaded file must be selected when requested during the application's first startup.

## Permissions and security

Gmail Labeler uses the OAuth `gmail.modify` permission scope, which is required to apply labels to messages.

Credentials are stored locally on the user's computer:

- `credentials.json`: OAuth credentials file.
- `token.json`: authentication token generated during authorization.

These files are stored in:

%APPDATA%\Gmail Labeler\credentials

Application preferences, such as the selected theme and language, are stored in:

%APPDATA%\Gmail Labeler\user_data

Credentials and authentication tokens must never be published.

## Installation

1. Run `Setup-GmailLabeler.exe`.
2. Follow the installation wizard.
3. Choose the installation folder.
4. Wait for the installation to complete.

The installer automatically prepares the required environment:

- checks whether Python is installed;
- installs Python if necessary;
- checks for and installs the required packages;
- builds the application using PyInstaller;
- creates the uninstaller;
- copies the application files to the selected folder;
- registers the application in Windows installed apps;
- offers the option to create a Desktop shortcut.

An Internet connection is required to download Python and the necessary packages when they are not already available.

## Usage

On the first startup, you will be asked to authorize access to your Google account.

After authentication, you can choose one of the three search modes:

### 1. Date range

Select a start date and an end date.

Both dates are included.

### 2. Emails before a date

Select a date.

Emails dated before the selected day are included, while the selected day itself is excluded.

### 3. Emails after a date

Select a date.

Emails dated from the selected day onward are included, including the selected day.

After choosing the search criteria, enter the name of the label you want to apply.

The application finds the matching messages and displays a confirmation prompt before proceeding.

If the label does not exist, it is created automatically.

## Message processing

To handle large numbers of emails, the application uses paginated searches and batch operations.

- Search retrieves up to 500 results per page.
- Modification operations can process up to 1,000 messages per batch.
- Processing can be interrupted.

## Email export

Gmail Labeler can also be used to prepare emails for export through Google Takeout.

You can apply a label to the messages you want to export and then use Google Takeout to export the messages associated with that label in MBOX format.

## Interface and preferences

The application supports Italian and English and includes six visual themes:

- Light
- Dark
- Ocean
- Forest
- Sunset
- Purple

The language is initially detected based on the Windows language settings. The selected theme and language preferences can be saved and retained between application sessions.

## Building the installer

The repository includes the scripts required to generate `Setup-GmailLabeler.exe`.

To start the compilation process:

1. Open the project's root folder.
2. Run `Compilator.bat`.
3. If PS2EXE is not installed, the script will ask whether you want to install it.
4. Confirm the installation if requested.
5. Wait for the compilation to complete.

The generated file, `Setup-GmailLabeler.exe`, will be created in the project's root folder.

The compilation uses PS2EXE with the following settings:

- executable without a console window;
- STA mode;
- administrator privileges required;
- x64 architecture;
- custom icon;
- executable title, description, company, product and version information.

**Note:** the PS2EXE compilation generates the installer executable. When the end user runs it, the installation wizard prepares the required environment and builds the Gmail Labeler application on that computer using PyInstaller.

## Uninstallation

The application includes an uninstaller.

During the removal process, the user can choose whether to also delete Google credentials and user data, including saved theme and language preferences.

If the user chooses not to delete them, these files remain on the computer.

## GitHub and distribution

GitHub is used to host the source code and distribute GmailLabeler.

The repository includes:

- PowerShell build and installation scripts;
- application icon;
- README files;
- LICENSE file.

The compiled installer (`Setup-GmailLabeler.exe`) is distributed through GitHub Releases.

User credentials and authentication tokens are stored locally on the user's computer and are not included in the repository or installer.
GitHub can be used to publish and distribute the project.

## License

This project is licensed under the MIT License.
See the [LICENSE](LICENSE) file for details.

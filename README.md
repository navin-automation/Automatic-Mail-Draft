# Automatic Mail Draft System

A powerful VBA/Macro-based automation tool that automatically creates Outlook/Exchange email drafts with attachments based on a dynamic configuration file (`.csv`). It streamlines bulk mailing workflows by mapping specific recipients, cc addresses, custom greetings, and relevant local reports seamlessly.

## 🚀 Features

- **Multi-Account Support**: Automatically binds and authenticates drafts using a specific sender address (`support@techroversolutions.com`).
- **Dynamic Routing**: Parses configuration data from `HotelContacts.csv` to map custom recipients (`To` and `CC` fields) dynamically.
- **Automated Attachments**: Scans the local `Downloads` folder to find and attach exact matching documents (Daily, MTD, and Revenue reports).
- **Customized Greetings**: Generates automated, personalized greetings for each dynamic entity or group.
- **Flexible Rules**: Includes built-in custom logic branching to handle multi-part report distribution rules for specific setups.

## 📁 Repository Structure

- `WORKING_ON_DRAFT.bas`: Core VBA automation script handling Outlook application binding, folder traversal, file attachment checks, and draft creation.
- `HotelContacts.csv`: Sample layout mapping configuration file for routing logic, emails, and greeting variables.

## 🛠️ How It Works

1. The script reads the configuration grid row-by-row.
2. It fetches the required file paths from the local machine's `Downloads` directory.
3. It maps the designated Primary (PM) and Secondary (GM) contacts.
4. It compiles and saves the generated draft safely inside the Outlook/Exchange Drafts folder without triggering accidental sends.

## 💻 Tech Stack

- **Language:** VBA (Visual Basic for Applications)
- **Environment:** Microsoft Outlook / Excel
- **Data Source:** CSV (Comma-Separated Values)

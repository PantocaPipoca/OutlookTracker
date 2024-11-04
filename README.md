# 📅 OutlookTracker

**Track Your Activities with Outlook Integration**

OutlookTracker is a personal project designed to log daily activities in an Outlook calendar. With weekly and monthly insights, it helps you analyze how you spend your time to focus on tasks that matter.

---

### 📝 Overview
Are you often wondering where your time went? Whether it's work, studying, or gaming, OutlookTracker captures everything and organizes it in your Outlook calendar. Spend time more consciously with automatic tracking, goal setting, and productivity insights.

---

## 🚀 Features
The following list shows the project's feature roadmap. Completed features are marked as ~crossed out~ (Feature list is not ordered in any way):

- [x] **Microsoft Authentication**: Secure login using Microsoft OAuth and Graph API.
- [x] **Calendar Event Logging**: Automatically logs categorized activities in Outlook Calendar.
- [x] **Active Window Tracking**: Tracks active apps and categorizes activities.
- [x] **Dynamic Event Updates**: Logs time spent on activities by updating calendar entries.
- [ ] **Custom User Interface**: Easy-to-use GUI for managing settings and tracking.
- [ ] **Monthly Statistics Generation**: Visualize monthly activities with time charts.
- [ ] **Multi-Device Tracking**: Expands tracking to mobile and tablets.
- [ ] **Cross-Platform Compatibility**: Support for iOS and Linux.
- [ ] **Enhanced Activity Categorization**: Add custom categories (e.g., reading, exercise).
- [ ] **Background Execution**: Runs in the background for continuous tracking.
- [ ] **User Notification System**: Alerts for authentication, tracking, and actions.
- [ ] **Configurable Settings**: Allows customization of intervals and categories.
- [ ] **Location-Based Tracking**: Logs activities with location context.
- [ ] **Automated Report Generation**: Export weekly/monthly summaries as downloadable reports.
- [ ] **Goal Setting**: Track progress with time-based activity goals.
- [ ] **Productivity Insights**: Provides insights on time usage patterns.
- [ ] **Smart Alerts**: Alerts for extended activities, like screen time breaks.
- [ ] **Data Backup and Sync**: Backup and sync data across devices.
- [ ] **Offline Mode**: Tracks offline, syncing when connected.
- [ ] **Customizable GUI Themes**: Offers light, dark, and other GUI themes.
- [ ] **Data Privacy Controls**: Lets users review, manage, and delete activity logs.

---

## 🛠️ Technologies Used
- **Python**: Core language for tracking, logging, and processing.
- **Microsoft Graph API**: Manages Outlook calendar events.
- **MSAL (Microsoft Authentication Library)**: Ensures secure authentication.

---

## 🎉 Getting Started

### ✅ Prerequisites
- **Microsoft 365 Account**: Needed to log activities in Outlook.
- **Python 3.x**: Ensure Python and pip are installed.
- **API Credentials**: Set up `CLIENT_ID` and `CLIENT_SECRET` in the Azure Portal.

### 🛠️ Installation

1. **Clone the repository**:
    ```bash
    git clone https://github.com/PantocaPipoca/OutlookTracker.git
    ```

2. **Install dependencies**:
    ```bash
    pip install -r requirements.txt
    ```

3. **Configure your environment variables**:
    Create a `.env` file:
    ```makefile
    OutlookAutom_CLIENT=your_client_id
    OutlookAutom_SECRET=your_client_secret
    ```

4. **Run the application**:
    ```bash
    python main.py
    ```

---

## 📊 Usage

1. **Start Tracking**: Use the GUI to input your user ID and start activity tracking.
2. **Review Activities**: Check your Outlook calendar for activity logs and time allocation.

---

## 🧩 Contributing
Feel free to submit feature requests, bug reports, or improvements via [Issues](https://github.com/PantocaPipoca/OutlookTracker/issues) on GitHub.

## 📜 License
This project is licensed under the MIT License.
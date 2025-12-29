# 🛡️ Community Management Guild Bot

A powerful automated Telegram bot designed to streamline guild management and performance tracking for **Lords Mobile**. This bot processes complex Excel data to generate insightful reports, track player performance, and automate administrative tasks.

Built with **Python**, **Flask**, **Pandas**, and the **Telegram Bot API**.

## 🚀 Key Features

* **📊 Automated Data Analysis**: instantly processes raw Excel (.xlsx) files to calculate player statistics.
* **🎨 Smart Formatting**: Generates a beautifully formatted, color-coded Excel report (identifying Top 3 ranks, low performers, and specific issues).
* **🔍 Advanced Filtering**: Automatically detects and filters out players who have been in the guild for less than 4 days or have 0% participation.
* **📈 Performance Tracking**: Commands to instantly view Top 10, Bottom 10, and Zero-Hunt members.
* **⚡ Webhook Integration**: Deployed using Flask and asyncio for real-time, lag-free responses.
* **🛠 Guild Management Tools**:
    * Search for specific players (`/i`).
    * Flag issues with players (`/problem`).
    * Manually include or exclude members from reports.

## 🛠️ Tech Stack

* **Language:** Python 3.10+
* **Core Libraries:**
    * `python-telegram-bot` (Async API)
    * `pandas` (Data manipulation & analysis)
    * `openpyxl` (Excel formatting & styling)
    * `flask` (Webhook server)
* **Deployment:** PythonAnywhere (WSGI)

## 🤖 Bot Commands

| Command | Description |
| :--- | :--- |
| `/summary` | View a comprehensive weekly summary (Total points, Avg Goal %, Top/Lowest Hunter). |
| `/top10` | Show the top 10 performing players. |
| `/bottom10` | Show the bottom 10 performing players. |
| `/0hunt` | List all members with 0 hunts or 0% goal completion. |
| `/i [Name]` | Search for a specific player's stats and issues. |
| `/problem [Name] [Issue]` | Log an issue/remark for a specific player (e.g., "Busy", "Sick"). |
| `/generate` | Generate and download the updated, color-coded Excel report. |
| `/filteredplayers` | View list of players automatically removed (e.g., <4 days in guild). |
| `/exclude [Name]` | Manually remove a player from the report. |
| `/include [Name]` | Manually add a previously filtered player back into the report. |
| `/clear` | Clear all stored data from memory. |

## ⚙️ How It Works

1.  **Upload:** An admin uploads a raw data `.xlsx` file to the bot.
2.  **Processing:** The bot uses `pandas` to clean the data, calculate "Days in Guild", and filter out ineligible players.
3.  **Interaction:** Admins can query the data or tag players with issues using bot commands.
4.  **Reporting:** The bot uses `openpyxl` to generate a final report with:
    * **Gold/Silver/Bronze** highlighting for top ranks.
    * **Color scales** (Green to Red) based on goal completion percentages.
    * **Visual cues** for specific rows (e.g., "Kidnape me" / Bank accounts).

## 🔧 Installation & Setup

1.  **Clone the repository**
    ```bash
    git clone [https://github.com/Shreyash9595/Community-management-TG-Bot.git](https://github.com/Shreyash9595/Community-management-TG-Bot.git)
    cd Community-management-TG-Bot
    ```

2.  **Install dependencies**
    ```bash
    pip install -r requirements.txt
    ```

3.  **Environment Variables**
    Create a `.env` file and add your Telegram Token:
    ```env
    TELEGRAM_TOKEN=your_bot_token_here
    ```

4.  **Run Local (Polling Mode)**
    *(For testing without webhook)*
    ```bash
    python main.py
    ```

## 🔒 Security Note
This repository does not contain the `WSGI File.txt` or `.env` file to protect sensitive API tokens. Configuration is handled via environment variables in the deployment environment.

---
*Created by [Shreyash Mungase](https://github.com/Shreyash9595)*

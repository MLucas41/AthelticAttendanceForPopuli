# Athletic Attendance Automation for Populi

A Python application designed to automate student-athlete attendance management by integrating with the Populi API.

## Overview

Managing student-athlete absences can be a time-consuming process for university athletic departments. When athletes travel for competitions, their absences need to be communicated to professors and accurately reflected in the university's student information system.

This project aims to streamline that process by automating attendance updates and professor notifications using the Populi API.

## Features

- **Student Data Retrieval:** Retrieves student information through the Populi API.
- **Athlete Management:** Identifies student-athletes and manages relevant roster information.
- **Schedule Integration:** Uses academic schedules and athletic travel schedules to identify class conflicts.
- **Attendance Automation:** Automates attendance updates for students missing class due to athletic events.
- **Professor Notifications:** Generates email notifications for professors regarding student-athlete absences.
- **User Interface:** Provides an interface for reviewing and managing attendance automation.

## Technologies Used

| Technology | Purpose |
|---|---|
| Python | Core application logic |
| Populi API | Student information and attendance integration |
| Tkinter | Desktop user interface |
| Requests | HTTP requests and API communication |
| python-dotenv | Environment variable management |
| uv | Python package and project management |

## Project Structure

```text
AthleticAttendanceForPopuli/
├── .venv/
├── .env
├── .gitignore
├── pyproject.toml
├── uv.lock
├── app.py
├── getStudentDatabase.py
└── README.md
```

*Update this structure to reflect your actual project files.*

## Installation

### Prerequisites

- Python 3.14+
- uv
- Access to a Populi API account with appropriate permissions

### 1. Clone the repository

```bash
git clone https://github.com/YOUR_USERNAME/AthleticAttendanceForPopuli.git
cd AthleticAttendanceForPopuli
```

### 2. Install dependencies

```bash
uv sync
```

### 3. Configure environment variables

Create a `.env` file in the project root:

```env
API_KEY=your_populi_api_key
```

**Important:** Never commit your `.env` file or expose API credentials.

### 4. Run the application

```bash
uv run app.py
```

## Configuration

The application requires valid Populi API credentials and appropriate API permissions.

API credentials should be stored in environment variables rather than hardcoded into the source code.

## Future Improvements

- Expand schedule synchronization and conflict detection.
- Improve error handling and API response validation.
- Add additional attendance reporting capabilities.
- Improve user experience and workflow efficiency.
- Expand support for additional athletic teams.

## Author

Developed by Luke Mertlich.

Computer Science Student | Corban University
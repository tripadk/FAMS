# Faculty Achievement Management System Backend

A Flask-based backend for managing faculty achievements using SQLite database.

## Setup

1. Install dependencies:
   ```
   pip install -r requirements.txt
   ```

2. Run the application:
   ```
   python app.py
   ```

The database `database.db` will be automatically created with the required tables when the app starts.

## Database Tables

- **Faculty**: Stores faculty information
- **Achievements**: Stores achievement records linked to faculty
- **Admin**: Stores admin user information

## Initialization

The `init_db()` function in `app.py` initializes the database and creates all tables automatically on app startup.
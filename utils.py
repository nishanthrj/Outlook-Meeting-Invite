import hashlib
from datetime import datetime
from dateutil import parser


def generate_group_id(input_string: str) -> str:
    """
    Generate a unique group ID.

    Args:
        input_string (str): The input string to be hashed.

    Returns:
        str: The hashed string (group ID).
    """
    sha256_hash = hashlib.sha256()
    sha256_hash.update(input_string.encode("utf-8"))
    hashed_string = sha256_hash.hexdigest()

    return hashed_string


def get_recurrence_pattern(occurrence: str, start_date: str) -> dict | None:
    """
    Get a recurrence pattern based on the occurrence.

    Parameters:
        occurrence (str): The recurrence type, e.g., "daily" or "weekly".
        start_date (str): The start date for the meeting.

    Returns:
        dict : The recurrence pattern.
    """
    if occurrence == "daily":
        return {
            "type": "weekly",
            "interval": 1,
            "daysOfWeek": ["monday", "tuesday", "wednesday", "thursday", "friday"],
        }
    elif occurrence == "weekly":
        return {
            "type": "weekly",
            "interval": 1,
            "daysOfWeek": [
                datetime.strptime(start_date, "%Y-%m-%d").strftime("%A").lower()
            ],
        }
    else:
        return None


def format_date(date: str) -> str:
    """
    Format date to match the format required by Microsoft Graph API.

    Args:
        date (str): Meeting date

    Returns:
        str: Formatted date
    """
    parsed_date = parser.parse(date)
    output_date = parsed_date.strftime("%Y-%m-%d")
    return output_date

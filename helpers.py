import time
from googleapiclient.errors import HttpError


def retry_operation(func, retries=3, delay=2, *args, **kwargs):
    """
    Helper function to retry an operation in case of failure.

    Args:
        func (callable): The function to retry.
        retries (int): The number of times to retry before failing.
        delay (int): Delay in seconds between retries.
        *args: Arguments for the function.
        **kwargs: Keyword arguments for the function.

    Returns:
        Any: The return value of the function, or None if it failed.
    """
    attempt = 0
    while attempt < retries:
        try:
            return func(*args, **kwargs)
        except HttpError as e:
            print(f"Retry {attempt + 1}/{retries}: Error {e.status_code}: {e.reason}")
            attempt += 1
            if attempt < retries:
                time.sleep(delay)
                delay *= 2  # Exponential backoff
    print(f"Operation failed after {retries} attempts.")
    return None

"""Connection to the desktop Outlook MAPI namespace."""

import win32com.client

def get_outlook_namespace():
    """
    Connects to classic Outlook desktop.
    """
    outlook = win32com.client.Dispatch(
        "Outlook.Application"
    )

    return outlook.GetNamespace("MAPI")

"""Feedback configuration for Case Conference Builder exports.

Replace these URLs with your residency's REDCap or Qualtrics feedback links.
"""

THANK_YOU_TITLE = "Thank you for participating"
THANK_YOU_MESSAGE = (
    "Please complete the brief feedback form so we can keep case conference "
    "clinically relevant, interactive, and useful."
)
FEEDBACK_INSTRUCTION = "Scan the QR code or use the link to provide feedback."

# QR code points to the full survey URL.
FEEDBACK_QR_URL = "https://redcap.ctsi.psu.edu/surveys/?s=REPLACE_ME"

# Display link can be a short link for laptop users.
FEEDBACK_DISPLAY_URL = "https://redcap.link/replace_me"

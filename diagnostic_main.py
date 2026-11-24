"""Compatibility entry point for the canonical Pharma RFP monitor.

The production implementation lives in main.py. Keeping this wrapper prevents
the older diagnostic runner from bypassing its URL, evidence, deadline, email,
and persistence safeguards.
"""

from main import main


if __name__ == "__main__":
    main()

import os

# Use the same @func/@arg decorators as xlwings Lite. Without this, xlwings replaces
# them with no-ops on macOS and Linux. Must be set before xlwings is imported.
os.environ["XLWINGS_ON_SERVER"] = "true"

import os
import glob

# Remove all generated event files in the current directory
for file in glob.glob("events*"):
    os.remove(file)
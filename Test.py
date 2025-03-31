import re
import Tools

# Input string
text = "Bonjour 123 monde Python3 est génial"

# Find all words that contain only letters (no digits)
words = re.findall(r'\b[a-zA-Z]+\b', text)[:2]

# Get the first two words if they exist
first_two_words = words

# Display the result
print(first_two_words)  # Output: ['Bonjour', 'monde']

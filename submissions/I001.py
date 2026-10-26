import re

def is_palindrome(text):
    s = re.sub(r"[^a-z0-9]", "", text.lower())
    return s == s[::-1]

assert is_palindrome("Racecar")
assert is_palindrome("A man, a plan, a canal: Panama")
assert not is_palindrome("python")

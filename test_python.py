# TreeSitterLexer PoC Test File
# This file tests Python syntax highlighting via tree-sitter

import os
import sys
from collections import defaultdict

# Constants
MAX_RETRIES = 3
PI = 3.14159
DEBUG = True

class DataProcessor:
    """A sample class for testing tree-sitter highlighting."""
    
    def __init__(self, name: str, count: int = 0):
        self.name = name
        self.count = count
        self._cache = {}
    
    def process(self, data: list) -> dict:
        """Process the input data and return results."""
        result = {}
        for i, item in enumerate(data):
            if item is None:
                continue
            elif isinstance(item, str):
                result[item] = len(item)
            else:
                result[str(item)] = item * 2
        return result
    
    @staticmethod
    def validate(value):
        if not value:
            raise ValueError("Value cannot be empty")
        return True

async def fetch_data(url: str) -> bytes:
    """Async function for testing coroutine highlighting."""
    try:
        response = await some_http_lib.get(url)
        return response.content
    except Exception as e:
        print(f"Error fetching {url}: {e}")
        return b""

# Builtin function calls
numbers = list(range(10))
total = sum(numbers)
filtered = filter(lambda x: x > 5, numbers)
mapped = map(str, numbers)

# Various literals
name = "hello world"
raw = r"raw\nstring"
multiline = """
This is a
multiline string
"""
byte_str = b"bytes"
fstring = f"Result: {total + 1}"
hex_num = 0xFF
oct_num = 0o77
bin_num = 0b1010
float_num = 1.5e-3
complex_num = 3 + 4j

# Control flow
for x in range(10):
    if x % 2 == 0:
        print(x)
    elif x == 7:
        break
    else:
        continue

while True:
    pass

# Pattern matching (Python 3.10+)
match total:
    case 0:
        print("zero")
    case n if n > 0:
        print(f"positive: {n}")

# Boolean and None
flag = True
other = False
nothing = None

if flag is not None and flag is True:
    print("flag is set")

# Operators
result = (10 + 20) * 3 // 4
bitwise = 0xFF & 0x0F | 0x10 ^ 0x01
shifted = 1 << 4
comparison = 10 >= 5 and 3 <= 7 or not False

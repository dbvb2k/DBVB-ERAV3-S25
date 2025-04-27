# basic import 
from mcp.server.fastmcp import FastMCP, Image
from mcp.server.fastmcp.prompts import base
from mcp.types import TextContent
from mcp import types
from PIL import Image as PILImage
import math
import sys
from pywinauto.application import Application
import win32gui
import win32con
import time
from win32api import GetSystemMetrics
import logging
import os
from datetime import datetime
from dotenv import load_dotenv



# Load environment variables
load_dotenv()

# Configure logging
log_dir = "logs"
if not os.path.exists(log_dir):
    os.makedirs(log_dir)

timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
log_file = os.path.join(log_dir, f"debug_{timestamp}.log")

# Configure logging
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_file),
        logging.StreamHandler()
    ]
)

logger = logging.getLogger(__name__)

# instantiate an MCP server client
mcp = FastMCP("Calculator")

# DEFINE TOOLS

#addition tool
@mcp.tool()
def add(a: int, b: int) -> int:
    """Add two numbers"""
    logger.info("CALLED: add(a: int, b: int) -> int:")
    return int(a + b)

@mcp.tool()
def add_list(l: list) -> int:
    """Add all numbers in a list"""
    logger.info("CALLED: add(l: list) -> int:")
    return sum(l)

# subtraction tool
@mcp.tool()
def subtract(a: int, b: int) -> int:
    """Subtract two numbers"""
    logger.info("CALLED: subtract(a: int, b: int) -> int:")
    return int(a - b)

# multiplication tool
@mcp.tool()
def multiply(a: int, b: int) -> int:
    """Multiply two numbers"""
    logger.info("CALLED: multiply(a: int, b: int) -> int:")
    return int(a * b)

#  division tool
@mcp.tool() 
def divide(a: int, b: int) -> float:
    """Divide two numbers"""
    logger.info("CALLED: divide(a: int, b: int) -> float:")
    return float(a / b)

# power tool
@mcp.tool()
def power(a: int, b: int) -> int:
    """Power of two numbers"""
    logger.info("CALLED: power(a: int, b: int) -> int:")
    return int(a ** b)

# square root tool
@mcp.tool()
def sqrt(a: int) -> float:
    """Square root of a number"""
    logger.info("CALLED: sqrt(a: int) -> float:")
    return float(a ** 0.5)

# cube root tool
@mcp.tool()
def cbrt(a: int) -> float:
    """Cube root of a number"""
    logger.info("CALLED: cbrt(a: int) -> float:")
    return float(a ** (1/3))

# factorial tool
@mcp.tool()
def factorial(a: int) -> int:
    """factorial of a number"""
    logger.info("CALLED: factorial(a: int) -> int:")
    return int(math.factorial(a))

# log tool
@mcp.tool()
def log(a: int) -> float:
    """log of a number"""
    logger.info("CALLED: log(a: int) -> float:")
    return float(math.log(a))

# remainder tool
@mcp.tool()
def remainder(a: int, b: int) -> int:
    """remainder of two numbers divison"""
    logger.info("CALLED: remainder(a: int, b: int) -> int:")
    return int(a % b)

# sin tool
@mcp.tool()
def sin(a: int) -> float:
    """sin of a number"""
    logger.info("CALLED: sin(a: int) -> float:")
    return float(math.sin(a))

# cos tool
@mcp.tool()
def cos(a: int) -> float:
    """cos of a number"""
    logger.info("CALLED: cos(a: int) -> float:")
    return float(math.cos(a))

# tan tool
@mcp.tool()
def tan(a: int) -> float:
    """tan of a number"""
    logger.info("CALLED: tan(a: int) -> float:")
    return float(math.tan(a))

# mine tool
@mcp.tool()
def mine(a: int, b: int) -> int:
    """special mining tool"""
    logger.info("CALLED: mine(a: int, b: int) -> int:")
    return int(a - b - b)

@mcp.tool()
def create_thumbnail(image_path: str) -> Image:
    """Create a thumbnail from an image"""
    logger.info("CALLED: create_thumbnail(image_path: str) -> Image:")
    img = PILImage.open(image_path)
    img.thumbnail((100, 100))
    return Image(data=img.tobytes(), format="png")

@mcp.tool()
def strings_to_chars_to_int(string: str) -> list[int]:
    """Return the ASCII values of the characters in a word"""
    logger.info("CALLED: strings_to_chars_to_int(string: str) -> list[int]:")
    return [int(ord(char)) for char in string]

@mcp.tool()
def int_list_to_exponential_sum(int_list: list) -> float:
    """Return sum of exponentials of numbers in a list"""
    logger.info("CALLED: int_list_to_exponential_sum(int_list: list) -> float:")
    return sum(math.exp(i) for i in int_list)

@mcp.tool()
def fibonacci_numbers(n: int) -> list:
    """Return the first n Fibonacci Numbers"""
    logger.info("CALLED: fibonacci_numbers(n: int) -> list:")
    if n <= 0:
        return []
    fib_sequence = [0, 1]
    for _ in range(2, n):
        fib_sequence.append(fib_sequence[-1] + fib_sequence[-2])
    return fib_sequence[:n]


@mcp.tool()
async def draw_rectangle(x1: int, y1: int, x2: int, y2: int) -> dict:
    """Draw a rectangle in Paint from (x1,y1) to (x2,y2)"""
    global paint_app
    try:
        if not paint_app:
            return {"content": [TextContent(type="text", text="Paint is not open. Please call open_paint first.")]}
        
        logger.debug("Starting rectangle drawing operation")
        paint_window = paint_app.window(class_name='MSPaintApp')
        
        # Ensure Paint window is active
        if not paint_window.has_focus():
            paint_window.set_focus()
            time.sleep(1)
            
        # Click Rectangle tool
        logger.debug("Selecting rectangle tool")
        paint_window.click_input(coords=(445, 70))
        time.sleep(1)
        
        # Get canvas area
        canvas = paint_window.child_window(class_name='MSPaintView')
        
        # Log canvas position and size for debugging
        canvas_rect = canvas.rectangle()
        logger.debug(f"Canvas Rectangle: {canvas_rect}")
        
        # Draw rectangle with logging
        logger.debug(f"Drawing rectangle from ({x1},{y1}) to ({x2},{y2})")

        # Draw rectangle - coordinates should be relative to the Paint window
        logger.debug(f"Clicking at: ({x1}, {y1})")
        canvas.click_input(coords=(x1, y1))
        time.sleep(1)

        logger.debug(f"Pressing mouse at: ({x1}, {y1})")
        canvas.press_mouse_input(coords=(x1, y1))
        time.sleep(1)

        logger.debug(f"Releasing mouse at: ({x2}, {y2})")
        canvas.release_mouse_input(coords=(x2, y2))
        time.sleep(1)  

        # logger.debug(f"Clicking at: ({x2}, {y2+40})")
        # canvas.click_input(coords=(x2, y2+40))
        # time.sleep(1)      

        return {
            "content": [TextContent(type="text", text=f"Rectangle drawn from ({x1},{y1}) to ({x2},{y2})")]
        }
    except Exception as e:
        logger.error(f"Error in draw_rectangle: {str(e)}")
        return {"content": [TextContent(type="text", text=f"Error drawing rectangle: {str(e)}")]}

@mcp.tool()
async def add_text_in_paint(text: str) -> dict:
    """Add text in Paint"""
    global paint_app
    try:
        if not paint_app:
            return {"content": [TextContent(type="text", text="Paint is not open. Please call open_paint first.")]}
        
        logger.debug("Starting text addition operation")
        paint_window = paint_app.window(class_name='MSPaintApp')
        
        # Ensure Paint window is active
        if not paint_window.has_focus():
            paint_window.set_focus()
            time.sleep(0.5)
        
        # Select green color
        logger.debug("Selecting green color")
        paint_window.click_input(coords=(895, 61))
        time.sleep(0.5)
        
        # Select Text tool
        logger.debug("Selecting text tool")
        paint_window.click_input(coords=(290, 70))
        time.sleep(0.5)
        
        # Get canvas
        canvas = paint_window.child_window(class_name='MSPaintView')
        
        # Click to start typing (inside rectangle)
        text_x, text_y = 500, 300  # Adjusted coordinates
        logger.debug(f"Clicking for text at ({text_x}, {text_y})")
        canvas.click_input(coords=(text_x, text_y))
        time.sleep(0.5)
        
        # Type text
        logger.debug(f"Typing text: '{text}'")
        paint_window.type_keys(text, with_spaces=True)
        time.sleep(0.5)
        
        # Click outside to finish
        canvas.click_input(coords=(50, 50))
        time.sleep(0.5)

        return {
            "content": [TextContent(type="text", text=f"Text:'{text}' added at ({text_x},{text_y})")]
        }
    except Exception as e:
        logger.error(f"Error in add_text_in_paint: {str(e)}")
        return {"content": [TextContent(type="text", text=f"Error adding text: {str(e)}")]}

@mcp.tool()
async def open_paint() -> dict:
    """Open Microsoft Paint maximized"""
    global paint_app
    try:
        logger.debug("Starting Paint opening operation")
        paint_app = Application().start('mspaint.exe')
        time.sleep(1)
        
        paint_window = paint_app.window(class_name='MSPaintApp')
        
        # Get initial window position
        initial_rect = paint_window.rectangle()
        logger.debug(f"Initial Paint window rectangle: {initial_rect}")
        
        # Maximize window
        win32gui.ShowWindow(paint_window.handle, win32con.SW_MAXIMIZE)
        time.sleep(0.5)
        
        # Get maximized position
        max_rect = paint_window.rectangle()
        logger.debug(f"Maximized Paint window rectangle: {max_rect}")
        
        # Get canvas
        canvas = paint_window.child_window(class_name='MSPaintView')
        canvas_rect = canvas.rectangle()
        logger.debug(f"Canvas rectangle: {canvas_rect}")
        
        return {
            "content": [TextContent(type="text", text="Paint opened successfully and maximized")]
        }
    except Exception as e:
        logger.error(f"Error in open_paint: {str(e)}")
        return {"content": [TextContent(type="text", text=f"Error opening Paint: {str(e)}")]}


# DEFINE RESOURCES

# Add a dynamic greeting resource
@mcp.resource("greeting://{name}")
def get_greeting(name: str) -> str:
    """Get a personalized greeting"""
    logger.info("CALLED: get_greeting(name: str) -> str:")
    return f"Hello, {name}!"


# DEFINE AVAILABLE PROMPTS
@mcp.prompt()
def review_code(code: str) -> str:
    return f"Please review this code:\n\n{code}"
    logger.info("CALLED: review_code(code: str) -> str:")


@mcp.prompt()
def debug_error(error: str) -> list[base.Message]:
    return [
        base.UserMessage("I'm seeing this error:"),
        base.UserMessage(error),
        base.AssistantMessage("I'll help debug that. What have you tried so far?"),
    ]

system_prompt = f"""
...
Examples:
- FUNCTION_CALL: add|5|3
- FUNCTION_CALL: strings_to_chars_to_int|INDIA
- FUNCTION_CALL: open_paint
- FUNCTION_CALL: draw_rectangle|100|100|500|400  # Large visible rectangle
- FUNCTION_CALL: add_text_in_paint|Result = 42  # Text will be placed at (100,100)
...
"""

if __name__ == "__main__":
    # Check if running with mcp dev command
    logger.info("STARTING")
    if len(sys.argv) > 1 and sys.argv[1] == "dev":
        mcp.run()  # Run without transport for dev server
    else:
        mcp.run(transport="stdio")  # Run with stdio for direct execution

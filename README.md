# Paint MCP Server

This project demonstrates an AI agent that performs calculations and visualizes results in Microsoft Paint using the Model-Code-Process (MCP) framework.

## Overview

The Paint MCP Server project consists of two main components:

1. **MCP Server (`example2.py`)**: Provides calculation tools and Paint automation capabilities
2. **AI Agent (`talk2mcp-2.py`)**: Communicates with the MCP server and uses Gemini LLM to process user queries

This system allows users to perform calculations and automatically visualize the results in Microsoft Paint through natural language queries.

## Features

- **Calculation Tools**: Addition, subtraction, multiplication, division, power, square root, factorial, and more
- **Data Processing**: Convert strings to ASCII values, calculate exponential sums of lists, generate Fibonacci sequences
- **Paint Automation**:
  - Open Microsoft Paint automatically
  - Draw rectangles with specified coordinates
  - Add text to Paint drawings
- **LLM Integration**: Uses Google's Gemini 1.5 Pro for natural language understanding

## Requirements

- Windows operating system (for Paint automation)
- Python 3.8+
- Google Gemini API key
- Required Python packages:
  - `mcp` (Model-Code-Process framework)
  - `google-generativeai`
  - `pywinauto`
  - `win32gui`, `win32con`, `win32com.client`, `win32api`, `win32process`
  - `psutil`
  - `python-dotenv`
  - `PIL` (Pillow)

## Setup

1. Clone this repository

2. Install required packages:
   ```
   pip install mcp google-generativeai pywinauto pywin32 psutil python-dotenv pillow
   ```

3. Create a `.env` file in the project root with your Gemini API key:
   ```
   GEMINI_API_KEY=your_api_key_here
   ```

4. Ensure you're using Windows with Microsoft Paint installed

## Usage

1. Start the AI agent:
   ```
   python talk2mcp-2.py
   ```

2. The agent will process the default query: "Find the ASCII values of characters in INDIA, calculate the sum of exponentials of those values, and visualize the result in Paint."

3. The agent will:
   - Calculate the ASCII values of "INDIA"
   - Calculate the sum of exponentials for these values
   - Open Microsoft Paint
   - Draw a rectangle
   - Add the result text to the Paint canvas

## How It Works

### MCP Server (`example2.py`)

The server defines tools that can be called by the AI agent:

1. **Calculation Tools**:
   - Basic operations: `add`, `subtract`, `multiply`, `divide`, etc.
   - Special functions: `strings_to_chars_to_int`, `int_list_to_exponential_sum`, etc.

2. **Paint Automation Tools**:
   - `open_paint`: Opens Microsoft Paint and maximizes the window
   - `draw_rectangle`: Draws a rectangle from coordinates (x1,y1) to (x2,y2)
   - `add_text_in_paint`: Adds text to the Paint canvas

### AI Agent (`talk2mcp-2.py`)

The agent communicates with the MCP server and manages the workflow:

1. Establishes a connection to the MCP server
2. Retrieves available tools from the server
3. Creates a system prompt with tool descriptions
4. Processes the user query using Google's Gemini LLM
5. Executes function calls based on the LLM's response
6. Handles the visualization workflow:
   - Opening Paint
   - Drawing a rectangle
   - Adding text with the result

## Logs

The project creates log files in a `logs` directory with timestamps to help with debugging:
- `example2.py` creates logs with prefix `debug_`
- `talk2mcp-2.py` currently logs to the console

## Extending the Project

### Adding New Calculation Tools

To add a new calculation tool, add a new function to `example2.py` with the `@mcp.tool()` decorator:

```python
@mcp.tool()
def my_new_function(param1: int, param2: int) -> float:
    """Description of what the function does"""
    logger.info("CALLED: my_new_function(param1: int, param2: int) -> float:")
    return float(some_calculation(param1, param2))
```

### Adding New Paint Tools

To add new Paint automation capabilities, add a new async function with the `@mcp.tool()` decorator:

```python
@mcp.tool()
async def my_new_paint_tool(param1: int, param2: str) -> dict:
    """Description of what the tool does"""
    global paint_app
    try:
        # Implementation here
        return {
            "content": [TextContent(type="text", text="Success message")]
        }
    except Exception as e:
        logger.error(f"Error in my_new_paint_tool: {str(e)}")
        return {"content": [TextContent(type="text", text=f"Error message: {str(e)}")]}
```

## Troubleshooting

- **Paint Automation Issues**: Make sure Paint is installed and accessible. The automation relies on window coordinates which may vary depending on your system resolution and UI settings.
- **LLM Generation Timeouts**: If you experience timeouts with the Gemini API, try increasing the timeout value in the `generate_with_timeout` function.
- **API Key Issues**: Ensure your Gemini API key is valid and properly configured in the `.env` file.

## License

This project is provided as an example implementation of the MCP framework. 
import os
from dotenv import load_dotenv
from mcp import ClientSession, StdioServerParameters, types
from mcp.client.stdio import stdio_client
import asyncio
import google.generativeai as genai
from concurrent.futures import TimeoutError
from functools import partial
import traceback
import time
import win32gui
import win32con
import win32com.client
import win32api
import win32process
import psutil
from pywinauto import Application
from mcp.types import TextContent
import argparse

# Load environment variables from .env file
load_dotenv()

# Configure the Gemini API
api_key = os.getenv("GEMINI_API_KEY")
print(f"API Key loaded: {'Yes' if api_key else 'No'}")  # Will print Yes/No without exposing the key
genai.configure(api_key=api_key)

max_iterations = 3
last_response = None
iteration = 0
iteration_response = []

# Global variable for email flag
send_email_flag = False

async def generate_with_timeout(client, prompt, timeout=30):
    """Generate content with a timeout"""
    print("Starting LLM generation...")
    try:
        # Convert the synchronous generate_content call to run in a thread
        loop = asyncio.get_event_loop()
        
        # Use gemini-pro model without the 'models/' prefix
        model = genai.GenerativeModel('gemini-1.5-pro')
        
        response = await asyncio.wait_for(
            loop.run_in_executor(
                None, 
                lambda: model.generate_content(contents=prompt)
            ),
            timeout=timeout
        )
        print("LLM generation completed")
        return response
    except TimeoutError:
        print("LLM generation timed out!")
        raise
    except Exception as e:
        print(f"Error in LLM generation: {e}")
        # If first attempt fails, try with gemini-1.5-pro
        try:
            print("Trying with gemini-1.5-pro...")
            model = genai.GenerativeModel('gemini-1.5-pro') 
            response = await asyncio.wait_for(
                loop.run_in_executor(
                    None, 
                    lambda: model.generate_content(contents=prompt)
                ),
                timeout=timeout
            )
            print("LLM generation completed with alternate model")
            return response
        except Exception as e2:
            print(f"Error with alternate model: {e2}")
            raise

def reset_state():
    """Reset all global variables to their initial state"""
    global last_response, iteration, iteration_response
    last_response = None
    iteration = 0
    iteration_response = []

async def main(send_email=False):
    global send_email_flag
    send_email_flag = send_email
    
    reset_state()  # Reset at the start of main
    print("Starting main execution...")
    try:
        # Create a single MCP server connection
        print("Establishing connection to MCP server...")
        server_params = StdioServerParameters(
            command="python",
            args=["example2.py"]
        )

        async with stdio_client(server_params) as (read, write):
            print("Connection established, creating session...")
            async with ClientSession(read, write) as session:
                print("Session created, initializing...")
                await session.initialize()
                
                # Get available tools
                print("Requesting tool list...")
                tools_result = await session.list_tools()
                tools = tools_result.tools
                print(f"Successfully retrieved {len(tools)} tools")

                # Create system prompt with available tools
                print("Creating system prompt...")
                print(f"Number of tools: {len(tools)}")
                
                try:
                    # First, let's inspect what a tool object looks like
                    # if tools:
                    #     print(f"First tool properties: {dir(tools[0])}")
                    #     print(f"First tool example: {tools[0]}")
                    
                    tools_description = []
                    for i, tool in enumerate(tools):
                        try:
                            # Get tool properties
                            params = tool.inputSchema
                            desc = getattr(tool, 'description', 'No description available')
                            name = getattr(tool, 'name', f'tool_{i}')
                            
                            # Format the input schema in a more readable way
                            if 'properties' in params:
                                param_details = []
                                for param_name, param_info in params['properties'].items():
                                    param_type = param_info.get('type', 'unknown')
                                    param_details.append(f"{param_name}: {param_type}")
                                params_str = ', '.join(param_details)
                            else:
                                params_str = 'no parameters'

                            tool_desc = f"{i+1}. {name}({params_str}) - {desc}"
                            tools_description.append(tool_desc)
                            print(f"Added description for tool: {tool_desc}")
                        except Exception as e:
                            print(f"Error processing tool {i}: {e}")
                            tools_description.append(f"{i+1}. Error processing tool")
                    
                    tools_description = "\n".join(tools_description)
                    print("Successfully created tools description")
                except Exception as e:
                    print(f"Error creating tools description: {e}")
                    tools_description = "Error loading tools"
                
                print("Created system prompt...")
                
                system_prompt = f"""You are an AI agent that solves problems and visualizes results in Microsoft Paint. You have access to various tools for calculations and visualization.

Available tools:
{tools_description}

You must respond with EXACTLY ONE line in one of these formats (no additional text):
1. For function calls:
   FUNCTION_CALL: function_name|param1|param2|...
   
2. For visualization:
   FUNCTION_CALL: open_paint
   FUNCTION_CALL: draw_rectangle|x1|y1|x2|y2
   FUNCTION_CALL: add_text_in_paint|text

3. For final answers:
   FINAL_ANSWER: [your_answer]

Important:
- When a function returns multiple values, you need to process all of them
- Only give FINAL_ANSWER when you have completed all necessary calculations
- Do not repeat function calls with the same parameters  
- When solving a problem that needs visualization:
  1. First perform all calculations
  2. Then call open_paint (for visualization)
  3. Then draw_rectangle for the frame (for visualization)
  4. Finally add_text_in_paint with the result (for visualization)

Examples:
- FUNCTION_CALL: add|5|3
- FUNCTION_CALL: strings_to_chars_to_int|INDIA
- FUNCTION_CALL: open_paint
- FUNCTION_CALL: draw_rectangle|400|300|1200|600        # Filled black rectangle
- FUNCTION_CALL: add_text_in_paint|Result = 42          # Black text at (500,400)
"""

                query = """Find the ASCII values of characters in INDIA, calculate the sum of exponentials of those values, and visualize the result in Paint."""
                print("Starting iteration loop...")
                
                # Use global iteration variables
                global iteration, last_response
                
                while iteration < max_iterations:
                    print(f"\n--- Iteration {iteration + 1} ---")
                    if last_response is None:
                        current_query = query
                    else:
                        current_query = current_query + "\n\n" + " ".join(iteration_response)
                        current_query = current_query + "  What should I do next?"

                    # Get model's response with timeout
                    print("Preparing to generate LLM response...")
                    prompt = f"{system_prompt}\n\nQuery: {current_query}"
                    try:
                        response = await generate_with_timeout(None, prompt)
                        response_text = response.text.strip()
                        print(f"LLM Response: {response_text}")
                        
                        # Find the FUNCTION_CALL line in the response
                        for line in response_text.split('\n'):
                            line = line.strip()
                            if line.startswith("FUNCTION_CALL:"):
                                response_text = line
                                break
                        
                        if response_text.startswith("FUNCTION_CALL:"):
                            _, function_info = response_text.split(":", 1)
                            parts = [p.strip() for p in function_info.split("|")]
                            func_name, params = parts[0], parts[1:]
                            
                            print(f"\nDEBUG: Raw function info: {function_info}")
                            print(f"DEBUG: Split parts: {parts}")
                            print(f"DEBUG: Function name: {func_name}")
                            print(f"DEBUG: Raw parameters: {params}")
                            
                            try:
                                # Find the matching tool to get its input schema
                                tool = next((t for t in tools if t.name == func_name), None)
                                if not tool:
                                    print(f"DEBUG: Available tools: {[t.name for t in tools]}")
                                    raise ValueError(f"Unknown tool: {func_name}")

                                print(f"DEBUG: Found tool: {tool.name}")
                                print(f"DEBUG: Tool schema: {tool.inputSchema}")

                                # Prepare arguments according to the tool's input schema
                                arguments = {}
                                schema_properties = tool.inputSchema.get('properties', {})
                                print(f"DEBUG: Schema properties: {schema_properties}")

                                for param_name, param_info in schema_properties.items():
                                    if not params:  # Check if we have enough parameters
                                        raise ValueError(f"Not enough parameters provided for {func_name}")
                                        
                                    value = params.pop(0)  # Get and remove the first parameter
                                    param_type = param_info.get('type', 'string')
                                    
                                    print(f"DEBUG: Converting parameter {param_name} with value {value} to type {param_type}")
                                    
                                    # Convert the value to the correct type based on the schema
                                    if param_type == 'integer':
                                        arguments[param_name] = int(value)
                                    elif param_type == 'number':
                                        arguments[param_name] = float(value)
                                    elif param_type == 'array':
                                        # Handle array input - convert all remaining parameters to integers
                                        if func_name == "int_list_to_exponential_sum":
                                            # For int_list_to_exponential_sum, use all parameters including the first one
                                            array_values = [int(value)] + [int(p.strip()) for p in params]
                                            arguments[param_name] = array_values
                                            # Clear the params list since we've used all values
                                            params.clear()
                                        else:
                                            # For other array parameters, handle as before
                                            if isinstance(value, str):
                                                value = value.strip('[]').split(',')
                                            arguments[param_name] = [int(x.strip()) for x in value]
                                    else:
                                        arguments[param_name] = str(value)

                                print(f"DEBUG: Final arguments: {arguments}")
                                print(f"DEBUG: Calling tool {func_name}")
                                
                                result = await session.call_tool(func_name, arguments=arguments)
                                print(f"DEBUG: Raw result: {result}")
                                
                                # Get the full result content
                                if hasattr(result, 'content'):
                                    print(f"DEBUG: Result has content attribute")
                                    # Handle multiple content items
                                    if isinstance(result.content, list):
                                        iteration_result = [
                                            item.text if hasattr(item, 'text') else str(item)
                                            for item in result.content
                                        ]
                                    else:
                                        iteration_result = str(result.content)
                                else:
                                    print(f"DEBUG: Result has no content attribute")
                                    iteration_result = str(result)
                                    
                                print(f"DEBUG: Final iteration result: {iteration_result}")
                                
                                # Format the response based on result type
                                if isinstance(iteration_result, list):
                                    result_str = f"[{', '.join(iteration_result)}]"
                                else:
                                    result_str = str(iteration_result)
                                
                                iteration_response.append(
                                    f"In the {iteration + 1} iteration you called {func_name} with {arguments} parameters, "
                                    f"and the function returned {result_str}."
                                )
                                last_response = iteration_result

                                # If we've completed the calculation, proceed with visualization
                                if func_name == "int_list_to_exponential_sum":
                                    print("\n===  AI Agent Execution (Calculation) Complete, Proceeding with Visualization ===")
 
                                    print("\nStep 1: Opening Microsoft Paint...")
                                    # Open Paint
                                    result = await session.call_tool("open_paint")
                                    print(f"✓ {result.content[0].text}")
                                    await asyncio.sleep(1)

                                    print("\nStep 2: Drawing rectangle frame...")
                                    # Draw rectangle
                                    result = await session.call_tool(
                                        "draw_rectangle",
                                        arguments={
                                            "x1": 400,
                                            "y1": 300,
                                            "x2": 1200,
                                            "y2": 600
                                        }
                                    )
                                    print(f"✓ {result.content[0].text}")

                                    print("\nStep 3: Adding result text...")
                                    # Add text with the result
                                    result = await session.call_tool(
                                        "add_text_in_paint",
                                        arguments={
                                            "text": f"Result = {result_str}"
                                        }
                                    )
                                    print(f"✓ {result.content[0].text}")
                                    print("\n=== Visualization Complete ===")
                                    print("The result has been displayed in Microsoft Paint.")
                                    print("You can find the visualization in the Paint window.")
                                    break

                            except Exception as e:
                                print(f"DEBUG: Error details: {str(e)}")
                                print(f"DEBUG: Error type: {type(e)}")
                                import traceback
                                traceback.print_exc()
                                iteration_response.append(f"Error in iteration {iteration + 1}: {str(e)}")
                                break

                        elif response_text.startswith("FINAL_ANSWER:"):
                            print("\n=== Agent Execution Complete ===")
                            break

                        iteration += 1
                    except Exception as e:
                        print(f"Failed to get LLM response: {e}")
                        break

    except Exception as e:
        print(f"Error in main execution: {e}")
        traceback.print_exc()
    finally:
        reset_state()  # Reset at the end of main

def find_paint_window():
    """Find Paint window using multiple methods"""
    def enum_windows_callback(hwnd, result):
        if win32gui.IsWindowVisible(hwnd):
            window_text = win32gui.GetWindowText(hwnd)
            class_name = win32gui.GetClassName(hwnd)
            # Paint window class is usually 'MSPaintApp' or contains 'Paint'
            if "Paint" in window_text or "MSPaintApp" in class_name:
                result.append(hwnd)
        return True

    paint_windows = []
    win32gui.EnumWindows(enum_windows_callback, paint_windows)
    return paint_windows[0] if paint_windows else None

def force_activate_window(hwnd):
    """Force activate a window using multiple methods"""
    if not hwnd:
        return False
    
    try:
        # Get current foreground window
        current_fg = win32gui.GetForegroundWindow()
        
        # Get the current window's thread
        current_thread = win32api.GetCurrentThreadId()
        
        # Get the target window's thread
        target_thread = win32process.GetWindowThreadProcessId(hwnd)[0]
        
        # Attach the threads
        win32process.AttachThreadInput(target_thread, current_thread, True)
        
        try:
            # Show the window
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            win32gui.ShowWindow(hwnd, win32con.SW_SHOW)
            
            # Force foreground
            win32gui.SetForegroundWindow(hwnd)
            
            # Maximize
            win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
            
            # Additional focus attempts
            win32gui.BringWindowToTop(hwnd)
            win32gui.SetFocus(hwnd)
            
        finally:
            # Detach threads
            win32process.AttachThreadInput(target_thread, current_thread, False)
        
        # Wait for window to be active
        time.sleep(0.5)
        return win32gui.GetForegroundWindow() == hwnd
        
    except Exception as e:
        print(f"Error activating window: {e}")
        return False

def ensure_paint_active():
    """Ensure Paint is running and active"""
    # Find Paint process
    paint_pid = None
    for proc in psutil.process_iter(['pid', 'name']):
        if 'mspaint' in proc.info['name'].lower():
            paint_pid = proc.info['pid']
            break
    
    if not paint_pid:
        print("DEBUG: Paint process not found")
        return False
    
    # Find Paint window
    hwnd = find_paint_window()
    if not hwnd:
        print("DEBUG: Paint window not found")
        return False
    
    # Force activate window
    success = True
    # success = force_activate_window(hwnd)
    # if success:
    #     print("DEBUG: Paint window successfully activated")
    # else:
    #     print("DEBUG: Failed to activate Paint window")
    
    return success

if __name__ == "__main__":
    asyncio.run(main())

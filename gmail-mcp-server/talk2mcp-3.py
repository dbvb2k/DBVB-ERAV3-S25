import os
from dotenv import load_dotenv
from mcp import ClientSession, StdioServerParameters, types
from mcp.client.stdio import stdio_client
from mcp.server.fastmcp import FastMCP
import asyncio
import google.generativeai as genai
from concurrent.futures import TimeoutError
from functools import partial
import traceback
import time
import logging
from datetime import datetime
import win32gui
import win32con
import win32com.client
import win32api
import win32process
import psutil
from pywinauto import Application
from mcp.types import TextContent
import argparse
import sys
import multiprocessing
import subprocess

# Configure logging
log_dir = "logs"
if not os.path.exists(log_dir):
    os.makedirs(log_dir)

timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
log_file = os.path.join(log_dir, f"talk2mcp_{timestamp}.log")

logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_file),
        logging.StreamHandler()
    ]
)

logger = logging.getLogger(__name__)

# Load environment variables from .env file
load_dotenv()

# Configure the Gemini API
api_key = os.getenv("GEMINI_API_KEY")
print(f"API Key loaded: {'Yes' if api_key else 'No'}")  # Will print Yes/No without exposing the key
genai.configure(api_key=api_key)

# Get recipient email from environment variables
recipient_email = os.getenv("GMAIL_RECIPIENT_EMAIL")
print(f"Recipient email loaded: {'Yes' if recipient_email else 'No'}")  # Will print Yes/No without exposing the email

max_iterations = 3
last_response = None
iteration = 0
iteration_response = []

async def generate_with_timeout(client, prompt, timeout=30):
    """Generate content with a timeout"""
    logger.info("Starting LLM generation...")
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
        logger.info("LLM generation completed")
        return response
    except TimeoutError:
        logger.error("LLM generation timed out!")
        raise
    except Exception as e:
        logger.error(f"Error in LLM generation: {e}")
        # If first attempt fails, try with gemini-1.5-pro
        try:
            logger.info("Trying with gemini-1.5-pro...")
            model = genai.GenerativeModel('gemini-1.5-pro') 
            response = await asyncio.wait_for(
                loop.run_in_executor(
                    None, 
                    lambda: model.generate_content(contents=prompt)
                ),
                timeout=timeout
            )
            logger.info("LLM generation completed with alternate model")
            return response
        except Exception as e2:
            logger.error(f"Error with alternate model: {e2}")
            raise

def reset_state():
    """Reset all global variables to their initial state"""
    global last_response, iteration, iteration_response
    last_response = None
    iteration = 0
    iteration_response = []

async def wait_for_server_startup(process, timeout=30):
    """Wait for server to start up with timeout"""
    start_time = time.time()
    while time.time() - start_time < timeout:
        if process.returncode is not None:
            # Get the error output from the process
            stderr = await process.stderr.read()
            error_msg = stderr.decode() if stderr else "No error message available"
            logger.error(f"Server process error output: {error_msg}")
            raise Exception(f"Server process exited with code {process.returncode}. Error: {error_msg}")
        
        # Check if the process is still running
        if process.returncode is None:
            # Try to read any output to see if server is ready
            try:
                stdout = await process.stdout.read(1024)
                if stdout:
                    output = stdout.decode()
                    logger.debug(f"Server output: {output}")
                    # Check for the SERVER_READY signal
                    if "SERVER_READY" in output:
                        logger.info("Server signaled it is ready")
                        return
            except Exception as e:
                logger.debug(f"Error reading stdout: {e}")
        
        await asyncio.sleep(0.1)
    
    # If we get here, the server didn't start in time
    logger.error("Server startup timed out")
    # Try to terminate the process gracefully
    try:
        process.terminate()
        await asyncio.sleep(1)
        if process.returncode is None:
            process.kill()
    except Exception as e:
        logger.error(f"Error terminating process: {e}")
    raise TimeoutError("Server startup timed out")

async def main():
    reset_state()  # Reset at the start of main
    logger.info("Starting main execution...")
    try:
        # Create a single MCP server connection
        logger.info("Establishing connection to MCP server...")
        
        # Verify environment variables
        creds_file_path = os.getenv("GMAIL_CREDS_FILE_PATH")
        token_path = os.getenv("GMAIL_TOKEN_PATH")
        
        if not creds_file_path or not token_path:
            logger.error("Missing required environment variables: GMAIL_CREDS_FILE_PATH or GMAIL_TOKEN_PATH")
            print("Error: Missing required environment variables. Please check your .env file.")
            return
            
        if not os.path.exists(creds_file_path):
            logger.error(f"Credentials file not found: {creds_file_path}")
            print(f"Error: Credentials file not found: {creds_file_path}")
            return
            
        logger.debug(f"Using credentials file: {creds_file_path}")
        logger.debug(f"Using token file: {token_path}")
        
        # Get the current directory and add it to Python path
        current_dir = os.path.dirname(os.path.abspath(__file__))
        if current_dir not in sys.path:
            sys.path.append(current_dir)
        
        # Create server parameters
        server_params = StdioServerParameters(
            command=sys.executable,
            args=["-u", "server.py", 
                  "--creds-file-path", creds_file_path,
                  "--token-path", token_path]
        )

        # Start the server process
        logger.info("Starting server process...")
        server_process = subprocess.Popen(
            [server_params.command] + server_params.args,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            bufsize=1
        )        
        
        # Wait for server to start
        logger.info("Waiting for server to start...")
        try:
            while True:
                output = server_process.stdout.readline()
                if not output:
                    # Check if process ended prematurely
                    if server_process.poll() is not None:
                        error = server_process.stderr.read()
                        logger.error(f"Server failed to start: {error}")
                        raise Exception(f"Server failed to start: {error}")
                    continue
                    
                logger.info(f"Server output: {output.strip()}")
                if "SERVER_READY" in output:
                    logger.info("Server is ready")
                    break
        except Exception as e:
            logger.error(f"Error waiting for server: {e}")
            if server_process.poll() is None:
                server_process.terminate()
            return

        # print(server_process.list_tools())

        try:
            # Create MCP client
            logger.info("Creating stdio client...")
            # Use the server parameters for stdio client
            async with stdio_client(server_params) as (read, write):
                logger.info("stdio_client created successfully")
                
                logger.info("Creating session...")
                async with ClientSession(read, write) as session:
                    logger.info("Session created, initializing...")
                    try:
                        # Add timeout for session initialization
                        await asyncio.wait_for(session.initialize(), timeout=30)
                        logger.info("Session initialized successfully")
                    except asyncio.TimeoutError:
                        logger.error("Session initialization timed out")
                        print("Error: Session initialization timed out. The server might be unresponsive.")
                        return
                    except Exception as e:
                        logger.error(f"Failed to initialize session: {str(e)}")
                        print(f"Error: Failed to initialize session: {str(e)}")
                        return
                
                    # Get available tools
                    logger.info("Requesting tool list...")
                    tools_result = await session.list_tools()
                    tools = tools_result.tools
                    logger.info(f"Successfully retrieved {len(tools)} tools")

                    # Create system prompt with available tools
                    logger.info("Creating system prompt...")
                    logger.info(f"Number of tools: {len(tools)}")
                
                    try:
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
                                logger.info(f"Added description for tool: {tool_desc}")
                            except Exception as e:
                                logger.error(f"Error processing tool {i}: {e}")
                                tools_description.append(f"{i+1}. Error processing tool")
                        
                        tools_description = "\n".join(tools_description)
                        logger.info("Successfully created tools description")
                    except Exception as e:
                        logger.error(f"Error creating tools description: {e}")
                        tools_description = "Error loading tools"
                    
                    logger.info("Created system prompt...")
                    
                    system_prompt = f"""You are an AI agent that solves problems and performs calculations. You have access to various tools for calculations and email sending.

    Available tools:
    {tools_description}

    You must respond with EXACTLY ONE line in one of these formats (no additional text):
    1. For function calls:
    FUNCTION_CALL: function_name|param1|param2|...
    
    2. For final answers:
    FINAL_ANSWER: [your_answer]

    Important:
    - When solving a problem that needs multiple steps:
    1. First perform all calculations
    2. Then send email with the results if requested
    - When a function returns multiple values, process all of them
    - Do not repeat function calls with the same parameters

    Examples:
    - FUNCTION_CALL: strings_to_chars_to_int|INDIA
    - FUNCTION_CALL: int_list_to_exponential_sum|[73, 78, 68, 73, 65]
    - FUNCTION_CALL: send-email|recipient_id|subject|message
    """

                    query = """Find the ASCII values of characters in INDIA, calculate the sum of exponentials of those values, and send the results via email."""
                    logger.info("Starting iteration loop...")
                    
                    # Use global iteration variables
                    global iteration, last_response
                    
                    while iteration < max_iterations:
                        logger.info(f"\n--- Iteration {iteration + 1} ---")
                        if last_response is None:
                            current_query = query
                        else:
                            current_query = current_query + "\n\n" + " ".join(iteration_response)
                            current_query = current_query + "  What should I do next?"

                        # Get model's response with timeout
                        logger.info("Preparing to generate LLM response...")
                        prompt = f"{system_prompt}\n\nQuery: {current_query}"
                        try:
                            response = await generate_with_timeout(None, prompt)
                            response_text = response.text.strip()
                            logger.info(f"LLM Response: {response_text}")
                            
                            # Split response into multiple lines and process each FUNCTION_CALL
                            function_calls = [line.strip() for line in response_text.split('\n') 
                                            if line.strip().startswith("FUNCTION_CALL:")]
                            
                            # Process each function call in sequence
                            last_calculation_result = None  # Store most recent calculation result
                            
                            for function_call in function_calls:
                                logger.info(f"\nProcessing function call: {function_call}")
                                response_text = function_call
                                
                                if response_text.startswith("FUNCTION_CALL:"):
                                    _, function_info = response_text.split(":", 1)
                                    parts = [p.strip() for p in function_info.split("|")]
                                    func_name, params = parts[0], parts[1:]
                                    
                                    logger.debug(f"\nDEBUG: Raw function info: {function_info}")
                                    logger.debug(f"DEBUG: Split parts: {parts}")
                                    logger.debug(f"DEBUG: Function name: {func_name}")
                                    logger.debug(f"DEBUG: Raw parameters: {params}")
                                    
                                    # If this is send-email following a calculation, ensure we use the latest result
                                    if func_name == "send-email" and last_calculation_result:
                                        # Find the message parameter (typically the last one)
                                        message_index = -1
                                        if len(params) >= 3:  # We need at least recipient, subject, and message
                                            message_index = 2  # Message is typically the third parameter
                                            
                                            # Check if the message seems to reference a calculation result
                                            if "sum" in params[message_index].lower() or "exponential" in params[message_index].lower():
                                                # Get current date and time
                                                current_datetime = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                                                
                                                # Update the message to include the correct calculation result with new format
                                                logger.info(f"Updating email message with latest calculation result: {last_calculation_result}")
                                                params[message_index] = f"""Hi User,

The problem statement was to find the 'Sum of exponentials of ASCII values of string [INDIA]'. This has been computed and here is the result:

The sum of exponentials is: {last_calculation_result} 
[Computed Date / Time: {current_datetime}]"""
                                    
                                    try:
                                        # Find the matching tool to get its input schema
                                        tool = next((t for t in tools if t.name == func_name), None)
                                        if not tool:
                                            logger.debug(f"DEBUG: Available tools: {[t.name for t in tools]}")
                                            raise ValueError(f"Unknown tool: {func_name}")

                                        logger.debug(f"DEBUG: Found tool: {tool.name}")
                                        
                                        # Prepare arguments based on tool schema
                                        arguments = {}
                                        schema_properties = tool.inputSchema.get('properties', {})
                                        
                                        for param_name, param_info in schema_properties.items():
                                            if not params:
                                                raise ValueError(f"Not enough parameters provided for {func_name}")
                                            
                                            value = params.pop(0)
                                            param_type = param_info.get('type', 'string')
                                            
                                            # Special handling for recipient_id in send-email
                                            if func_name == "send-email" and param_name == "recipient_id" and value == "recipient_id":
                                                if recipient_email:
                                                    arguments[param_name] = recipient_email
                                                else:
                                                    raise ValueError("No recipient email found in environment variables")
                                            # Normal parameter processing
                                            elif param_type == 'integer':
                                                arguments[param_name] = int(value)
                                            elif param_type == 'number':
                                                arguments[param_name] = float(value)
                                            elif param_type == 'array':
                                                if isinstance(value, str):
                                                    value = value.strip('[]').split(',')
                                                arguments[param_name] = [int(x.strip()) for x in value]
                                            else:
                                                arguments[param_name] = str(value)

                                        logger.debug(f"DEBUG: Final arguments: {arguments}")
                                        
                                        # Make sure email message contains the latest calculation result if needed
                                        if func_name == "send-email" and "message" in arguments and last_calculation_result:
                                            # Check if the message seems to reference a calculation result
                                            if ("sum" in arguments["message"].lower() or 
                                               "exponential" in arguments["message"].lower() or 
                                               "calculation" in arguments["message"].lower()):
                                                # Get current date and time
                                                current_datetime = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                                                
                                                # Format the email with the specified template
                                                arguments["message"] = f"""Hi User,

The problem statement was to find the 'Sum of exponentials of ASCII values of string [INDIA]'. This has been computed and here is the result:

The sum of exponentials is: {last_calculation_result} [Computed Date / Time: {current_datetime}]"""
                                                
                                                logger.info(f"Updated email message with calculation result and formatted template")
                                        
                                        # Call the tool
                                        result = await session.call_tool(func_name, arguments=arguments)
                                    
                                        # Get the full result content
                                        if hasattr(result, 'content'):
                                            logger.debug(f"DEBUG: Result has content attribute")
                                            # Handle multiple content items
                                            if isinstance(result.content, list):
                                                iteration_result = [
                                                    item.text if hasattr(item, 'text') else str(item)
                                                    for item in result.content
                                                ]
                                            else:
                                                iteration_result = str(result.content)
                                        else:
                                            logger.debug(f"DEBUG: Result has no content attribute")
                                            iteration_result = str(result)
                                        
                                        logger.debug(f"DEBUG: Final iteration result: {iteration_result}")
                                        
                                        # Format the response based on result type
                                        if isinstance(iteration_result, list):
                                            result_str = f"[{', '.join(iteration_result)}]"
                                        else:
                                            result_str = str(iteration_result)
                                        
                                        # Store the last result for possible use in next function calls
                                        last_response = iteration_result
                                        
                                        # If this is a calculation function, store the result for potential email use
                                        if func_name == "int_list_to_exponential_sum" or func_name == "strings_to_chars_to_int":
                                            if isinstance(iteration_result, list) and iteration_result:
                                                last_calculation_result = iteration_result[0]
                                            else:
                                                last_calculation_result = iteration_result
                                            logger.info(f"Stored calculation result: {last_calculation_result}")
                                        
                                        iteration_response.append(
                                            f"In the {iteration + 1} iteration you called {func_name} with {arguments} parameters, "
                                            f"and the function returned {result_str}."
                                        )
                                        
                                    except Exception as e:
                                        logger.error(f"DEBUG: Error in function call {func_name}: {e}")
                                        traceback.print_exc()
                                        continue  # Continue with next function call even if one fails
                                    
                            # Break the main loop after processing all function calls
                            break
                            
                        except Exception as e:
                            logger.error(f"Failed to get LLM response: {e}")
                            break

                        if response_text.startswith("FINAL_ANSWER:"):
                            logger.info("\n=== Agent Execution Complete ===")
                            break

                        iteration += 1

        except Exception as e:
            logger.error(f"Error in stdio_client or session creation: {str(e)}")
            logger.error(traceback.format_exc())  # NEW: Full traceback
            print(f"Error: Failed to create connection: {str(e)}")
            return
        finally:
            # Clean up server process
            if server_process.poll() is None:
                server_process.terminate()
                server_process.wait()

    except Exception as e:
        logger.error(f"Error in main execution: {str(e)}")
        logger.error(f"Traceback: {traceback.format_exc()}")
        print(f"Error: {str(e)}")
        print("Check the log file for detailed error information.")
    finally:
        reset_state()  # Reset at the end of main

if __name__ == "__main__":
    print("Starting application...")
    try:
        asyncio.run(main())
    except Exception as e:
        print(f"Fatal error: {str(e)}")
        print("Check the log file for detailed error information.")
    print("Application completed.")
    
    

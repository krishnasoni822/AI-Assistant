import pyttsx3
import sys
import speech_recognition as sr
from textblob import TextBlob
import time
import pandas as pd
import matplotlib.pyplot as plt
import os

# Initialize Text-to-Speech engine
engine = pyttsx3.init()
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)
engine.setProperty('rate', 150)

tasks = []

# Speak Function
def speak(text):
    for sentence in text.split('.'):
        if sentence.strip():
            engine.say(sentence.strip())
            engine.runAndWait()
            time.sleep(0.5)

# Listen Function
def listen():
    recognizer = sr.Recognizer()
    with sr.Microphone() as source:
        print("🎤 Assistant is listening... Speak now, Krishna.")
        recognizer.adjust_for_ambient_noise(source, duration=1)
        audio = recognizer.listen(source)
    try:
        command = recognizer.recognize_google(audio)
        print("You said:", command)
        return command.lower()
    except sr.UnknownValueError:
        speak("Sorry Krishna, I couldn't understand.")
        return ""
    except sr.RequestError:
        speak("Network issue, Krishna.")
        return ""

# Sample sales data (for demo)
sales_data = {
    "Month": ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
              "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"],
    "Product_A_Sales": [12000, 13500, 12800, 15000, 16000, 15500,
                        17000, 16500, 15800, 17200, 18000, 19000],
    "Product_B_Sales": [10000, 9500, 11000, 12000, 11500, 13000,
                        12500, 14000, 13500, 14500, 15000, 15500],
    "Product_C_Sales": [8000, 8500, 9000, 9500, 10000, 10500,
                        11000, 11500, 12000, 13000, 13500, 14000],
}
df_sales = pd.DataFrame(sales_data)


# =============== NEW FUNCTION =====================
def handle_excel_chart_command(command):
    """Handle user request to create a chart from an Excel file"""
    try:
        # Ask for Excel file path
        speak("Please enter the full path of your Excel file, Krishna.")
        file_path = input("Enter Excel file path: ").strip().replace('"', '')

        if not os.path.exists(file_path):
            speak("Sorry Krishna, I cannot find that file.")
            print("❌ File not found.")
            return

        # Load Excel file
        df = pd.read_excel(file_path)
        speak("File loaded successfully.")
        print("\n✅ Excel Data Preview:\n", df.head())

        # Ask for chart type
        speak("Which chart would you like to create? Bar, Pie, or Line?")
        chart_type = input("Enter chart type: ").strip().lower()

        # Ask for columns
        speak("Please enter the column name for X axis.")
        x_col = input(f"Available columns: {list(df.columns)}\nX-axis column: ").strip()

        speak("Please enter the column name for Y axis.")
        y_col = input(f"Available columns: {list(df.columns)}\nY-axis column: ").strip()

        # Create chart
        plt.figure(figsize=(10, 5))

        if chart_type == "bar":
            plt.bar(df[x_col], df[y_col], color="skyblue")
            plt.title(f"{y_col} by {x_col}")
            plt.xlabel(x_col)
            plt.ylabel(y_col)

        elif chart_type == "pie":
            if df[y_col].sum() == 0:
                speak("Pie chart is not suitable for zero values.")
                return
            plt.pie(df[y_col], labels=df[x_col], autopct="%1.1f%%", startangle=90)
            plt.title(f"{y_col} Distribution by {x_col}")

        elif chart_type == "line":
            plt.plot(df[x_col], df[y_col], marker='o', color='green')
            plt.title(f"{y_col} over {x_col}")
            plt.xlabel(x_col)
            plt.ylabel(y_col)
            plt.grid(True)

        else:
            speak("Unknown chart type. Please choose bar, pie, or line.")
            return

        plt.tight_layout()
        plt.show()
        speak(f"Here is your {chart_type} chart, Krishna.")

    except Exception as e:
        speak("Sorry Krishna, there was an error while creating the chart.")
        print(f"❌ Error: {e}")
# ==================================================


# Existing calculation function
def handle_calculation_commands(command):
    if "calculate" in command or "solve" in command:
        command = command.replace("calculate", "").replace("solve", "").strip()
    try:
        command = command.lower().replace("^", "**").strip()
        command = command.replace("plus", "+").replace("minus", "-").replace("times", "*").replace("divided by", "/")
        result = eval(command)
        speak(f"The result of {command} is {result}")
        print(f"The result of {command} is {result}")
    except Exception:
        speak("Sorry Krishna, I couldn't calculate that. Please try again.")
        print("❌ Error: Unable to calculate expression.")


from datetime import datetime

def tell_time():
    now = datetime.now()
    current_time = now.strftime("%H:%M:%S")
    speak(f"The current time is {current_time}")
    print(f"🕒 Current time: {current_time}")




# Chart handler for built-in data
def handle_chart_command(command):
    command = command.lower()

    if "product a" in command:
        product = "Product_A_Sales"
    elif "product b" in command:
        product = "Product_B_Sales"
    elif "product c" in command:
        product = "Product_C_Sales"
    else:
        product = None

    if "pie" in command:
        chart_type = "pie"
    elif "bar" in command:
        chart_type = "bar"
    elif "histogram" in command or "hist" in command:
        chart_type = "histogram"
    elif "line" in command:
        chart_type = "line"
    else:
        chart_type = None

    if chart_type == "pie":
        speak("Creating a pie chart for you, Krishna.")
        values = df_sales[["Product_A_Sales", "Product_B_Sales", "Product_C_Sales"]].sum()
        plt.figure(figsize=(6, 6))
        plt.pie(values, labels=["Product A", "Product B", "Product C"], autopct='%1.1f%%', startangle=90)
        plt.title("Total Sales Share (Pie Chart)")
        plt.show()

    elif chart_type == "bar":
        speak("Creating a bar chart for you, Krishna.")
        plt.figure(figsize=(10, 5))
        plt.plot(df_sales["Month"], df_sales["Product_A_Sales"], label="Product A")
        plt.plot(df_sales["Month"], df_sales["Product_B_Sales"], label="Product B")
        plt.plot(df_sales["Month"], df_sales["Product_C_Sales"], label="Product C")
        plt.title("Monthly Sales Comparison (Bar Chart)")
        plt.xlabel("Month")
        plt.ylabel("Sales (₹)")
        plt.legend()
        plt.tight_layout()
        plt.show()

    elif chart_type == "line":
        speak("Creating a line chart for you, Krishna.")
        plt.figure(figsize=(10, 5))
        plt.plot(df_sales["Month"], df_sales["Product_A_Sales"], marker='o', label="Product A")
        plt.plot(df_sales["Month"], df_sales["Product_B_Sales"], marker='s', label="Product B")
        plt.plot(df_sales["Month"], df_sales["Product_C_Sales"], marker='^', label="Product C")
        plt.title("Monthly Sales Trend (Line Chart)")
        plt.xlabel("Month")
        plt.ylabel("Sales (₹)")
        plt.legend()
        plt.grid(True)
        plt.tight_layout()
        plt.show()

    else:
        speak("Sorry Krishna, I couldn't understand the chart type you requested.")
        print("Try saying: 'create a pie chart', 'make a bar chart', or 'show histogram'.")
# average
def calculate_average(df, column_name):
    if column_name not in df.columns:
        speak(f"Column {column_name} does not exist in the data.")
        print(f"❌ Column {column_name} not found.")
        return
    avg = df[column_name].mean()
    speak(f"The average of {column_name} is {avg:.2f}")
    print(f"📊 Average of {column_name}: {avg:.2f}")
# sales total
def total_sales_for_month(df, month_name):
    month_name = month_name.capitalize()  # ensure format matches
    if "Month" not in df.columns:
        speak("The data does not have a Month column.")
        print("❌ Month column not found.")
        return
    row = df[df["Month"] == month_name]
    if row.empty:
        speak(f"No data found for {month_name}.")
        print(f"❌ No sales data for {month_name}")
        return
    total_sales = row[["Product_A_Sales", "Product_B_Sales", "Product_C_Sales"]].sum(axis=1).values[0]
    speak(f"The total sales for {month_name} are {total_sales}")
    print(f"📊 Total sales for {month_name}: {total_sales}")



# Task Manager
def handle_task_commands(command):
    if "add task" in command or "create project" in command:
        task = command.replace("add task", "").replace("create project", "").strip()
        if task:
            tasks.append(task)
            speak(f"Task added: {task}")
        else:
            speak("Please specify the task after 'add task'.")
    elif "show tasks" in command or "list tasks" in command:
        if tasks:
            speak("Here are your tasks:")
            for i, task in enumerate(tasks, 1):
                speak(f"{i}. {task}")
        else:
            speak("Your task list is empty.")
    elif "remove task" in command:
        try:
            number = int(command.split()[-1])
            if 1 <= number <= len(tasks):
                removed = tasks.pop(number - 1)
                speak(f"Removed task: {removed}")
            else:
                speak("Invalid task number.")
        except:
            speak("Please specify the task number to remove, like 'remove task 2'.")
    else:
        speak("Sorry, I can only manage tasks or chat with you.")


# Chat
def chat_response(command):
    if any(greet in command for greet in ["hello", "namaste", "hi", "hey", "good morning", "good afternoon", "good evening"]):
        speak("Hello Boss! How can I help you today?")
        print("Hello Boss! How can I help you today?")
    elif "your name" in command or "who are you" in command:
        speak("I am your assistant. My name is Rora.")
    elif "my name" in command:
        speak("Yes boss, your name is Krishna.")
    elif "purpose" in command:
        speak("I am your virtual assistant designed to help you with tasks and calculations.")
        print("I am your virtual assistant designed to help you with tasks and calculations.")
    elif "how are you" in command:
        speak("I am fine, Krishna. How about you?")
    elif any(word in command for word in ["thank you", "thankyou", "shukriya"]):
        speak("You're welcome, Krishna.")
    elif "i am fine" in command:
        speak("thats good, Krishna.")
    else:
        speak("I'm sorry Krishna, I am not able to answer that right now.")


# MAIN
def main():
    speak("Hello Krishna, I am ready to talk.")
    while True:
        command = input("You: ").lower()

        if any(exit_word in command for exit_word in ["exit", "quit", "bye", "goodbye", "get lost"]):
            speak("Goodbye Krishna. Have a great day!")
            break

        elif "excel" in command or "file" in command:
            handle_excel_chart_command(command)
        
        elif "time" in command or "current time" in command:
             tell_time()
        
        elif "average" in command:
             speak("Please enter the column name to calculate average:")
             col = input(f"Available columns: {list(df_sales.columns)}\nColumn: ").strip()
             calculate_average(df_sales, col)

        elif "total sales" in command or "sales for" in command:
             speak("Which month do you want the total sales for?")
             month = input("Enter month name (e.g., January): ").strip()
             total_sales_for_month(df_sales, month)

         
        elif "chart" in command or "graph" in command or "plot" in command:
            handle_chart_command(command)

        elif "task" in command:
            handle_task_commands(command)

        elif "calculate" in command or "solve" in command:
            handle_calculation_commands(command)

        else:
            chat_response(command)


# RUN
if __name__ == "__main__":
    main()


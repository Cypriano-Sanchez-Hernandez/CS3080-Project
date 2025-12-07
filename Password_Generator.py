import secrets # Used for cryptographically secure random selection
import tkinter as tk # GUI Framework
import pandas as pd # For creating and writing Excel files
import os  # File system interactions
import tkinter.simpledialog as simpledialog # Popup dialog for input
import platform # Determine operating system for file opening
from datetime import datetime # Timestamp for saved passwords

#----------------------------------------
#          Password Generator Class
#----------------------------------------
class Password_Generator:
    """
    This class:
    - Stores user preference for password creating
    - Builds the character pool based on these preferences
    - Generates secure random passwords
    """

    def __init__(self):
        # Prefs holds user-selected options like length, types of characters, ...
        self.prefs ={}
        # The pool will contain all allowed characters based on the users preferences
        self.pool = ""

    def set_preferences(self, length, lowercase, uppercase, numbers, symbols):
        """
        This saves the user's selected preferences into the prefs dictionary.
        These options determine how the password generation acts.
        """

        # Stores each chosen preference in a dictionary so it can be used later
        self.prefs["length_input"] = length
        self.prefs["symbols"] = symbols
        self.prefs["uppercase"] = uppercase
        self.prefs["lowercase"] = lowercase
        self.prefs["numbers"] = numbers
        return self.prefs
    
    def create_pool(self):
        """
        This builds a string of characters the generator can use.
        Only characters from user-selected options are included.
        """

        # This resets the pool each time a new password is requested.
        self.pool= ""

        # Each conditional block appends a category of characters to the pool,
        # which ensures only selected character types are included.

        # Adds lowercase letters if it is chosen.
        if self.prefs["lowercase"]:
            self.pool += 'abcdefghijklmnopqrstuvwxyz'
        
        # Adds uppercase letters if it is chosen.
        if self.prefs["uppercase"]: 
            self.pool += 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'

        # Adds digits if it is chosen.
        if self.prefs["numbers"]: 
            self.pool += '0123456789'

        # Adds symbols if it is chosen.
        if self.prefs["symbols"]: 
            self.pool += '!@#$%^&*()_+-=[]}{|;:,.<>/?' 

    def generate_password(self):
        """
        This creates the password by randomly selecting characters from the created pool using secrets.choice() 
        which is cryptographically secure compared to random.choice().
        """
        
        password = ""

        # Loop repeats once per character in the password based on the preferred length

        # This picks a random character from the pool for each position in the password
        for i in range(self.prefs["length_input"]):
            password += secrets.choice(self.pool)
        return password

#---------------------------------
#           GUI Class
#---------------------------------
class PasswordGUI:
    """
    This class creates and manages the Tkinter GUI.
    This acts as the front-end interface for the Password_Generator class.
    """

    def __init__(self, master):
        # This is the window setup
        self.master = master
        self.master.title("Password Generator")
        
        # This creates an instance of the generator class
        self.generator = Password_Generator()

        # This is the main title in the GUI
        tk.Label(master, text="Password Generator", font=("Times New Roman", 45, "underline"), fg="black", bg="#2887BE").pack(padx=10, pady=(10, 0))

        # This is a prompt for password length
        tk.Label(master, text="Type In Value For Password Length:", font=("Times New Roman", 30, "italic"), fg="black", bg="#2887BE").pack(padx=10, pady=(10, 0))

        # Input box for length
        self.length_entry = tk.Entry(master, font=("Times New Roman", 35), width=3, justify='center')
        self.length_entry.pack(padx=10, pady=10)

        # Checkboxes for character options
        self.lower_var = tk.BooleanVar(value=False)
        self.upper_var = tk.BooleanVar(value=False)
        self.num_var = tk.BooleanVar(value=False)
        self.sym_var = tk.BooleanVar(value=False)

        # Checkboxes displayed in the GUI
        tk.Checkbutton(master, font=("Times New Roman", 30, "italic"), fg="black", bg="#2887BE", text="Include Lowercase?", variable=self.lower_var).pack(padx=10, pady=2)
        tk.Checkbutton(master, font=("Times New Roman", 30, "italic"), fg="black", bg="#2887BE", text="Include Uppercase?", variable=self.upper_var).pack(padx=10, pady=2)
        tk.Checkbutton(master, font=("Times New Roman", 30, "italic"), fg="black", bg="#2887BE", text="Include Numbers?", variable=self.num_var).pack(padx=10, pady=2)
        tk.Checkbutton(master, font=("Times New Roman", 30, "italic"), fg="black", bg="#2887BE", text="Include Symbols?", variable=self.sym_var).pack(padx=10, pady=2)
        
        # Button to trigger password creation
        tk.Button(master, text="Generate Password", font=("Times New Roman", 27, "bold"), fg="black", command=self.generate_password, bg="#058420").pack(padx=10, pady=10)

        # This is a warning label
        tk.Label(master, text="*Do Not Have Excel File Open While Generating and Saving New Password!", font=("Times New Roman", 27), fg="#E3A849", bg="#2887BE").pack(padx=10, pady=(10, 10))

        # This button opens the Excel file 
        tk.Button(master, text="Open Excel File", font=("Times New Roman", 27, "bold"), fg="black", command=self.open_excel, bg="#058420").pack(padx=10, pady=10)

        # This is a label for displaying generated passwords or errors.
        self.result_label = tk.Label(master, text="", font=("Times New Roman", 20, "bold"), fg="white", bg="#811081")
        self.result_label.pack(padx=10, pady=(10, 10))

    #----------------------------------------------------
    #       Function to open existing Excel file
    #----------------------------------------------------
    def open_excel(self):
        # Name of the excel file
        file_path = "passwords_sheet.xlsx"

        # This ensures the file exists
        if not os.path.exists(file_path):
            self.result_label.config(text="Excel file does not exist yet.")
            return
        
        # This means the opening method depends on OS
        system = platform.system()

        try:
            # os.startfile only works on Windows
            if system == "Windows":
                os.startfile(file_path) # Windows-only function
        except Exception as e:
            # Displays any error while trying to open the file
            self.result_label.config(text=f"Error opening file: {e}")

    #----------------------------------------------
    #    Main function that generates password
    #----------------------------------------------
    def generate_password(self):
        """
        This handles:
        - Validating user input
        - Collecting preferences
        - Creating password
        - Saving it to Excel
        """

        # This validates length input
        try:
            length = int(self.length_entry.get())

            # This ensures length is not negative or zero
            if length <= 0:
                self.result_label.config(text="Length has to be positive.", font=("Arial", 20))
                return
        except ValueError:
            # User typed something that is not a number
            self.result_label.config(text="Length has to be a number.")
            return
        
        # This reads checkbox values
        lowercase = self.lower_var.get()
        uppercase = self.upper_var.get()
        numbers = self.num_var.get()
        symbols = self. sym_var.get()

        # This stores preferences
        self.generator.set_preferences(length, lowercase, uppercase, numbers, symbols)

        # This builds the character pool
        self.generator.create_pool()

        # If the pool is empty, the user selected no character types
        if not self.generator.pool:
            self.result_label.config(text="Select at least one character type.")
            return
        
        # This generates the password
        password = self.generator.generate_password()
        self.result_label.config(text=password)

        # This ask the user what the password is for
        purpose = simpledialog.askstring("Purpose", "What is this password for?")

        # Builds timestamp for records
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

        # This saves only if the user provided a purpose
        if purpose:
            self.save_excel(purpose, password, timestamp)
    
    #---------------------------------------
    #  Save Password Record To Excel File
    #---------------------------------------
    def save_excel(self, purpose, password, timestamp):
        """
        This creates or appends password data into an Excel spreadsheet.
        The spreadsheet has columns: Purpose, Password, and Timestamp.
        """

        # Converts data row into a DataFrame
        df = pd.DataFrame([{
            "Purpose": purpose.upper(), # Converts purpose to uppercase for consistency
            "Password": password,
            "Timestamp": timestamp
        }])

        file_path = "passwords_sheet.xlsx"

        # This creates the excel file if it does not exist
        if not os.path.exists(file_path):
            # This creates file and sheet for first-time setup
            with pd.ExcelWriter(file_path, engine="xlsxwriter") as writer:
                df.to_excel(writer, index=False, sheet_name="Passwords")

                # Adjusts the column widths for better readability
                worksheet = writer.sheets["Passwords"]
                worksheet.set_column(0, 0, 25)
                worksheet.set_column(1, 1, 35)
                worksheet.set_column(2, 2, 25)

        # This appends passwords to the existing file
        else:
            # Reads previous stored passwords
            existing = pd.read_excel(file_path)

            # Combines old and new entries
            updated = pd.concat([existing, df], ignore_index=True)

            # Writes updated data back to the file
            with pd.ExcelWriter(file_path, engine="xlsxwriter") as writer:
                updated.to_excel(writer, index=False, sheet_name="Passwords")

                worksheet = writer.sheets["Passwords"]
                worksheet.set_column(0, 0, 25)
                worksheet.set_column(1, 1, 35)
                worksheet.set_column(2, 2, 25)

#------------------------
#     Main Program
#------------------------
if __name__ == "__main__":
    # Creates the main Tkinter window
    root = tk.Tk()

    # Initializes and attaches GUI to the window
    app = PasswordGUI(root)

    # Sets background color for entire window
    root.configure(bg="#2887BE")

    # Base dimensions of the GUI window
    root.geometry("1200x800")

    # Start Tkinter main loop (keeps window open)
    root.mainloop()
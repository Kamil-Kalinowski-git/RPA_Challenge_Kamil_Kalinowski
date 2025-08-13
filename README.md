# Data Entry Automation - RPA Challenge

This project is a script that automates the process of data entry into the web application at `rpachallenge.com`.

---

## Objective

The goal of this task is to create a Robotic Process Automation (RPA) script that will:
1.  Open a web browser and navigate to `rpachallenge.com`.
2.  Download the Excel data file provided on the website.
3.  Fill in the form on the page, dynamically matching the data to the fields, as the layout changes in each round.
4.  After completing all rounds, retrieve the final score.
5.  Save the result to a text file.
6.  Close the browser.

---

## How to Run the Project

To run this script, follow these steps:
1.  Open the project in your preferred code editor (e.g., Visual Studio Code).
2.  Install dependencies by opening a terminal within the editor and running the following command:
    ```bash
    pip install -r requirements.txt
    ```
3.  Run the script:
    ```bash
    python main.py
    ```
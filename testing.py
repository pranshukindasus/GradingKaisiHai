#!/usr/bin/env python3
"""
Unified Professor Grade Lookup

Fetches the courses taught by a professor from your intranet site and merges them
with grade distributions in grades.xlsx.

Requirements:
  • Python 3.8+
  • pip install pandas selenium webdriver-manager openpyxl

Usage:
  python unified_professor_grades.py
"""
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
import time

GRADES_FILE = "grades.xlsx"
INSTR_PARAM = "instructname"
BASE_URL = "http://172.26.142.68/dccourse/"
ALLOWED_GRADES = ["A*", "A", "B+", "B", "C+", "C", "D+", "D", "F", "S", "X"]

import sys, os

if getattr(sys, 'frozen', False):
    # Running in PyInstaller bundle
    BASE_DIR = sys._MEIPASS
else:
    # Running in normal Python environment
    BASE_DIR = os.path.dirname(__file__)

GRADES_FILE = os.path.join(BASE_DIR, "grades.xlsx")


def wait_for_table_stabilize(driver, max_wait=60, stable_duration=6):
    """
    Dynamically waits for the table to "stop updating" within the `stu_course` frame.
    Specifically, every second we parse the table and check its row count.
    If the row count (and columns) remain unchanged for `stable_duration` seconds,
    we conclude it has stabilized.

    :param driver: the WebDriver in the `stu_course` frame.
    :param max_wait: the maximum total time in seconds to keep checking.
    :param stable_duration: how many seconds of no change needed to conclude it is stable.
    """
    start_time = time.time()
    last_df = None
    stable_secs = 0

    while True:
        # get current table HTML, parse it
        current_html = driver.page_source
        try:
            df_current = pd.read_html(current_html, header=None)[0]
        except ValueError:
            # if for some reason there's no table
            df_current = pd.DataFrame()
        # rename columns so we can do a shape check safely
        if not df_current.empty:
            df_current.columns = df_current.iloc[0]
            df_current = df_current.drop(index=0).reset_index(drop=True)

        if last_df is not None:
            # check if shape changed
            if df_current.shape == last_df.shape:
                stable_secs += 1
            else:
                stable_secs = 0
        last_df = df_current

        # if stable for stable_duration seconds, done
        if stable_secs >= stable_duration:
            break

        # if we exceed max_wait, break anyway
        if (time.time() - start_time) > max_wait:
            break

        time.sleep(1)


def get_professor_courses(prof_name: str) -> pd.DataFrame:
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--disable-gpu")  # often recommended on Windows

    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=chrome_options
    )
    try:
        driver.get(BASE_URL)
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, INSTR_PARAM))
        )
        driver.find_element(By.NAME, INSTR_PARAM).send_keys(prof_name)
        driver.find_element(By.NAME, "showlist").click()

        # Switch to results frame
        WebDriverWait(driver, 10).until(
            EC.frame_to_be_available_and_switch_to_it((By.NAME, "stu_course"))
        )

        # Wait until the table is found
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.TAG_NAME, "table"))
        )

        # Dynamically wait for the table to stop updating
        wait_for_table_stabilize(driver, max_wait=60, stable_duration=6)

        print(f"\nPlease wait...")
        # Finally parse the entire table
        html = driver.page_source
        raw = pd.read_html(html, header=None)[0]
        raw.columns = raw.iloc[0]
        df = raw.drop(index=0).reset_index(drop=True)
        df = df.rename(columns={
            'ACADEMIC YEAR': 'Year',
            'SEM': 'Semester',
            'COURSE NAME': 'Course'
        })[['Year', 'Semester', 'Course']]
        return df
    finally:
        driver.quit()


def main():
    while True:
        prof_name = input("\nEnter the Professor's name, for example \"Rik Dey\": ").strip()
        if not prof_name:
            print("No Professor name entered. Exiting.")
            break

        print(f"Fetching courses for '{prof_name}'...")
        courses_df = get_professor_courses(prof_name)

        if courses_df.empty:
            print("No courses found for this professor.")
            continue  # go back to the input prompt
        grades_df = pd.read_excel(GRADES_FILE)
        #grades_df = grades_df[grades_df['Grade'].isin(ALLOWED_GRADES)]
        # Convert Count column to integer type
        grades_df['Count'] = grades_df['Count'].astype(int)
        # Standardize types for merge
        for df in (courses_df, grades_df):
            df['Year'] = df['Year'].astype(str).str.strip()
            df['Course'] = df['Course'].astype(str).str.strip()
        courses_df['Semester'] = courses_df['Semester'].astype(int)
        grades_df['Semester'] = grades_df['Semester'].astype(int)

        merged = pd.merge(
            courses_df, grades_df,
            on=['Year', 'Semester', 'Course'],
            how='left'
        )

        pivot = merged.pivot_table(
            index=['Year', 'Semester', 'Course'],
            columns='Grade',
            values='Count',
            aggfunc='sum',
            fill_value=0
        ).reset_index()

        print(f"\nGrade distribution for courses taught by {prof_name}:")
        print(pivot.to_string(index=False))

if __name__ == '__main__':
    print('\nWelcome to Grading Kaisi Hai?! \n')
    print(
        f'Please Note:\n1. This program runs only on Windows with Chrome installed. \n\n2. You NEED to be connected to IITK\'s campus network to use this. Use Forticlient VPN if you\'re off-campus. \n\n3. The Program can display grades\' distributions till 2024-25 Odd Semester. \n\n4. Please try different permutations of the Professor\'s name whom you wish to look for.\nFor Example, Professor "Manoj K Harbola" goes by the name M K Harbola in the dataset. You may also type "Harbola" and get the same results. \nHowever "Manoj K Harbola" will not work. You may visit "http://172.26.142.68/dccourse/" to look up valid names in case of confusion. \n\n5. The Program displays grades of all Professors who satisfy the name you enter. \nFor Example, "Abhishek" will display grades for "Abhishek Gupta" as well as "Abhishek Sarkar". \nHence, please be careful with the name you enter. ')

    main()

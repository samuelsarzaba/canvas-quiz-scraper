from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import Workbook
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.alert import Alert
import time
from typing import Dict, List

quiz_link = ""


def login(driver, username, password):
    wait = WebDriverWait(driver, 30)

    login_to_canvas_button = wait.until(
        EC.element_to_be_clickable((By.XPATH, '//a[@title="Login to Canvas"]'))
    )
    login_to_canvas_button.click()

    username_field = wait.until(EC.presence_of_element_located((By.ID, "username")))
    username_field.send_keys(username)

    password_field = driver.find_element(By.ID, "password")
    password_field.send_keys(password)

    login_button = driver.find_element(
        By.XPATH, '//button[@type="submit" and text()="Login"]'
    )
    login_button.click()

    button = wait.until(
        EC.element_to_be_clickable((By.ID, "dont-trust-browser-button"))
    )
    button.click()


def get_attempts(wait: WebDriverWait) -> List[str]:
    attempt_container = wait.until(
        EC.presence_of_element_located((By.CLASS_NAME, "ic-Table"))
    )

    attempts = attempt_container.find_elements(By.TAG_NAME, "tr")[1:]

    attempt_links = set()
    for attempt in attempts:
        try:
            link_element = attempt.find_element(By.TAG_NAME, "a")
            attempt_link = link_element.get_attribute("href")
            attempt_links.add(attempt_link)
        except:
            continue

    return list(attempt_links)


def scrape_quiz(wait: WebDriverWait, data_dict: Dict[str, List[str]]):
    time.sleep(1)

    question_container = wait.until(
        EC.presence_of_element_located((By.ID, "questions"))
    )

    questions = question_container.find_elements(By.CLASS_NAME, "display_question")

    data_dict.update(
        {
            question.find_element(By.CLASS_NAME, "question_text").text: [
                answer.text.replace("Correct Answer\n", "").strip()
                for answer in question.find_elements(By.CLASS_NAME, "correct_answer")
            ]
            for question in questions
        }
    )


def save_dict_to_excel(data_dict, filename="output.xlsx"):
    wb = Workbook()
    ws = wb.active
    ws.title = "Quiz Results"

    ws.append(["Question", "Correct Answers"])

    for question, answers in data_dict.items():
        ws.append([question, ", ".join(answers)])

    for column in ws.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2) * 1.2
        ws.column_dimensions[column_letter].width = adjusted_width

    wb.save(filename)
    print(f"Data saved to {filename}")


def main():
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
    wait = WebDriverWait(driver, 10)
    data_dict = {}

    try:
        driver.get(quiz_link)

        login(driver, "", "")

        attempts = get_attempts(wait)

        for attempt in attempts:
            driver.get(attempt)
            scrape_quiz(wait, data_dict)

        save_dict_to_excel(data_dict)

    finally:
        driver.quit()


if __name__ == "__main__":
    main()

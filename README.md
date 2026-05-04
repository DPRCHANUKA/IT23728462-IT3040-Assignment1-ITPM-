# IT3040 Assignment 1 - Transliteration Accuracy Testing

## Student
Registration Number: IT23728462 (CHANUKA D P R)

## Description
This project automates negative test cases for the Pixel Suite Chat Sinhala Transliteration function.

## Requirements
- Python 3.11 
- Playwright
- openpyxl

## Installation
pip install -U pip
pip install playwright openpyxl
python -m playwright install

## pixelssuite Run Command
python .\test_automation.py --excel "$PWD\IT23728462_Assignment 1 - Test cases.xlsx" --url "https://www.pixelssuite.com/chat-translator" --wait-ms 5000 --type-delay-ms 80 --slow-mo-ms 200 --save-every 1 --keep-open

## tmrtools run command 
python .\tmrtools_automation.py --excel "$PWD\IT23728462_tmrt_tool.xlsx" --url "https://tmrtools.com/" --wait-ms 6000 --retries 30 --retry-wait-ms 2000 --type-delay-ms 30 --slow-mo-ms 0 --save-every 1


## Files
- test_automation.py
- IT23728462_Assignment 1 - Test cases.xlsx
- IT23728462_GitHub_Repository_Link.txt
- README.md


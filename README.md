# Vocabulary (Foreign Words Learning Method)

- This README describes not only the program but the whole method, which I have developed for myself and use almost every day.
- Have you ever heard of the flashcard method of learning words? I hope my interpretation will help save some time and trees.

## How It Works (Briefly)

- First, work with materials in your target language (books, articles, videos). Write down every word whose meaning or pronunciation you don't know in a special workbook. This will become your personal study dictionary.

- If you don't have such a dictionary yet, don't worry: the program can generate one for you based on the template.xlsx file.

- You'll need a spreadsheet application compatible with your OS (Microsoft Office, LibreOffice Calc, etc.), but this shouldn't be problematic.

- Once you've collected enough words, you can start using this program.

- It helps you memorize new vocabulary deeply with minimal effort and no tedious drilling.

- Just 15 minutes of daily practice can expand your vocabulary by several thousand words, transforming foreign language use from a challenge into routine.

<!-- Let's look at this process in detail. -->

## Installation

### Requirements:

- Python >= 3.10
- [Poetry](https://python-poetry.org/)
- Tkinter (pre-installed on Windows/macOS; requires manual installation on Linux)

**Linux Installation:**

Arch Linux:
```bash
sudo pacman -S tk
```

Debian/Ubuntu:
```bash
sudo apt install python3-tk
```

### Installation Steps:

1. Open your terminal in the desired parent directory

2. Clone the repository:
    ```bash
    git clone https://github.com/benabub/vocabulary.git
    ```

3. Enter the project directory:
    ```bash
    cd vocabulary/vocabulary
    ```

4. Set up the environment:
    ```bash
    poetry install
    ```

## Usage

1. Navigate to the project:
    ```bash
    cd your_parent_dir/vocabulary/vocabulary
    ```

2. Launch the application:
    ```bash
    poetry run python main.py
    ```

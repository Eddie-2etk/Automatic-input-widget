# Automatic Input Widget

A small desktop tool for **automatically typing text into Microsoft Word character by character**.  
Instead of pasting content directly, it simulates real keyboard input so the text appears as if it is being typed naturally.

## Features

- Type into a **new Word document**
- Type into the **current cursor position**
- Adjustable typing speed
- Optional **human-like typo simulation**
- Multiple typing styles:
  - Natural
  - Thoughtful
  - Clumsy
  - Hybrid
- Adjustable typo rate and thinking pauses
- Estimated typing time display
- Load text from clipboard
- Restart typing from the beginning
- Global hotkeys:
  - `Space`: Pause / Resume
  - `Esc`: Stop

## Tech Stack

- Python
- Tkinter
- AppleScript (`osascript`)
- `pynput`
- Microsoft Word
- macOS System Events

## Requirements

- macOS
- Python 3.10+
- Microsoft Word
- Required Python package:
  - `pynput`

Install dependencies with:

```bash
pip install pynput
```

## How to Run

```bash
python app.py
```

## How It Works

The program opens a GUI where you can paste or type your text, then choose one of the following modes:

### 1. Type into a New Word Document
- Automatically opens Microsoft Word
- Creates a new document
- Types the text into it character by character

### 2. Start Typing After 3 Seconds
- Gives you 3 seconds to switch to Word or another target window
- Lets you place the cursor where you want
- Starts typing automatically at the current cursor position

## Permissions

Because the app uses `System Events` to simulate keyboard input, macOS may require you to grant permissions such as:

- Accessibility
- Automation

If the typing does not work properly, check your system permission settings.

## Notes

- This tool simulates keyboard input instead of pasting text directly.
- It is designed for situations where the typing should look natural.
- Avoid switching windows during typing, or the text may be entered in the wrong place.
- This project is currently intended for **macOS + Microsoft Word**.

## File Structure

```bash
.
├── app.py
└── README.md
```

## Future Improvements

- Add more typing styles
- Save user preferences
- Support more target applications
- Provide a packaged executable version
- Improve logging and error handling

## License

For personal use and learning purposes.

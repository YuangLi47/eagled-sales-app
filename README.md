# Eagled Sales App

A light-weight Python desktop application for generating quotations and packing lists, built with Tkinter and pandas.

## Features

- GUI
- Automatic calculations
- Export to PDF
- Mac-ready packaging
- Limited Data Management Capabilities

## Future Improvements

This project was originally created as a self-learning exercise to explore **Tkinter** for the first time and to reduce manual errors in generating quotations and packing lists. Its primary goal is to improve document processing efficiency.

Please note:  
There may be known and undetected bugs within the current codebase due to its early-stage nature.

### Potential Future Updates:
- Refactor the codebase for better structure and maintainability  
- Replace Tkinter with a modern GUI library (e.g., PyQt or Electron) to enhance the user interface  
- Add user authentication for secure access  
- Integrate cloud-based data management  
- Improve user interaction features (e.g., notifications, input validation, form flow)

## Mock Data Notice

All data used in this project is mock data and does not represent real business information. This is done to protect the confidentiality of Eagled Company’s sensitive data.

## Setup

```bash

python -m venv .venv 
source .venv/bin/activate  # or .\.venv\Scripts\activate on Windows
pip install -r requirements.txt

# Package the app
python pyinstaller.py

# Run the app
python main.py
```

## License

This project is licensed under the Creative Commons Attribution-NonCommercial 4.0 International License (CC BY-NC 4.0).

You are free to:
- Use, copy, and modify the code for personal or educational purposes.
- Share or adapt it with proper attribution.

You may not:
- Use this project or any part of it for commercial purposes.

For full license details, see the [LICENSE](./LICENSE) file or visit  
[https://creativecommons.org/licenses/by-nc/4.0/](https://creativecommons.org/licenses/by-nc/4.0/)
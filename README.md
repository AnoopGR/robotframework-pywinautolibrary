# PywinautoLibrary for Robot Framework

This is a custom Robot Framework library for automating Windows GUI applications using the **Pywinauto** library. 
It enables users to control and interact with desktop applications on Windows through Robot Framework test cases.

## What is Pywinauto?

**Pywinauto** is a Python package that allows you to automate and interact with GUI applications running on Windows. 
It supports actions like mouse clicks, keyboard input, and window manipulation, making it an excellent choice for testing Windows desktop applications. 
Pywinauto works by simulating user interactions with the application's controls, making it ideal for automating repetitive tasks and GUI testing.
Pywinauto supports both native Windows GUI controls (via the `win32` backend) and modern UI elements (via the `uia` backend).

## Installation

1. **Install Robot Framework**:
   If you don't already have Robot Framework installed, you can install it using pip:
   ```bash
   pip install robotframework

2. **Install Pywinauto:** Pywinauto is the core library for automating Windows applications:
   ```bash
   pip install pywinauto

3. **Install PywinautoLibrary:** After cloning this repository, you can install the PywinautoLibrary with the following commands:
   ```bash
   git clone https://github.com/AnoopGR/robotframework-pywinautolibrary.git
   cd PywinautoLibrary
   pip install .

## Usage

Once the library is installed, you can use it in your Robot Framework test cases to automate Windows applications. 
Here’s an example of how to use the library:

```robot
*** Settings ***
Library           PywinautoLibrary
    
*** Test Cases ***
Open Notepad, Type Text and Save
     Launch Application    notepad.exe
     Get Dialog From Regex   .*Untitled
     Type Text            Edit    Hello, World!
     Menu Select   File->Save As
     Get Dialog    Save As
     Type Text    edit1    test.txt
     Click    Save
     Close Application
```

## Keyword Documentation

See [Keyword Documentation](https://anoopgr.github.io/robotframework-pywinautolibrary/PywinautoLibrary.html) for available keywords.

## Contributing

Contributions are welcome! To contribute to the project, please follow these steps:

1. Fork the repository.
2. Create a new feature branch (git checkout -b feature-name).
3. Make your changes and commit them (git commit -m 'Add new feature').
4. Push your changes (git push origin feature-name).
5. Open a pull request.
6. Please make sure your code adheres to the existing coding style and includes tests for any new functionality.

## License
This project is licensed under the [MIT License](https://opensource.org/licenses/MIT).

## Reference

Thank you to the authors of **[robotframework-winbot](https://code.google.com/archive/p/robotframework-winbot/)** for their initial work on providing GUI automation capabilities to Robot Framework.

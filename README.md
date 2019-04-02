# roster

This repository serves as a programming project to assess your understanding of certain Python concepts and give you an opportunity to demonstrate your Python proficiency by crafting a solution class and sharing your design decisions.

## the problem

Mrs. Jones, a teacher at Billings Elementary School, has been asked by the administration to track her students' grades in a roster stored in an Excel file. Unfortunatly, Mrs. Jones had a bad experience with Excel as a child and as solemnly sworn never to use Excel again. Your task is to create a Python class for reading and manipulating a provided Excel file so Mrs. Jones doesn't haven't to use Excel to successfully record her students' grades.

### grading

You will be graded on:
- the architecture and elegance of your code
- the cleanliness and documentation of your code
- the number of unit tests that you pass

It is not necessarily expected that you will be able to construct a solution that can pass all the unittests, so focus more on quality if you feel you will have a hard time completing the project.

## python setup

For your convenience, scripts for setting up a Python virtual environment are provided. You must have python installed to use these scripts.

### MacOS and Unix

To setup a Python 2.7 environment on MacOS or linux use the ``setup.sh`` script.

```bash
hostname:username$ sh setup.sh
```

For a Python 3 environment, use the ``setup3.sh`` script. Note that the ``openpyxl`` dependency requires Python >= 3.5.

### Windows

To setup a Python 2.7 environment on Windows use the ``setup.bat`` script.

```batch
C:\Users\username\Documents\roster> setup.bat
```

For a Python 3 environment, use the ``setup3.bat`` script. Note that the ``openpyxl`` dependency requires Python >= 3.5.


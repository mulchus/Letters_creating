In english | [По-русски](../README.md)

# Script for preparing letters of commendation files


### How to use?
Python should already be installed.
It is recommended to use virtualenv/venv to isolate the project.
(https://docs.python.org/3/library/venv.html)
Then use pip to install dependencies.
Open the command line with the Win+R keys and enter:
```commandline
pip install -r requirements.txt
```


### Setting environment variables
There are no environment variables.

In the root folder of the script there should be an Excel file named `awards.xlsx `,
which lists the awardees.
The format of all cells is common. Column headers:
```
num organization surname_name_patronymic post ii_years award_type
```
A template file with a sample filling is attached.

The file may be located in a different path.
The path is set in the command line of the script startup. If the script is run without a database path,
the file is searched by default ``awards.xlsx`` in the root directory with the script.
If it is not there, the path is taken from the settings in the file setup.txt , example:
```
PATH_TO_WINE_FILE='e:\Python\Awards\awards.xlsx '
```


## Launch

- Download the code
- Run the command
```commandline
python main.py [the path to the database.xlsx]
```
- The script will display a list of created award files and the results in the Awards folder, arranged in folders by the name of the organization


## Project goal
Created to solve work tasks
# Task for Data Engineer Job Application at the BFI
---
Sergio Menna 30/08/2024

## Task Overview

The objective of this task is to explore title-level data from the top 15 films in the Weekend Box Office report for 26/07/2024 - 28/07/2024. This involves retrieving additional film details from the OMDb API and analysing box office performance data from July and August 2024.

## How It Works

### Steps:

1. **Extraction**: Extract all data from the box office Excel worksheets.
2. **Selection**: Identify the top 15 film titles for the weekend of 23rd-25th August 2024.
3. **API Retrieval**: Retrieve detailed film information from the OMDb API for each of the 15 films.
4. **Sheet Generation**: Create individual sheets in a new Excel file for each film, containing the data retrieved from OMDb.
5. **Data Integration**: Append the relevant weekend box office data to each corresponding sheet.

## Project Structure

The project was built using Python 3 and Visual Studio Code. It includes the following files:

- `main.py`: The main script containing all the code needed to run the task.
- `README.md`: The file you are currently reading.
- `requirements.txt`: A file listing all the Python libraries required to run the task.
- `.env`: A file containing the API_KEY for accessing the OMDb API.
- `.gitignore`: A list of files and directories to be excluded from version control on GitHub.
- `data/`: A folder containing the BFI box office reports.
- `Task_output.xlsx`: The output Excel file generated by the script.

## Value Proposition

This project provides a consolidated view of film details alongside their box office performance throughout their release. Please note that this is a prototype, and Excel was chosen for simplicity in output format.

## Future Development

To further develop this project, I would:

- Transition the output to a more user-friendly web format, allowing easier navigation through different weekends and providing dedicated pages summarising data by Country of Origin, Distributor, and Year.
- Integrate additional data sources using the IMDb ID obtained from OMDb.
- Implement performance comparisons through graphs and develop forecast models.

## How to Reproduce This Project Locally

To reproduce this demo:

1. Download all files from this GitHub repository.
2. Obtain a personal API_KEY from the OMDb API by signing up at [this link](https://www.omdbapi.com/apikey.aspx).
3. Create a new file called `.env` in the project directory.
4. In the `.env` file, store the API_KEY in the following format:

    ```
    API_KEY=your_api_key
    ```

   Replace `your_api_key` with the actual key obtained from OMDb.

## Fun Fact

The Python code was primarily written while listening to the OST of the film *Challengers*.


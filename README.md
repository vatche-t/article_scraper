# Academic Article Scraper

This Python project provides a tool to scrape academic articles from Google Scholar and NCBI, offering an easy-to-use GUI for specifying search queries, choosing the number of pages to scrape, and saving the gathered data to an Excel file. The tool also visualizes the distribution of articles over time, among other insights.

## Features

- Web scraping from Google Scholar and NCBI.
- User-friendly GUI built with tkinter.
- Logging of scraping process.
- Data visualization with matplotlib.
- Export of scraped data and visuals to Excel.

## Prerequisites

Before running this project, ensure you have Python 3.11 installed on your machine. You'll also need to install the following Python packages:

- `tkinter` (usually comes with Python)
- `matplotlib`
- `pandas`
- `requests`
- `beautifulsoup4`
- `loguru`

You can install the necessary packages using pip:

```bash
pip install matplotlib pandas requests beautifulsoup4 loguru
```

## Usage

To use this tool, run the Python script, and the GUI will appear. Follow these steps:

1. Enter the title or keyword of the article you're interested in.
2. Specify the number of pages you want to scrape.
3. (Optional) Choose an output folder for the Excel file.
4. Select the source from Google Scholar, NCBI, or both.
5. Click on "Extract Data" to start the scraping process.

The status of the scraping will be logged in the GUI, and you will be notified once the data has been successfully saved to an Excel file.

## Project Structure

- `scrape_ncbi_articles` and `scrape_scholar_articles`: Functions to scrape articles from NCBI and Google Scholar, respectively.
- `TkinterHandler`: Custom logging handler to redirect log output to the tkinter interface.
- `save_to_excel`: Saves the scraped articles to an Excel file, along with generated plots.
- `generate_and_save_plots`: Generates and saves plots to the Excel file to visualize the data.
- GUI setup section: Configures the GUI elements and binds actions.

## Contributing

Contributions to improve this project are welcome. Please follow the standard fork-and-pull request workflow.
\

Last updated on: 2024-04-16

Last updated on: 2024-04-18

Last updated on: 2024-04-19

Last updated on: 2024-04-22

Last updated on: 2024-04-27

Last updated on: 2024-04-28

Last updated on: 2024-04-28

Last updated on: 2024-05-04

Last updated on: 2024-05-05

Last updated on: 2024-05-09

Last updated on: 2024-05-11

Last updated on: 2024-05-11

Last updated on: 2024-05-12

Last updated on: 2024-05-13

Last updated on: 2024-12-06
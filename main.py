import re
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, scrolledtext
import matplotlib.pyplot as plt
import io
from loguru import logger
import pandas as pd
import requests
from bs4 import BeautifulSoup

class TkinterHandler:
    def __init__(self, widget):
        self.widget = widget

    def write(self, message):
        self.widget.configure(state='normal')
        self.widget.insert(tk.END, message)
        self.widget.configure(state='disabled')
        # Auto-scroll to the end
        self.widget.see(tk.END)

    def flush(self):
        pass

def scrape_ncbi_articles(query, num_pages):
    logger.info(f"Scraping NCBI articles for query: '{query}' across {num_pages} pages")
    articles = []
    
    base_url = "https://pubmed.ncbi.nlm.nih.gov/"
    query_string = query.replace(' ', '+')

    for page in range(num_pages):
        url = f"{base_url}?term={query_string}&page={page + 1}"
        logger.debug(f"Fetching URL: {url}")
        
        response = requests.get(url)
        if response.status_code == 200:
            logger.debug("Successfully fetched the page")
            soup = BeautifulSoup(response.text, 'html.parser')
            results = soup.find_all('article', class_='full-docsum')
            
            for result in results:
                title_tag = result.find('a', class_='docsum-title')
                title = title_tag.text.strip() if title_tag else 'No Title Found'
                link = base_url + title_tag['href'] if title_tag else 'No Link Found'

                author_tag = result.find('span', class_='docsum-authors full-authors')
                author = author_tag.text.strip() if author_tag else 'No Author Found'
                
                journal_tag = result.find('span', class_='docsum-journal-citation full-journal-citation')
                journal = journal_tag.text.strip() if journal_tag else 'No Journal Found'
                
                date_raw_tag = result.find('span', class_='docsum-journal-citation short-journal-citation')
                date_raw = date_raw_tag.text.strip() if date_raw_tag else 'No Date Found'
                date = re.findall(r'\d{4}', date_raw)[0] if date_raw else 'No Date Found'
                
                articles.append({"Title": title, "Author": author, "Journal": journal, "Date": date, "Link": link})
                logger.debug(f"Article scraped: {title} | {date}")
        else:
            logger.error(f"Failed to fetch the page, status code: {response.status_code}")
            break

    logger.info(f"Finished scraping {len(articles)} articles.")
    return articles

def scrape_scholar_articles(query, num_pages):
    logger.info(f"Starting to scrape articles for query: '{query}' across {num_pages} pages")
    articles = []
    page = 0
    while page < num_pages:
        url = f"https://scholar.google.com/scholar?start={page*10}&q={query}&hl=en&as_sdt=0,5"
        logger.debug(f"Fetching URL: {url}")
        response = requests.get(url)
        if response.status_code == 200:
            logger.debug("Successfully fetched the page")
        else:
            logger.error(f"Failed to fetch the page, status code: {response.status_code}")
            break
        soup = BeautifulSoup(response.text, "html.parser")
        results = soup.find_all("div", class_="gs_ri")
        for result in results:
            title = result.find("h3", class_="gs_rt").text
            authors_and_date = result.find("div", class_="gs_a").text
            link = result.find("a")["href"]

            # Attempt to extract the year using regex. This pattern looks for four consecutive digits.
            date_match = re.search(r'\b\d{4}\b', authors_and_date)
            date = date_match.group(0) if date_match else 'No Date Found'

            articles.append({"Title": title, "Authors": authors_and_date, "Date": date, "Link": link})
            logger.debug(f"Article scraped: {title} | Date: {date}")
        page += 1
        logger.info(f"Page {page} scraped successfully.")
    logger.info("Finished scraping articles")
    return articles

def save_to_excel(articles, filename):
    logger.info(f"Saving articles to Excel file: {filename}")
    df = pd.DataFrame(articles)

    # Using ExcelWriter to allow for adding charts or images
    with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Articles')

        # Call the function to generate and save plots
        generate_and_save_plots(articles, writer)

    logger.info("Data saved to Excel successfully.")


def generate_and_save_plots(articles, excel_writer):
    df = pd.DataFrame(articles)
    
    # Prepare the DataFrame
    if 'Date' in df.columns and not df['Date'].isnull().all():
        df['Year'] = pd.to_datetime(df['Date'], errors='coerce').dt.year
        df = df.dropna(subset=['Year'])
        df['Year'] = df['Year'].astype(int)
    else:
        logger.warning("No valid 'Date' information found for plotting.")
        return

    # Plot 1: Frequency of Articles by Year
    plt.figure(figsize=(10, 6))
    df['Year'].value_counts().sort_index().plot(kind='bar')
    plt.title('Frequency of Articles by Year')
    plt.xlabel('Year')
    plt.ylabel('Number of Articles')
    save_plot_to_excel(excel_writer, 'Yearly Frequency')

    # Plot 2: Frequency of Articles by Journal
    if 'Journal' in df.columns and not df['Journal'].isnull().all():
        plt.figure(figsize=(10, 6))
        df['Journal'].value_counts().head(10).plot(kind='bar')  # Top 10 journals
        plt.title('Top 10 Journals by Number of Articles')
        plt.xlabel('Journal')
        plt.ylabel('Number of Articles')
        save_plot_to_excel(excel_writer, 'Top Journals')

    # Plot 3: Frequency of Articles by Author
    if 'Authors' in df.columns and not df['Authors'].isnull().all():
        all_authors = df['Authors'].str.split(',').explode()  # This line changes to use explode()
        all_authors = all_authors.str.strip()  # Clean up whitespace around author names
        author_counts = all_authors.value_counts().head(10)  # Get top 10 authors
        
        plt.figure(figsize=(10, 6))
        author_counts.plot(kind='bar')
        plt.title('Top 10 Authors by Number of Articles')
        plt.xlabel('Author')
        plt.ylabel('Number of Articles')
        save_plot_to_excel(excel_writer, 'Top Authors')

    # Plot 4: Distribution of Articles Over Time (Line Plot)
    plt.figure(figsize=(10, 6))
    df.groupby('Year').size().plot(kind='line')
    plt.title('Distribution of Articles Over Time')
    plt.xlabel('Year')
    plt.ylabel('Number of Articles')
    save_plot_to_excel(excel_writer, 'Article Distribution Over Time')
    
    df['Title Length'] = df['Title'].apply(len)
    
    plt.figure(figsize=(10, 6))
    plt.scatter(df['Year'], df['Title Length'], alpha=0.5)
    plt.title('Article Complexity (Title Length) Over Time')
    plt.xlabel('Publication Year')
    plt.ylabel('Title Length (as a proxy for complexity)')
    save_plot_to_excel(excel_writer, 'Article Complexity Over Time')

def save_plot_to_excel(excel_writer, sheet_name):
    buffer = io.BytesIO()
    plt.savefig(buffer, format='png')
    buffer.seek(0)

    worksheet = excel_writer.book.add_worksheet(sheet_name)
    worksheet.insert_image('A1', f'{sheet_name}.png', {'image_data': buffer})
    plt.close()

# Function to open folder browser
def browse_folder():
    logger.debug("Opening folder browser")
    folder_path = filedialog.askdirectory()
    entry_folder.delete(0, tk.END)
    entry_folder.insert(tk.END, folder_path)
    logger.info(f"Folder selected: {folder_path}")

def scrape_articles():
    query = entry_query.get()
    num_pages = int(entry_pages.get())
    source = combo_source.get()

    articles = []
    filename_suffix = "scholar_articles"  # Default suffix

    if source == "Google Scholar":
        articles = scrape_scholar_articles(query, num_pages)
        filename_suffix = "google_scholar_articles"
    elif source == "NCBI":
        articles = scrape_ncbi_articles(query, num_pages)
        filename_suffix = "ncbi_articles"
    elif source == "Both":
        articles += scrape_scholar_articles(query, num_pages)
        articles += scrape_ncbi_articles(query, num_pages)
        filename_suffix = "combined_articles"

    folder_path = entry_folder.get()
    filename = f"{folder_path}/{filename_suffix}.xlsx" if folder_path else f"{filename_suffix}.xlsx"

    save_to_excel(articles, filename)
    label_status.config(text=f"Extraction complete. Data saved to {filename}.")
    logger.info("Article scraping and saving process completed.")

# GUI setup
window = tk.Tk()
window.title("Article Scraper")
window.geometry("500x400")

# Redirect logging to the Text widget
log_area = scrolledtext.ScrolledText(window, state='disabled', height=10)
log_area.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)
logger.remove()  # Remove default logger
logger.add(TkinterHandler(log_area), format="{time} {level} {message}", level="DEBUG")

# Article Title or Keyword
label_query = tk.Label(window, text="Article Title or Keyword:")
label_query.pack()
entry_query = tk.Entry(window, width=40)
entry_query.pack()

# Number of Pages
label_pages = tk.Label(window, text="Number of Pages:")
label_pages.pack()
entry_pages = tk.Entry(window, width=40)
entry_pages.pack()

# Output Folder (optional)
label_folder = tk.Label(window, text="Output Folder (optional):")
label_folder.pack()
entry_folder = tk.Entry(window, width=40)
entry_folder.pack()

# Browse button
button_browse = tk.Button(window, text="Browse", command=browse_folder)
button_browse.pack()

# Source selection
label_source = tk.Label(window, text="Select Source:")
label_source.pack()
combo_source = ttk.Combobox(window, values=["Google Scholar", "NCBI", "Both"], state="readonly")
combo_source.pack()
combo_source.set("Google Scholar")

# Extract button
button_extract = tk.Button(window, text="Extract Data", command=scrape_articles)
button_extract.pack()

# Status label
label_status = tk.Label(window, text="")
label_status.pack()

window.mainloop()

# Package Import
from selenium import webdriver # "pip install selenium" or https://pypi.org/project/selenium/
# driver = webdriver.Chrome() # Initialize WebDriver on Chrome
# driver = webdriver.Edge() # Initialize WebDriver on Edge
# driver = webdriver.Firefox() # Initialize WebDriver on Firefox
# driver = webdriver.Ie() # Initialize WebDriver on Ie
# driver = webdriver.Safari() # Initialize WebDriver on Safari
from selenium.webdriver.common.by import By # ...

from openpyxl import Workbook # "pip install openpyxl" or https://pypi.org/project/openpyxl/
wb = Workbook() # Create a new workbook
ws = wb.active # Get default worksheet



# Custom Methods
alphabet = "ZABCDEFGHIJKLMNOPQRSTUVWXY" # For Excel cell column indexing
def cell_write(x, y, t): # Method for writing to a specific cell (18278 * 1048576)
	out = "" # Initialize the column index
	if x > 702: # Case for three letters
		if (x - 26) % 676 == 0: # Case for Zs
			out += alphabet[((x - 26) // 676 - 1) % 26] # Set third letter
		else: # Case for other three letters
			out += alphabet[((x - 26) // 676) % 26] # Set third letter
	if x > 26: # Case for two letters
		if x % 26 == 0: # Case for Zs
			out += alphabet[(x // 26 - 1) % 26] # Set second letter
		else: # Case for other two letters
			out += alphabet[(x // 26) % 26] # Set second letter
	c = ws[out + alphabet[x % 26] + str(y)] # Access a specific cell
	c.value = t # Log corresponding text

dvr = [] # Initialize the driver list
dvrs = [] # Initialize the drivers list
def downloader(url): # Method for downloading files on a specific page
	dvr.append(webdriver.Chrome()) # Initialize WebDriver on Chrome
	dvr[len(dvr) - 1].get(url) # Navigate to page
	dvrs.append([])# Initialize the drivers list
	files = dvr[len(dvr) - 1].find_elements(By.CSS_SELECTOR, "td[data-before=\"Title\"]") # Find all files
	file_links = [file.find_element(By.TAG_NAME, "a").get_attribute("href") for file in files] # Create file links list
	for l in file_links: # Iterate over all file links
		dvrs[len(dvrs) - 1].append(webdriver.Chrome(options=opt)) # Create a new driver
		dvrs[len(dvrs) - 1][len(dvrs[len(dvrs) - 1]) - 1].get(l) # Download the file

opt = webdriver.ChromeOptions() # Initialize Chrome options
opt.add_experimental_option("prefs", { # Add customized options
"plugins.always_open_pdf_externally": True, # Prevent .pdf from opening in a new tab
# "download.default_directory" : "...", # Change download location
}) # ...



# File Logging
downloader("https://bakerhughesrigcount.gcs-web.com/na-rig-count") # Download files on North America Rig Count page

downloader("https://bakerhughesrigcount.gcs-web.com/intl-rig-count") # Download files on International Rig Count page



# Table Logging
driver = webdriver.Chrome() # Initialize WebDriver on Chrome
driver.get("https://bakerhughesrigcount.gcs-web.com/") # Navigate to home page
file_name = driver.find_element(By.TAG_NAME, "h1").text # Find the Excel file name

heads = driver.find_elements(By.TAG_NAME, "th") # Find all heads
for head in heads: # Iterate over all heads

	col = heads.index(head) + 1 # Set the col index
	row = 1 # Set the row index
	cell_write(col, row, head.text) # Log the head
	row += 1 # Increase row index
	
	datas = driver.find_elements(By.CSS_SELECTOR, "td[data-before=\"" + head.text + "\"]") # Find all datas
	for data in datas: # Iterate over all datas
		
		cell_write(col, row, data.text) # Log the data
		row += 1 # Increase row index

wb.save(file_name + ".xlsx") # Save the Excel

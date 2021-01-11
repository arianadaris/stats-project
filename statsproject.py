import requests
from requests_html import HTMLSession
import xlsxwriter # for creating new workbooks and writing

url = "https://pris.iaea.org/PRIS/CountryStatistics/ReactorDetails.aspx?current=609" # link to GINNA data

totalYears = []
totalHours = []

# Statistic calculations
lowest = 0
highest = 0
mean = 0
deviation = 0

# Excel workbook
workbook = xlsxwriter.Workbook('statsproject.xlsx')
worksheet = workbook.add_worksheet()

def extract_data():
    """ 
    Method that extracts year and annual time online (h) for R.E. Ginna Nuclear Power Plant.
    """
    try:
        session = HTMLSession()
        response = session.get(url)

        # Focus on Operating History table rows
        historyTable = (response.html.find('tbody'))[0]
        tableRows = historyTable.find('tr')
        tableRows.remove(tableRows[0]) # remove 1969 row

        # Extract year
        for row in tableRows:
            year = (int) (row.find('td:nth-child(1)')[0].text)
            totalYears.append(year)

        # Extract hours each year
        for row in tableRows:
            hour = (int) (row.find('td:nth-child(4)')[0].text)
            totalHours.append(hour) 
    except requests.exceptions.RequestException as e:
        print(e)


def sort_data():
    """ 
    Method that uses selection sort to sort data
    """
    for i in range(len(totalHours)):
        min = i
        for j in range(i+1, len(totalHours)):
            if totalHours[min] > totalHours[j]:
                min = j

        totalHours[i], totalHours[min] = totalHours[min], totalHours[i]

    lowest = totalHours[0]
    highest = totalHours[len(totalHours)-1]


def calculate_sample():
    """ 
    Method that calculates sample mean and standard deviation for annual time online per year.
    """
    global mean, deviation
    
    # Sample mean calculation
    total = 0
    for hour in totalHours:
        total += hour
    mean = round(total/len(totalHours), 5) # round to 5 decimal
    
    # Sample standard deviation calculation
    sum = 0 # summation of (x-μ)^2
    for hour in totalHours:
        sum += (hour - mean)**2
    deviation = round((sum/len(totalHours))**(1/2), 5) # round to 5 decimal


def create_graph():
    """ 
    Method that creates a workbook and formats the data into a bar graph.
    """
    # Label columns, set column width
    worksheet.write('A1', 'Years')
    worksheet.write('B1', 'Annual Time Online (h)')
    worksheet.set_column(2, 1, 20)

    # Display sample mean and standard deviation
    worksheet.write('C1', 'Sample Mean = {}'.format(mean))
    worksheet.write('C2', 'Sample Standard Deviation = {}'.format(deviation))
    worksheet.set_column(3, 1, 40)

    # Add extracted data to sheet
    worksheet.write_column('A2', totalYears)
    worksheet.write_column('B2', totalHours)

    # Create a chart
    chart = workbook.add_chart({'type': 'column'})

    # Set chart title, legend, column titles and size
    chart.set_title({
        'name': 'Reactor Time Online per Year',
        'name_font': {'size': 32, 'bold': True}
        })
    chart.set_legend({'none': True})
    chart.set_x_axis({
        'name': 'Frequency',
        'name_font': {'size': 24, 'bold': False},
        'num_font': {'size': 16, 'italics': True},
        })
    chart.set_y_axis({
        'name': 'Hours Online (h/yr)',
        'name_font': {'size': 24, 'bold': False},
        'num_font': {'size': 16},
        'min': 4500,
        'max': 11000,
        })

    # Configure chart by adding data
    chart.add_series({
        'categories': '=Sheet1!$A$2:$A$51', # frequency
        'values': '=Sheet1!$B$2:$B$51', # hours
        'border': {'color': '#33658A'},
        'fill': {'color': '#86BBD8'},
        'gap': 25,
    })

    # Insert chart into sheet, close workbook
    worksheet.insert_chart('D1', chart, {'x_scale': 2, 'y_scale': 2})
    workbook.close()
    print("statsproject.xlsx created.")



if __name__ == "__main__":
    extract_data()
    sort_data()
    calculate_sample()
    create_graph()
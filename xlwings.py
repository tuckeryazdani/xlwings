from contextlib import nullcontext
import xlwings as xw
import random

# What this program does:
# - Creates 5 random cash amounts with 1 being a negative number as required by XIRR
# - Creates 5 random dates for which every date is unique and greater than the previous as required by XIRR
# - Calculates examples of formula use:
# - XIRR,XNPV,Index-Match,SUMIF(Positive),COUNTIF(Positive),Max,Min

# Ideas to mention 
# Python Anywhere to run the code constantly from the cloud or windows scheduler to run locally
# Pandas for DataFrame and SQL

if __name__ = '__main__':
    #Open prexisting workbook. Set first worksheet equal to "ws1".
    wb=xw.Book("book1.xlsx")
    ws1=wb.sheets[0]
    ws1.clear()

    cellRange=str(input("Enter Integer")) # Can change this to input() for demo purposes
    cellRangeInt=int(cellRange)+1
    ws1.cells(1,'A').value='Amount'
    ws1.cells(1,'B').value='Dates'

    ws1.range('B2:B'+str(cellRangeInt)).number_format='dd-mm-yyyy'

    # Random Cash amount creation
    for i in range(2,cellRangeInt+1):
        # Make the first iteration give a negative number as is a condition for XIRR
        if i==2:
            ws1.cells(i,'A').value=round(random.uniform(-1000,-2000),0)
        else:
            ws1.cells(i,'A').value=round(random.uniform(0,1000),0)

    # Random Date Creation
    for i in range(2,cellRangeInt+1):
        # Set value for day for first iteration
        if i==2:
            day = str(int(random.uniform(0,5)))
        elif i>2:
            # Set day to be at a minmum 1 day later than the previous day
            day=str(int(random.uniform(int(day)+1,int(day)+3)))
        ws1.cells(i,'B').value='=DATE(2022,02,'+day+')'

    ws1.cells(1,'C').value='Rate'
    ws1.cells(2,'C').value=''+str(round(random.uniform(0,1),2))

    ws1.cells(1,'D').value='XIRR'
    ws1.cells(2,'D').value='=XIRR(A2:A'+cellRange+',B2:B'+cellRange+')'

    ws1.cells(3,'D').value='XNPV'
    ws1.cells(4,'D').value='=XNPV(C2,A2:A'+cellRange+',B2:B'+cellRange+')'

    ws1.cells(1,'E').value='IndexMatch (Get date of MIN)'
    ws1.cells(2,'E').number_format='dd-mm-yyyy'
    ws1.cells(2,'E').value='=INDEX(A:B,MATCH(A2,A:A),2)'

    ws1.cells(3,'E').value='SUMIF (Sum of Positives)'
    ws1.cells(4,'E').value='=SUMIF(A:A,">0")'

    ws1.cells(5,'E').value='COUNTIF (Count Positives)'
    ws1.cells(6,'E').value='=COUNTIF(A:A,">0")'

    ws1.cells(7,'E').value='Max Amount'
    ws1.cells(8,'E').value='=MAX(A:A)'

    ws1.cells(9,'E').value='Min Amount'
    ws1.cells(10,'E').value='=MIN(A:A)'

    if input("Display Offset").lower()=="y":
        ws1.cells(1,'F').value='Offset 2x2 -1 Row/Col'
        ws1.cells(1,'G').value='Formulas Refered by Offset'
        ws1.cells(2,'F').value='=OFFSET(A1:B6,1,3)'

    ws1.cells(9,'E').value='Indirect (A2)'
    ws1.cells(10,'E').value='=INDIRECT("A2")'

    rng = ws1.range('A2:E'+str(cellRangeInt))
    values = rng.value
    formulas = rng.formula
    print(values)
    print("----------------------------------")
    print(formulas)

    if input("Update values?").lower()=="y":
        ws1.cells(1,'G').value="Old "+ws1.cells(1,'A').value+"(-100)"
        for i in range(len(values)):
            # Put old values into "G" subtract 100 from each value and put new values in "A"
            ws1.cells(2+i,'G').value=values[i][0]
            ws1.cells(2+i,'A').value=int(ws1.cells(2+i,'A').value-100)

    # For fun charts
    #chart=ws1.charts.add()
    #chart.set_source_data(ws1.range('A:B').expand())
    #chart.chart_type='line_markers'
    #chart.name='Line Graph'

    # Add save statement You can add "example.xlsx" in parameters to rename
    #wb.save("DemoComplete.xlsx")
    #ws1.clear()
    #wb.close()

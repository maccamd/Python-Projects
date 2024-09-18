from matplotlib.transforms import Transform
import openpyxl as xl
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.widgets import TextBox
from matplotlib.backends.backend_pdf import PdfPages
from PIL import Image

def vdc1_distrubtion():
    df = pd.read_excel(r"P:\CORR Project\Data Trackers\RDEC 1\VDC1 Q3.xlsx.xlsm", sheet_name='Summary Sheet')
    #change formatting to be able to interogate
    df.columns = df.columns.str.strip().str.lower().str.replace(' ', '_')
    df.columns = df.columns.str.replace(' ', '_').str.replace('(', '').str.replace(')', '').str.replace('/', '_').str.replace('+', '').str.replace('#', '').str.replace(',', '')
    #remove columns that arnt needed
    df = df.drop(columns=['comments', 'barometerkpa', 'driver', 'test_name', 'date', 'mileage_miles', 'violations', 'distance_km', 'inertial_work_rating__%','temperature_°c'])
    #print(df)
    #filter criteria 1
    wltc_test = df[(df.cycle == 'WLTC Class3b TYPE 1') | (df.cycle == 'WLTC Class3b TYPE 1 Development')]
    #filter criteria 2
    valid_test = wltc_test[((wltc_test.valid_test_can_be_used_for_conformity_y_n == 'Y') & (wltc_test.test_result_valid_y_n == 'Y'))]
    #print (valid_test)

    #Figure attributes to set page up 
    logo = Image.open("P:\CORR Project\Data Analysis\Extras\MPT Logo.png")
    fig = plt.figure(figsize=[8.25,11.75])
    plt.subplots_adjust(left=0.062, bottom=0.06, right=0.98, top=0.952, wspace=0.321, hspace=0.433)
    newax = fig.add_axes([0.77, 0.78, 0.2, 0.2], anchor='NE',)
    newax.imshow(logo)
    newax.axis('off')
    
    #histograms for each compound, w is the bin width when generating the plot
    #subplot(column,row,graph No)
    # 1  2  3
    # 4  5  6
    # 7  8  9
    # 10 11 12

    #w = 2 np.arange(min(valid_test.co_km_mg), max(valid_test.co_km_mg) + w, w)
    plt.subplot(4,3,1)
    plt.hist(valid_test.co_km_mg, edgecolor='black', bins=10, range=[158,240])

    plt.title('CO Distribution')
    plt.xlabel('CO mg/km')
    plt.ylabel('Frequency')

    #w = 1 np.arange(min(valid_test.co2_km_g), max(valid_test.co2_km_g) + w, w)
    plt.subplot(4,3,2)
    plt.hist(valid_test.co2_km_g, edgecolor='black', bins=10, range=[180,193])
    plt.title('CO2 Distribution')
    plt.xlabel('CO2 g/km')
    plt.ylabel('Frequency')

    #description text
    plt.text(0.77, 0.90,'Data from ' + str(len(valid_test)) + ' valid tests', transform=plt.gcf().transFigure)
    plt.text(0.77, 0.88,'on BT70 BHV in RDEC 1', transform=plt.gcf().transFigure)

    plt.subplot(4,3,4)
    plt.hist(valid_test.thc_km_mg, edgecolor='black', bins=10, range=[10,16])
    plt.title('THC Distribution')
    plt.xlabel('THC mg/km')
    plt.ylabel('Frequency')

    plt.subplot(4,3,5)
    plt.hist(valid_test.nox_km_mg, edgecolor='black', bins=10, range=[9,17])
    plt.title('NOx Distribution')
    plt.xlabel('NOx mg/km')
    plt.ylabel('Frequency')

    plt.subplot(4,3,6)
    plt.hist(valid_test.thc__nox_km_mg, edgecolor='black', bins=10, range=[20,31])
    plt.title('THC + NOx Distribution')
    plt.xlabel('THC + NOx mg/km')
    plt.ylabel('Frequency')

    #w = 10000000000 np.arange(min(valid_test.pn_km_), max(valid_test.pn_km_) + w, w)
    plt.subplot(4,3,7)
    plt.hist(valid_test.pn_km_, edgecolor='black', bins=10)
    plt.title('PN Distribution')
    plt.ticklabel_format(style='sci', axis='x', scilimits=(10,11))
    plt.xlabel('PN #/km')
    plt.ylabel('Frequency')

    #w = 0.05 np.arange(min(valid_test.pm_km_mg), max(valid_test.pm_km_mg) + w, w)
    plt.subplot(4,3,8)
    plt.hist(valid_test.pm_km_mg, edgecolor='black', bins=10, range=[0,0.5])
    plt.title('PM Distribution')
    plt.xlabel('PM mg/km')
    plt.ylabel('Frequency')

    plt.subplot(4,3,10)
    plt.hist(valid_test.n20_mg_km, edgecolor='black', bins=10, range=[0.05,1.9])
    plt.title('N2O Distribution')
    plt.xlabel('N2O mg/km')
    plt.ylabel('Frequency')

    plt.subplot(4,3,11)
    plt.hist(valid_test.l_100km, edgecolor='black', bins=10, range=[7.7,8.6])
    plt.title('Fuel Consumption Distribution')
    plt.xlabel('Fuel conusmption l100/km')
    plt.ylabel('Frequency')

    #w = 0.1 np.arange(min(valid_test.ch4_mg_km), max(valid_test.ch4_mg_km) + w, w)
    plt.subplot(4,3,9)
    plt.hist(valid_test.ch4_mg_km, edgecolor='black', bins=10, range=[1.7,3.2])
    plt.title('CH4 Distribution')
    plt.xlabel('CH4 mg/km')
    plt.ylabel('Frequency')

    #plt.show() #test for graph to show

    filename = PdfPages("Chamber Distribution VDC1.pdf")
    filename.savefig(fig)
    filename.close()
    
def vdc2_distrubtion():
    #Open file
    df = pd.read_excel(r"P:\CORR Project\Data Trackers\RDEC 2\VDC2 Q3.xlsx.xlsm", sheet_name='Summary Sheet')
    #change formatting to be able to interogate
    df.columns = df.columns.str.strip().str.lower().str.replace(' ', '_')
    df.columns = df.columns.str.replace(' ', '_').str.replace('(', '').str.replace(')', '').str.replace('/', '_').str.replace('+', '').str.replace('#', '').str.replace(',', '')
    #remove columns that arnt needed
    df = df.drop(columns=['comments', 'barometerkpa', 'driver', 'test_name', 'date', 'mileage_miles', 'violations', 'distance_km', 'inertial_work_rating__%','temperature_°c'])
    #print(df)
    #filter criteria 1
    wltc_test = df[(df.cycle == 'WLTC Class3b TYPE 1') | (df.cycle == 'WLTC Class3b TYPE 1 Development')]
    #filter criteria 2
    valid_test = wltc_test[((wltc_test.valid_test_can_be_used_for_conformity_y_n == 'Y') & (wltc_test.test_result_valid_y_n == 'Y'))]
    #print (valid_test)

    #Figure attributes to set page up 
    logo = Image.open("P:\CORR Project\Data Analysis\Extras\MPT Logo.png")
    fig = plt.figure(figsize=[8.25,11.75])
    plt.subplots_adjust(left=0.062, bottom=0.06, right=0.98, top=0.952, wspace=0.321, hspace=0.433)
    newax = fig.add_axes([0.77, 0.78, 0.2, 0.2], anchor='NE',)
    newax.imshow(logo)
    newax.axis('off')
    
    #histograms for each compound, w is the bin width when generating the plot
    #subplot(column,row,graph No)
    # 1  2  3
    # 4  5  6
    # 7  8  9
    # 10 11 12

    #w = 2 np.arange(min(valid_test.co_km_mg), max(valid_test.co_km_mg) + w, w)
    plt.subplot(4,3,1)
    plt.hist(valid_test.co_km_mg, edgecolor='black', bins=7, range=[158,240])

    plt.title('CO Distribution')
    plt.xlabel('CO mg/km')
    plt.ylabel('Frequency')

    #w = 1 np.arange(min(valid_test.co2_km_g), max(valid_test.co2_km_g) + w, w)
    plt.subplot(4,3,2)
    plt.hist(valid_test.co2_km_g, edgecolor='black', bins=4, range=[180,193])
    plt.title('CO2 Distribution')
    plt.xlabel('CO2 g/km')
    plt.ylabel('Frequency')

    #description text
    plt.text(0.77, 0.90,'Data from ' + str(len(valid_test)) + ' valid tests', transform=plt.gcf().transFigure)
    plt.text(0.77, 0.88,'on BT70 BHV in RDEC 2', transform=plt.gcf().transFigure)

    plt.subplot(4,3,4)
    plt.hist(valid_test.thc_km_mg, edgecolor='black', bins=[10,11.2,12.4,13.6,14.8,16], range=[10,16])
    plt.title('THC Distribution')
    plt.xlabel('THC mg/km')
    plt.ylabel('Frequency')

    plt.subplot(4,3,5)
    plt.hist(valid_test.nox_km_mg, edgecolor='black', bins=7, range=[9,17])
    plt.title('NOx Distribution')
    plt.xlabel('NOx mg/km')
    plt.ylabel('Frequency')

    plt.subplot(4,3,6)
    plt.hist(valid_test.thc__nox_km_mg, edgecolor='black', bins=5, range=[20,31])
    plt.title('THC + NOx Distribution')
    plt.xlabel('THC + NOx mg/km')
    plt.ylabel('Frequency')

    #w = 10000000000 np.arange(min(valid_test.pn_km_), max(valid_test.pn_km_) + w, w)
    plt.subplot(4,3,7)
    plt.hist(valid_test.pn_km_, edgecolor='black', bins=6)
    plt.title('PN Distribution')
    plt.ticklabel_format(style='sci', axis='x', scilimits=(10,11))
    plt.xlabel('PN #/km')
    plt.ylabel('Frequency')

    #w = 0.05 np.arange(min(valid_test.pm_km_mg), max(valid_test.pm_km_mg) + w, w)
    plt.subplot(4,3,8)
    plt.hist(valid_test.pm_km_mg, edgecolor='black', bins=6, range=[0,0.5])
    plt.title('PM Distribution')
    plt.xlabel('PM mg/km')
    plt.ylabel('Frequency')

    plt.subplot(4,3,10)
    plt.hist(valid_test.n20_mg_km, edgecolor='black', bins=5, range=[0.05,0.7])
    plt.title('N2O Distribution')
    plt.xlabel('N2O mg/km')
    plt.ylabel('Frequency')

    plt.subplot(4,3,11)
    plt.hist(valid_test.l_100km, edgecolor='black', bins=6, range=[7.7,8.6])
    plt.title('Fuel Consumption Distribution')
    plt.xlabel('Fuel conusmption l100/km')
    plt.ylabel('Frequency')

    #w = 0.1 np.arange(min(valid_test.ch4_mg_km), max(valid_test.ch4_mg_km) + w, w)
    plt.subplot(4,3,9)
    plt.hist(valid_test.ch4_mg_km, edgecolor='black', bins=6, range=[1.7,3.2])
    plt.title('CH4 Distribution')
    plt.xlabel('CH4 mg/km')
    plt.ylabel('Frequency')

    #plt.show() #test for graph to show

    filename = PdfPages("Chamber Distribution VDC2.pdf")
    filename.savefig(fig)
    filename.close()
    
vdc1_distrubtion()
vdc2_distrubtion()
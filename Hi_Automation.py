import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

#Python Pandas is a Python data analysis library. It can read, filter and re-arrange small and large data sets and output them in a range of formats including Excel.

pd.set_option('display.max_rows', 500)
pd.set_option('display.max_columns', 500)
pd.set_option('display.width', 1000)

"""
Leitungswiederstand :R = l / (γ * A)[Ohm]
l=Leitungslänge(Cable length) in meter
γ=griechisches Gamma; Leitfähigkeit(Conductivity) in (m/Ohm*mm²)
A = Leitungsquerschnitt(Wire Cross Section) in mm²
"""

"""
Die Formel, die hinter dem Ohmschen Gesetz steckt, ist folgende:
    U = R * I

    U = Spannung am Verbraucher in Volt
    R = (Leitungs-)Widerstand in Ohm
    I = Strombedarf des Verbrauchers in Ampere
"""

"""
Spannungsfall berechnen;
Versorgungsspannung = X Volt
Spannung am Verbraucher = Y Volt (U=R*I)
Spannungsfall in Volt = Z Volt (Versorgungsspannung - Spannung am Verbraucher)[V]
"""
excel_file_1 = 'DataScience1.xlsx'

#Input Data by using PANDAs library to create a data frame
df1 = pd.read_excel(excel_file_1, sheet_name='First')
df2 = pd.read_excel(excel_file_1, sheet_name='Second')


print(df1)

#Pulling data from a specific column
print(df1['Name'])

#This finction is created to calculate the voltage drop on the cable
def voltage_drop(row_current, row_er):
    return row_current*row_er

#This function is created to calculate the consumer voltage
def consumer_voltage(row_supply, row_cable):
    return row_supply-row_cable

#Lamda is an anonymous function which allows you to input more than one variable
df1['Voltage Drop(V)'] = df1.apply(lambda row: voltage_drop(row['Current(A)'], row['Electrical Resistance(Ohm)']), axis = 1)
df2['Voltage Drop(V)'] = df2.apply(lambda row: voltage_drop(row['Current(A)'], row['Electrical Resistance(Ohm)']), axis = 1)

#Apply a function along an axis of the DataFrame.

#Objects passed to the function are Series objects whose index is either the DataFrame’s index (axis=0) or the DataFrame’s columns (axis=1).
#By default (result_type=None), the final return type is inferred from the return type of the applied function. Otherwise, it depends on the result_type argument.
df1['Consumer Voltage(V)'] = df1.apply(lambda row: consumer_voltage(row['Supply Voltage(V)'], row['Voltage Drop(V)']), axis = 1)
df2['Consumer Voltage(V)'] = df2.apply(lambda row: consumer_voltage(row['Supply Voltage(V)'], row['Voltage Drop(V)']), axis = 1)

df_plot = pd.DataFrame(columns=['Product_Name', 'Voltage_Drop(V)'])
df_plot['Product_Name'] = ['Product_1', 'Product_2']
df_plot['Voltage_Drop(V)'] = [float(df1['Voltage Drop(V)'].mean()), float(df2['Voltage Drop(V)'].mean())]
plt.figure(figsize=(5,10))
plt.bar(x=df_plot['Product_Name'], height=df_plot['Voltage_Drop(V)'])
plt.title('Average Voltage Drops of the Products')
plt.show()

#plt.savefig('ofke.png')
writer = pd.ExcelWriter('output1.xlsx', engine='xlsxwriter')

df1.to_excel(writer, sheet_name='First', index=False)
df2.to_excel(writer, sheet_name='Second', index=False)

#Close the Pandas Excel writer and output the Excel file
writer.save()
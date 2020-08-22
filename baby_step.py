import speech_recognition as sr
import pandas as pd
import matplotlib.pyplot as plt
r = sr.Recognizer()
with sr.Microphone() as source:
    while True:
        print("Listening")
        audio = r.listen (source)
        voice_data = r.recognize_google(audio)

        print(voice_data)w

        if 'Sales' in voice_data:
            graph = pd.read_excel('C:\\Users\\16478\\Desktop\\data_base\\sales.xlsx','Sales')
            plt.plot(graph['Months'],graph['Quantity'],graph['Return'])
            plt.title('Monthly Sales/Returns')
            plt.legend(['Quantity','Return'])
            plt.xlabel('Months')
            plt.ylabel('Quantity/Return')
            plt.show()
        elif 'no' in voice_data:
            graph = pd.read_excel('C:\\Users\\16478\\Desktop\\data_base\\sales.xlsx','Sales')
            plt.bar(graph['Months'],graph['Quantity'],graph['Return'])
            plt.title('Monthly Sales/Returns')
            plt.legend(['Quantity','Return'])
            plt.xlabel('Months')
            plt.ylabel('Quantity/Return')
            plt.show()
        elif 'yes' in voice_data:
            graph = pd.read_excel('C:\\Users\\16478\\Desktop\\data_base\\sales.xlsx','Sales')
            plt.scatter(graph['Months'],graph['Quantity'],graph['Return'])
            plt.title('Monthly Sales/Returns')
            plt.legend(['Quantity','Return'])
            plt.xlabel('Months')
            plt.ylabel('Quantity/Return')
            plt.show()

        elif 'items' in voice_data:
            graph = pd.read_excel('C:\\Users\\16478\\Desktop\\data_base\\sales.xlsx','Items')
            plt.plot(graph['Months'],graph['Table'],graph['Chair'],graph['Carpet'])
            plt.title('Monthly Products Quantities')
            #plt.legend(['Table','Chair','Carpet'])
            plt.legend((Table, Chair, Carpet), ('Table', 'Chair', 'Carpet'))
            plt.xlabel('Months')
            plt.ylabel('Products')
            plt.show()

        elif 'employee' in voice_data:
            graph = pd.read_excel('C:\\Users\\16478\\Desktop\\data_base\\employee.xlsx')
            plt.plot(graph['Name'],graph['Salary'],graph['Production'])
            plt.title('Employee Salary/Products')
            plt.legend(['Salary','Production'])
            plt.xlabel('Employee Name')
            plt.ylabel('Salary Production')
            plt.show()

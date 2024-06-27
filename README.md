Informaçãos em Português:

Script de estudo da utilização do openpyxl e selenium.

A finalizade do Script é acessar uma planliha onde existem informações de Clientes, compras feitas e de qual forma foi efetuado o pagamento. Com os dados dos clientes verificar em um site qual a situação do pagamento e então alimentar uma outra planilha.

Utilizada a biblioteca openpyxl feito o acesso as planilhas, obtidos os dados para consulta e depois adicionadas informações após análise.
Já com a bibliteca Selenium feita utilização do Browser Firefox para ter acesso ao site desejado, efetuar as consultar através de acesso aos componentes de tela com XPath.

---

Study Script for Using Openpyxl and Selenium
Purpose:

This script aims to access a spreadsheet containing customer information, purchases made, and payment methods. Using customer data, it will check the payment status on a website and populate another spreadsheet with the results.

Tools:

Openpyxl: Used to access spreadsheets, retrieve data for queries, and add information after analysis.
Selenium: Utilizes the Firefox browser to access the desired website, perform queries through access to screen components with XPath.
Detailed Description:

Data Acquisition:

Open the source spreadsheet containing customer information using openpyxl.
Extract relevant data such as customer names, purchase details, and payment methods.
Payment Status Verification:

Utilize Selenium to launch the Firefox browser and navigate to the payment status checking website.
For each customer, input their corresponding payment information from the extracted data.
Scrape the website's response to obtain the payment status for each customer.
Data Consolidation:

Open the target spreadsheet using openpyxl.
Create a structured format to store the consolidated information, including customer details, purchase information, payment methods, and verified payment statuses.
Populate the target spreadsheet with the extracted and verified data.

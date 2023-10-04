/**
 * @description This app.js file take data of Customers.json file and it as 
 * customers.xlsx file
 * @author Siddhant Prajapati
 * @version 1.0.0
 * 
 */

/**
 * @description import Customers.json file
 * @param type=String     value=path of Customers.json file
 * @return type=object    value=customers.json
 */
const customerData = require('../data/Customers.json')

/**
 * @description xlsx npm package is use export data as .xlsx file extension
 */
const XLSX = require('xlsx')

/**
 * @description path package is useful to set the path where excel file 
 * is save or store
 */
const path = require('path')

/**
 * @description ageConvert method find age of with provided dateOfBith
 * @param type=date    value=birthdate
 * @return type=number     value=age         
 */
 ageConvert=(dateOfBirth) => {
    const today = new Date()
    const birthdate = new Date(dateOfBirth)
    const age = today.getFullYear() - birthdate.getFullYear() - 
             (today.getMonth() < birthdate.getMonth() || 
             (today.getMonth() === birthdate.getMonth() && today.getDate() < birthdate.getDate()));
  return age;
 }

 /**
  * @description getCustomerUsefulData function retrive(select) only useful data from
  * customers object data
  * @params  type=object    value=customer
  * @return  type=object    value=reduced customer data
  */
getCustomerUsefulData = (customer) => {
    customer = customer.customers
    return {
        customerId : customer.customerId,
        First_Name :customer.name.first,
        Last_Name :customer.name.last,
        Email : customer.email,
        Age : ageConvert(customer.dateOfBirth)
    }
}

/**
 * @description map only required field of customers.json file
 */
const usefulCustomers =customerData.map((customer) => {
    return getCustomerUsefulData(customer)
}) 


const workSheetName = 'Customers'  //Set name of worksheet it will shown as the name of worksheet
const filePathOfExcel = './customers.xlsx'  //Set the path where excel file is store

/**
 * @description workSheetColumnName array contain name of columns of excel file, it will use 
 * to map objects field with excels column
 */
const workSheetColumnName = [
    "Customer ID",
    "First Name",
    "Last Name",
    "Email",
    "Age"
]

/**
 * 
 * @param {*} usefulCustomers = Object of useful fields customers 
 * @param {*} workSheetColumnName = array of all column name of Excel file 
 * @param {*} workSheetName = Name of excel worksheet
 * @param {*} filePath = set the file path of excel file
 * @returns 
 */
const exportUsersToExcel = (usefulCustomers, workSheetColumnName, workSheetName, filePath) => {
    
    const customers =usefulCustomers.map(customer => {
        return [
            customer.customerId, 
            customer.First_Name,
            customer.Last_Name,
            customer.Email,
            customer.Age
        ]
    })

    const workBook = XLSX.utils.book_new(); //create new workbook
    const workSheetData = [
        workSheetColumnName,
        ...customers
    ]
    const worksheet = XLSX.utils.aoa_to_sheet(workSheetData)
    XLSX.utils.book_append_sheet(workBook, worksheet, workSheetName)
    XLSX.writeFile(workBook, path.resolve(filePath))
    return true
}

exportUsersToExcel(usefulCustomers, workSheetColumnName, workSheetName, filePathOfExcel)
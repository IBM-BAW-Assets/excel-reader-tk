# IBM Business Automation Workflow Excel Reader Toolkit
#### Demonstrates reading Excel sheets using Apache POI.
Medium article [here](https://medium.com/ibm-business-automation-ap-el/ibm-business-automation-workflow-excel-reader-toolkit-fd9fe606d14d)

## Introduction
Spreadsheets are commonly used for business and are popular with users. It is not uncommon for users to request access to some data as a spreadsheet. Several years ago, a proof of technology for using Apache POI was shared in the community (by Neil Kolban). We have revisited this in a recent project and decided it was worth providing readers with an updated version using the latest versions IBM Business Automation Workflow and updated version of Apache POI. We also made several updates to the library and what is supported in terms of excel data structures. This article details this updated library how it works and data types it supports.

## Excel Reader Toolkit
This toolkit consists of external service using the java code to read excel, required java library (POI) and a CSHS example to use the service.

## ReadExcel Java Class
You can find the java code in the GitHub repository
[ReadExcel.java](https://github.com/IBM-BAW-Assets/excel-reader-tk/blob/master/java/ReadExcel.java), the source is also currently in the Jar file “ReadExcel.jar” in the toolkit so you could also just open the jar file from inside the toolkit to access the source.

The Class Read Excel converts a base 64 encoded input String to a ByteArrayInputStream and reads the excel data. Using [Apache POI](https://poi.apache.org/) classes, the class iterates the workbook converting the excel input to a JSONObject tree which is later returned as JSON string.

The reader interprets each cell is returned with the structure thus
```
{“colIndex”:0,”type”:”string”,”value”:”#”}
```
The types that are returned mapped to Excel types in the following way

- STRING string
- BOOLEAN boolean,
- BLANK null
- _NONE null
- ERROR null
- NUMERIC Date, int, or decimal. More detailed numeric not handled would need updates for more robust needs
- FORMULA formulas are evaluated, and value returned TYPES (STRING, NUMERIC, BOOLEAN, ERROR)
- Dates are a string in this format “yyyy-MM-dd’T’HH:mm:ss.SSSZ”
```
If you have need to support other cell types the source can be modified and updated to meet your needs.
```

## Example
In the toolkit, we have built an example Client Side Human Service (CSHS) under the name “Read Excel Example”. This example demonstrates an extracting data from excel spreadsheet and mapping that output to a business object in for use in the IBM Business Automation Workflow. In our first simple example we do the following:
1. Upload the excel sheet
2. Read the Excel Data
3. Show the output from reading the file

Shown in Figure 1 is three UI components which help us to complete this simple example.


![alt text](https://miro.medium.com/max/585/1*LJm5uuyKLbFSQPw4aquehg.png)

*Figure 1: Read Excel Example Components*

1. BPM File Uploader — uploads the file into BPM Document Store, this coach is included in Content Management toolkit.
2. Service Call — calls the service flow which will run the java to read the uploaded excel file.
3. Output Text — shows the output of the service call.

In the BPM File Uploader UI component in the “On Upload complete” event, the code here strips the BPM Document ID from the event message this is then passed into the service call and used to pull the document from the BPM document store in subsequent steps. This example script is shown in Figure 2.

![alt text](https://miro.medium.com/max/700/1*RQ-bXtwx5SGeiiFlC9Gpyw.png)

*Figure 2: Getting BPM Document ID*

In the invoked service call maps to the service flow shown in Figure 3.

![alt text](https://miro.medium.com/max/700/1*-AUuK66Xgx4Ku081p8i4vw.png)

*Figure 3: Read Excel Service Call*

1. Content Integration Task — get document content from BPM Document Store, the BPM Document ID will be used as the input.
2. Service Task — call the external service that using the ReadExcel java code.
3. Error Handler — output error message if occurred

In component 1 of the service flow, we used the document ID that send from the CSHS as shown in Figure 2 as an input. The purpose of this component is to get the base64 data format of the excel file and passed it to the java service. To serve this purpose we are using Content Integration Task and choose Get document content as the operation. Figure 4 shows the implementation and data mapping of the component. The output of this operation should be bound using ECMContentStream BO which will be in Content Management toolkit.

![alt text](https://miro.medium.com/max/700/1*1M_zZylI7lwB_YwR6aOuPw.png)

*Figure 4: Component 1 Implementation and Data Mapping*

As for component 2, the input is not exactly the same data type as component 1 output where it used ECMContentStream BO, instead it only needs the base64 data format of the document, we can find this in the BO in the content parameter. Figure 5 shows the input mapping of component 2.

![alt text](https://miro.medium.com/max/700/1*emM34bP5Yz5HM6-0QCv3Rw.png)

*Figure 5: Component 2 Input Mapping*

With the CSHS and the Service flow established we shall now execute a simple example. Figure 5 shows the simple spreadsheet that will be used.
You can also find the example excel in our [GitHub repository](https://github.com/IBM-BAW-Assets/excel-reader-tk/blob/master/files/Personal%20Data.xlsx)

![alt text](https://miro.medium.com/max/960/1*iBFgIQWbo_s5WdsmHu0rxQ.png)

*Figure 6: Excel Spreadsheet Content*

Figure 6 shows the raw output of a stingified JSON object, this will be mapped to your business object. To further illustrate the returned output from the file we have placed a human readable layout for this example in Figure 7 and in [GitHub](https://github.com/IBM-BAW-Assets/excel-reader-tk/blob/master/files/output.json)

![alt text](https://miro.medium.com/max/700/1*aIAI_1U5GvDj5piZzsSDLg.png)

*Figure 7: JSON Output*

The output starts with the list of sheets in the excel and has a tree structure that navigates to the cell level for each row.
The java code read all the data from the spreadsheet, it also included the table header row which is “Name”, “Age”, “Phone Number” and “Email”.

![alt text](https://miro.medium.com/max/622/1*G5hJk-rxyVVO2Kb0L_tBNA.png)

*Figure 8: ReadExcel JSON Object Structure*

When we already know the structure of the JSON, we can now take the data we want and use it in our Business Object.
In Read Excel CSHS, there is a simple private variable using a business object (BO) called “PersonalData”. We will convert the JSON output from the excel reader service to this BO and later can be used to other component in BAW.

![alt text](https://miro.medium.com/max/551/1*QxriYZ0XGlgoSwHwdkbEaQ.png)

*Figure 9: PersonalData business object*


In Figure 8, we can take a look at the client script used to convert the JSON to BO. As we have seen the JSON structure previously, we can loop based on the number of the rows minus the table header. From there we can set each variable of the BO following the JSON.

![alt text](https://miro.medium.com/max/673/1*wpi-MktUF2dqeuaEOkYW3w.png)

*Figure 10: Convert JSON Script*

Onto the next coach view, we show how the newly mapped business object is bound to the tables personData list. In Figure 9, as final result, we can see that the records from excel spreadsheet can consumed and shown on screen with tables. And that concluded the simple example of reading and consuming excel spreadsheet data.

![alt text](https://miro.medium.com/max/431/1*-43aBCvmDy5eeWBZS5NwFg.png)

*Figure 11: Table using Binded data*

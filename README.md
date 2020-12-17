Note: MongoDB should be installed on this system
MongoDB Installation: https://fastdl.mongodb.org/windows/mongodb-windows-x86_64-4.4.0-signed.msi
Pymongo and Openpyxl modules should also be installed on your ide or system.

Using a data set, A Time table for students and invigilators and a list of students in each room is created. It is done using MongoDB advanced operators and Python3
Input:
A .xlsx file  with the name "Exam Allootment" should be present within the same directory as the python file with the following structure
The first sheet should contain the folllowing data starting from the 4th row
1st column - Student Name
2nd column - Student ID
3rd column - Branch/Department
4th column - Semester/Year

The second sheet should contain the following data starting from the 5th row
2nd column - Faculty/Invigilator name
5th column - Experience/Designation(in my code New faculties can have 5 duties during one exam cycle,Assisstant professor can have two 3 duties and Associate professors can have 2 duties. This can easily be changed in the lines 150-155 of the code)

The third sheets should contain the following data from the 5th row
1st column - Room number or name
2nd column - Capacity 
4th column - Room Position(Floor/ Block etc)

The remaining sheets are of the same structure and should contain the following data
1st row,1st column - Department
1st row,2nd column - Year/Semester
From 3rd row onwards
1st column- Course code
2nd column - Course Name

The project has two allotment modes
Internal exam mode - 4 days of 2 exams maximum
End semester exam mode - 7 days of one exam each

An .xlsx empty document named "Output" should be present within the same directory of the python code

Output:
The  Output document mentioned in line 34 of this document is in the following structure
first sheet is empty
n sheets of the following structure:
1st row,1st column - Room number/name
1st row, 2nd column - Room position

n+1th sheet containing the following data
1st column - Faculty/Invigilator name
2nd column - Date of Duty
3rd column - Time of Duty
4th column - Student branch that they have to invigilate
5th column - Exam subject
6th column - Exam code

m sheets for each date of exam with the following data
1st column - Faculty/Invigilator name
2nd column - Date of Duty
3rd column - Time of Duty
4th column - Student branch that they have to invigilate
5th column - Exam subject





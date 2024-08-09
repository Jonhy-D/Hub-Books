# Hub Reading

## Description

This project is a reading hub making an application for Windows with Visual Basic 6, its main function is to load the list of available books so the user can put which ones he has read and which ones he did not like.

## Technical requirements

This project was generated with Microsoft Visual Basic 6.

## How to install the project

### First Step: 

You need to have Visual Basic 6 installed for this project to work.

### Second Step: 

When you have installed Visual Basic 6 you will have to open the project, go to the modules section and in the module called **ConnectionDB** and in the ConecctionString put:

```
"Provider=SQLOLEDB;Data Source="NAME OF SERVER SQL";Initial Catalog="NAME OF DATA BASE";User ID="USER SERVER";Password="PASSWORD OF USER SERVER";"
```

### Final Step

Run the proyect.

### Remember need a database

The database is in the file called **HubOfReading.dacpac**

# Sprint 4

### Goals
Make a reading hub program for Windows with the following:

- Mega Books Catalog.
- Books you've already read.
- Books you want to read.
- Books you didn't like.

All with SQL Server and Visual Basic 6.

## Entity-Relationship Diagram

This is the entity relationship diagram of the hub entertainment mega project database according to my point of view.

![Entity-Relationship Diagram Image](/Images/Entity-Relationship-HUB-R.webp)

## Images of Aplication Windows - Hub Of Reading

- This image shows the user interface.

![Profile Aplication](/Images/Profile.webp)

- This image shows the interface of the book library.
 
![Library Aplication](/Images/Library.webp)

- This image shows the book details interface.

![Details of book](/Images/Detailsb.webp)

## Process

- First I created the first main tables of the entire database.
- After having created those tables to be able to store those items, then I had to see the functions of my application to be able to make the other tables that are the user's favorites and Since in MySQL there are relationship tables.
- I had to think about how the user relates to those tables and thus create the other tables of favorite books and those read in addition to those you don't like.
- After creating the database, we now begin to create the interface, which was somewhat tedious because it is not something we currently use, so it is difficult to understand what each thing is used for.
- The last thing was to bring all the buttons we created to life and give them functionality, but without the Challenger we wouldn't have been able to move forward, really.

## Issues

- I think this sprint has now surpassed the bugs I was getting from the other project.
- Because it is simply very difficult to find how to fix it because the errors are not at all legible or understandable.
- Because I made a mistake and spent almost a whole day reading articles from people who posted years ago on websites that are also super old and so it was very difficult to find a solution that would truly solve my problems.
- I think my problem was more that I didn't like the interface of how to work with Visual Basic because after fixing an error you can continue with the next one and more or less you get an idea of ​​what's going on.

## Table Sprint Review

| **What was done well?** | **What can I do differently?** | **What didn't go well?** |
------------------|----------------------------|-----------------------
| I think that in this sprint it was easier to create things because the sessions with the challenger were extensive so he had the opportunity to teach us several things that could be useful, because I was able to carry out the sprint in my own way as I completed it. | I don't know what to answer in this sprint because I can manage my times very well, the only thing that was a problem was error handling and there is not a wide range of resources. | For me, working with Visual Basic 6 is what is not good because it is too outdated and there are not many options to extract information. |
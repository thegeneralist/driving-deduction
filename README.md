# driving-deduction
A JS project to calculate driving for expenses and list out your meetings for sales purposes

This project was almost entirely built using Cursor. 

Currently it requires that you create your own google API keys in a developer mode (ie, it isn't built for production–just for perosnal use though per the license, commercial use is fine)

It serves 2 purposes:
1. calculate mileage driven (there and back) for tax deduction/expense reimbursement purposes. 
2. creates a CSV with every meeting and every meeting attendee in a given time period, so you can do whatever salesy thing you want to do with that information (maybe send a holiday e-card?)

It takes 3 inputs:
start date
end date
how many one-way miles away to ignore a drive (will explain below)

It outputs 3 things:
- A JSON file, which for most people probably won't be that helpful
- A CSV ending in Mileage that includes the meetings and mileage, as well as total mileage, for the period. 
- A CSV ending in Meetings listing all the meetings and meeting attendees for the period. 

Happy selling!
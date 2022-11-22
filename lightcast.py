import pandas as pd
import openpyxl as opyxl
import matplotlib.pyplot as plt

# The four databases to be used, read and converted to Data Frames
# IMPORTANT: File paths will likely need to be readjusted if using the program on a different computer.
employmentDF = pd.read_excel(r"C:\Users\Calvin\OneDrive\playground\playground\Occ_Employment.xlsx")
genderDistributionDF = pd.read_excel(r"C:\Users\Calvin\OneDrive\playground\playground\Occ_Gender_Distribution.xlsx")
skillsDF = pd.read_excel(r"C:\Users\Calvin\OneDrive\playground\playground\Skills.xlsx")
workActivitiesDF = pd.read_excel(r"C:\Users\Calvin\OneDrive\playground\playground\Work_Activities.xlsx")


# Key words in skills and work activites that determine if they're analytical attributes, chosen from my definition of analytical
analyticalWords = ["Analysis", "Analyze", "Analyzing", "Observation", "Observe", "Observing", "Evaluation", 'Evaluate', 
"Evaluating", "Judgment", "Judge", "Judging", "Processing Information", "Information Processing", "Process Information",
"Conclusion", "Conclude", "Concluding", "Identify Pattern", "Identifying Pattern", "Logic", "Logical", "Examine", "Examining", "Examination"]

# IsAnalyticalSkill: String -> 1 for True, 0 for False
def IsAnalytical(string):
    for word in analyticalWords:
        if word in string:
            return 1
    return 0

# Use the map function to create a new column in the skills and work activities Data Frames marking if 
# each attribute is considered analytical based on the element name
skillsDF["Indicator"] = skillsDF["Element Name"].map(IsAnalytical)
workActivitiesDF["Indicator"] = workActivitiesDF["Element Name"].map(IsAnalytical)

# Calculating the number of points that should be added 
# to the job's rating by multiplying its indicator by its data value
skillsDF["Points"] = skillsDF["Indicator"] * skillsDF["Data Value"] * skillsDF["Data Value"]
workActivitiesDF["Points"] = workActivitiesDF["Indicator"] * workActivitiesDF["Data Value"] * workActivitiesDF["Data Value"]

# Sums is a dictionary that will store the 6-digit job IDs as keys and keep track of their ratings as ppoints are added as the values.
sums = {}

# Shorten will be used to shorten 8-digit job IDs down to the 6-digit versions
# Shorten: String -> String
def Shorten(jobID):
    newJobID = jobID[:7]
    return newJobID


# Creating the 6-digit job ID entries in the dictionary and starting them out with 0 points
for index in range(genderDistributionDF.shape[0]):
    sums[genderDistributionDF.loc[index, "2018 SOC Code"]] = 0

# Iteratet through the skills Data Frame adding the earned points
for index in range(skillsDF.shape[0]):
    sums[Shorten(skillsDF.loc[index, "O*NET-SOC Code"])] += skillsDF.loc[index, "Points"]

# Iterate through the work activities Data Frame adding the earned points
for index in range(workActivitiesDF.shape[0]):
    sums[Shorten(workActivitiesDF.loc[index, "O*NET-SOC Code"])] += workActivitiesDF.loc[index, "Points"]


# Turn the sums dictionary into a column of the genderDistribution Data Frame
genderDistributionDF['Analytical Rating']= genderDistributionDF['2018 SOC Code'].map(sums)

# Sort gender distribution DF by the rating with highest scores at the top
genderDistributionDF.sort_values(by=["Analytical Rating"], inplace=True, ascending=False)
print(genderDistributionDF.to_string())

# Create a scatter plot with the analytical rating on x axis, proportion of Female on the y axis
plt.scatter(genderDistributionDF['Analytical Rating'], genderDistributionDF['Proportion Female'])
plt.show()
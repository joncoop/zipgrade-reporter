import json
import statistics
import docx

csv_data = "apcsa2014practiceexamExport.csv"

with open(csv_data, 'r') as f:
    contents = f.read().splitlines()

fields = contents[0].split(',')

records = []

for line in contents[1:]:
    r = {}
    values = line.split(',')
    
    for field, value in zip(fields, values):
        r[field] = value

    records.append(r)

pretty = json.dumps(records, indent=4, sort_keys=False)
print(pretty)

# now put it in word
document = docx.Document()
document.add_heading('Score Report', 0)

for r in records:
    title = r['QuizName']
    section = r['QuizClass']
    last = r['LastName']
    first = r['FirstName']
    zip_id = r['ZipGradeID']
    earned = r['EarnedPts']
    possible = r['PossiblePts']
    percent = r['PercentCorrect']

    if 'Key100' in r:
        sheet_size = 100
    elif 'Key50' in r:
        sheet_size = 50
    else:
        sheet_size = 25

    result = ""
    wrong = ""
    num_wrong = 0
    for i in range(1, sheet_size + 1):
        correct = r['Key' + str(i)]

        if len(correct) > 0:
            answer = r['Stu' + str(i)]

            if len(answer) == 0:
                answer = "-"
                
            num_wrong += 1
            wrong += "\t" + str(i) + ". " + answer

            if answer != correct:
                wrong += " (" + correct + ")"

                if i < 10:
                    wrong += "\t"
            else:
                wrong += "\t"
            
            if num_wrong % 5 == 0:
                wrong += "\n"
        
    result += "Name: " + last + ", " + first + " (ID=" + zip_id + ")\n"
    result += "Test: " + title + "\n"
    result += "Class: " + section + "\n"
    result += "Raw: " + earned + "/" + possible + "\n"
    result += "Percent: " +  percent + "\n"
    result += "Incorrect responses: (your answer, correct)" + "\n"
    result += wrong
    result += "\n\n"

    paragraph = document.add_paragraph(result)
    paragraph.paragraph_format.keep_together = True

document.save('report.docx')

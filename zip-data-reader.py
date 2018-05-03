import json
import statistics
import docx

def get_import_path():
    '''
    Gets path to ZipGrade data file.
    
    Right now it's hard coded, but this could be obtained from GUI later.
    '''

    return "apcsa2014practiceexamExport.csv"

def get_export_path():
    '''
    Gets path to save report.
    
    Right now it's hard coded as an empty string so the report will be
    saved in the same file as the program is run.
    '''

    return ""

def get_export_filename():
    '''
    Gets path to save report.
    
    Right now it's hard coded, but this could be obtained from GUI later.
    '''

    return "sample_report.docx"

def csv_to_json(path):
    '''
    Reads ZipGrade csv export file and stores as JSON data
    '''
    
    with open(path, 'r') as f:
        contents = f.read().splitlines()

    fields = contents[0].split(',')

    records = []

    for line in contents[1:]:
        r = {}
        values = line.split(',')
        
        for field, value in zip(fields, values):
            r[field] = value

        records.append(r)

    #pretty = json.dumps(records, indent=4, sort_keys=False)
    #print(pretty)

    sort_by = lambda k: k['LastName'] + " " + k['FirstName']
    records = sorted(records, key=sort_by , reverse=False)
    
    return records

def get_raw_scores(records):
    scores = []

    for r in records:
        scores.append(int(r['EarnedPts']))
        
    return scores
        
def get_percentages(records):
    scores = []

    for r in records:
        s = float(r['PercentCorrect'])
        scores.append(s)
        
    return scores
        
def make_cover_page(records, document):
    '''
    Creates cover page with summary statistics.
    '''

    r = records[0]
    title = r['QuizName']
    date_created = r['QuizCreated']
    date_exported = r['DataExported']
    num_scores = len(records)
    possible_points = r['PossiblePts']

    scores = get_raw_scores(records)
    percentages = get_percentages(records)

    mean_raw = statistics.mean(scores)
    max_raw = max(scores)
    min_raw = min(scores)

    mean_percent = statistics.mean(percentages)
    max_percent = max(percentages)
    min_percent = min(percentages)
    
    document.add_heading('ZipGrade Score Report', 0)
    document.add_heading(title, 1)

    p = document.add_paragraph()
    p.add_run("Date Created: ").bold = True
    p.add_run(date_created + "\n")
    p.add_run("Date Exported: ").bold = True
    p.add_run(date_exported)
    
    document.add_heading('Summary Statistics', 1)

    p = document.add_paragraph()
    p.add_run("Number of Scores: ").bold = True
    p.add_run(str(num_scores) + "\n")
    p.add_run("Points Possible: ").bold = True
    p.add_run(str(possible_points))

    p = document.add_paragraph()
    p.add_run("Mean (raw): ").bold = True
    p.add_run(str(max_raw) + "\n")
    p.add_run("Max (raw): ").bold = True
    p.add_run(str(min_raw) + "\n")
    p.add_run("Min (raw): ").bold = True
    p.add_run(str(min_raw))

    p = document.add_paragraph()
    p.add_run("Mean (%): ").bold = True
    p.add_run(str(mean_percent) + "\n")
    p.add_run("Max(%): ").bold = True
    p.add_run(str(max_percent) + "\n")
    p.add_run("Min (%): ").bold = True
    p.add_run(str(min_percent))
   
    document.add_page_break()

def make_score_summary(records, document):
    '''
    Creates summary of indidual student scores.
    '''

    document.add_heading('Individual Scores', 1)
    table = document.add_table(rows=1, cols=4)
    table.style = 'Medium Shading 1'
    
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Name'
    hdr_cells[1].text = 'Raw'
    hdr_cells[2].text = 'Possible'
    hdr_cells[3].text = 'Percent'

    for r in records:
        row_cells = table.add_row().cells
        row_cells[0].text = r['LastName'] + ", " + r['FirstName']
        row_cells[1].text = r['EarnedPts']
        row_cells[2].text = r['PossiblePts']
        row_cells[3].text = r['PercentCorrect'] + "%"
        
    document.add_page_break()

def make_individual_report(record):
    '''
    Creates report for a student
    '''
    
    title = record['QuizName']
    section = record['QuizClass']
    last = record['LastName']
    first = record['FirstName']
    zip_id = record['ZipGradeID']
    earned = record['EarnedPts']
    possible = record['PossiblePts']
    percent = record['PercentCorrect']

    if 'Key100' in record:
        sheet_size = 100
    elif 'Key50' in record:
        sheet_size = 50
    else:
        sheet_size = 25

    result = ""
    wrong = ""
    num_wrong = 0
    for i in range(1, sheet_size + 1):
        correct = record['Key' + str(i)]

        if len(correct) > 0:
            answer = record['Stu' + str(i)]

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
        
    result += "Name: " + last + ", " + first + "\n"
    result += "ID: " + zip_id + "\n"
    result += "Test: " + title + "\n"
    result += "Class: " + section + "\n"
    result += "Raw: " + earned + "/" + possible + "\n"
    result += "Percent: " +  percent + "%\n"
    result += "Response Summary: (Correct)\n"
    result += wrong
    result += "\n\n"
    
    return result
    
def generate_word_doc(records, save_path):
    '''
    Creates Word Doc with individual score reports.
    '''

    document = docx.Document()
    make_cover_page(records, document)
    make_score_summary(records, document)
    
    for r in records:
        student_report = make_individual_report(r)
        paragraph = document.add_paragraph(student_report)
        paragraph.paragraph_format.keep_together = True

    document.save(save_path)

# go!
if __name__ == "__main__":
    import_path = get_import_path()
    export_path = get_export_path()
    export_filename = get_export_filename()
    
    data = csv_to_json(import_path)
    save_path = export_path + export_filename
    generate_word_doc(data, save_path)
